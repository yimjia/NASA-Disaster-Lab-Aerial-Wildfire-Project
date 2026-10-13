#!/usr/bin/env python3
"""
NASA Aerial Wildfire research RAG.

Local, offline, no API keys. Indexes every PDF in the References folder,
searches with hybrid BM25 + embeddings, and prints cited passages that
Claude Code (or you) can answer from.

Usage:
  rag index                 build / refresh the index (PDF and DOCX; only changed files are re-read)
  rag search "query" [-k N] [--doc SUBSTR] [--full]
  rag docs                  list indexed documents with page counts
  rag page "DOC SUBSTR" N   print one page of one document verbatim
  rag ask "question" [-k N] answer with Claude Code CLI (claude -p) using retrieved passages
"""
import argparse, hashlib, json, os, re, subprocess, sys
for _k, _v in {"HF_HUB_DISABLE_PROGRESS_BARS": "1", "TRANSFORMERS_VERBOSITY": "error",
               "TOKENIZERS_PARALLELISM": "false", "HF_HUB_DISABLE_TELEMETRY": "1"}.items():
    os.environ.setdefault(_k, _v)
from pathlib import Path
if (Path.home() / ".cache/huggingface/hub/models--BAAI--bge-small-en-v1.5").exists():
    os.environ.setdefault("HF_HUB_OFFLINE", "1")   # model cached, skip hub checks and warnings

REFS = Path(os.environ.get(
    "NASA_RAG_REFS",
    "/Users/vickylopez/Desktop/NASA_Aerial_Wildfire_Project/References"))
HOME = Path(__file__).resolve().parent
IDX = HOME / "index"
IDX.mkdir(exist_ok=True)
CHUNKS = IDX / "chunks.jsonl"
MANIFEST = IDX / "manifest.json"
EMB = IDX / "embeddings.npy"
MODEL = "BAAI/bge-small-en-v1.5"      # 33M params, strong for retrieval, runs fast on M4
CHUNK_CHARS, OVERLAP = 1400, 250
CONOPS_BOOST = 1.35                   # ConOps is the backbone of the research; rank it up


# ------------------------------------------------------------------ helpers
def sha1(p: Path) -> str:
    h = hashlib.sha1()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def short_name(fname: str) -> str:
    """'Brewer & Clements 2020 Meteorological profiling...pdf' -> 'Brewer & Clements 2020'"""
    stem = Path(fname).stem
    m = re.match(r"^(.*?\b(?:19|20)\d{2}\b)", stem)
    if m and len(m.group(1)) < 60:
        return m.group(1).strip()
    if "conops" in stem.lower():
        return "NASA ACERO ConOps 2026"
    return stem[:50]


def is_conops(fname: str) -> bool:
    s = fname.lower()
    return "conops" in s or "concept of operations" in s


def pdf_pages(p: Path):
    """Yield (page_no, text) using pdftotext with form feeds between pages."""
    out = subprocess.run(["pdftotext", "-layout", str(p), "-"],
                         capture_output=True, text=True, errors="replace").stdout
    for i, page in enumerate(out.split("\f"), start=1):
        t = re.sub(r"[ \t]+", " ", page)
        t = re.sub(r"\n{3,}", "\n\n", t).strip()
        if t:
            yield i, t


def docx_pages(p: Path, per_page=25):
    """Word file: paragraphs grouped into pseudo-pages of `per_page` paragraphs."""
    import zipfile
    xml = zipfile.ZipFile(p).read("word/document.xml").decode("utf8", "replace")
    paras = []
    for blk in re.findall(r"<w:p[ >].*?</w:p>", xml, flags=re.S):
        t = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", blk, flags=re.S))
        t = t.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").strip()
        if t:
            paras.append(t)
    for i in range(0, len(paras), per_page):
        yield i // per_page + 1, "\n".join(paras[i:i + per_page])


def extract_pages(p: Path):
    return docx_pages(p) if p.suffix.lower() == ".docx" else pdf_pages(p)


def chunk_pages(pages):
    """Sliding window over the concatenated page text; each chunk remembers its page span."""
    buf, spans = "", []
    for pno, text in pages:
        start = len(buf)
        buf += text + "\n\n"
        spans.append((start, len(buf), pno))

    def page_of(pos):
        for s, e, p in spans:
            if s <= pos < e:
                return p
        return spans[-1][2] if spans else 1

    pos = 0
    while pos < len(buf):
        end = min(len(buf), pos + CHUNK_CHARS)
        if end < len(buf):                      # break on a sentence/paragraph boundary
            cut = max(buf.rfind("\n\n", pos, end), buf.rfind(". ", pos, end))
            if cut > pos + CHUNK_CHARS // 2:
                end = cut + 1
        text = buf[pos:end].strip()
        if len(text) > 80:
            yield page_of(pos), page_of(end - 1), text
        if end >= len(buf):
            break
        pos = max(end - OVERLAP, pos + 1)


def load_chunks():
    if not CHUNKS.exists():
        sys.exit("No index yet. Run:  rag index")
    with open(CHUNKS) as f:
        return [json.loads(l) for l in f]


def embedder():
    from sentence_transformers import SentenceTransformer
    return SentenceTransformer(MODEL)


# ------------------------------------------------------------------ index
def cmd_index(args):
    import numpy as np
    pdfs = sorted(p for p in REFS.rglob("*") if p.suffix.lower() in (".pdf", ".docx") and not p.name.startswith("."))
    old = json.load(open(MANIFEST)) if MANIFEST.exists() else {}
    old_chunks = load_chunks() if CHUNKS.exists() and not args.rebuild else []
    old_emb = np.load(EMB) if EMB.exists() and old_chunks else None
    by_hash = {}
    for i, c in enumerate(old_chunks):
        by_hash.setdefault(c["hash"], []).append(i)

    new_manifest, chunks, emb_rows, todo, seen = {}, [], [], [], set()
    for p in pdfs:
        h = sha1(p)
        if h in seen:
            print(f"  skip duplicate: {p.name}")
            continue
        seen.add(h)
        new_manifest[p.name] = h
        if h in by_hash and old_emb is not None:
            for i in by_hash[h]:
                c = dict(old_chunks[i]); c["file"] = p.name; c["doc"] = short_name(p.name)
                chunks.append(c); emb_rows.append(old_emb[i])
        else:
            todo.append((p, h))

    if todo:
        print(f"Extracting {len(todo)} PDF(s)...")
        fresh = []
        for p, h in todo:
            pages = list(extract_pages(p))
            n = 0
            for p1, p2, text in chunk_pages(pages):
                fresh.append({"file": p.name, "doc": short_name(p.name), "hash": h,
                              "conops": is_conops(p.name), "p1": p1, "p2": p2, "text": text})
                n += 1
            print(f"  {n:4d} chunks  {len(pages):3d} pages  {p.name[:70]}")
            if n == 0:
                print(f"  WARNING: no text in {p.name} (scanned? run OCR)")
        print(f"Embedding {len(fresh)} chunks with {MODEL} (first run downloads ~130MB)...")
        model = embedder()
        vecs = model.encode([c["text"] for c in fresh], batch_size=32,
                            normalize_embeddings=True, show_progress_bar=True)
        chunks.extend(fresh); emb_rows.extend(list(vecs))

    with open(CHUNKS, "w") as f:
        for c in chunks:
            f.write(json.dumps(c) + "\n")
    np.save(EMB, np.array(emb_rows, dtype="float32"))
    json.dump(new_manifest, open(MANIFEST, "w"), indent=1)
    print(f"Index ready: {len(new_manifest)} documents, {len(chunks)} chunks -> {IDX}")


# ------------------------------------------------------------------ search
def tokenize(s):
    return re.findall(r"[a-z0-9]+", s.lower())


def retrieve(query, k=8, doc_filter=None, conops_only=False):
    import numpy as np
    from rank_bm25 import BM25Okapi
    chunks = load_chunks()
    emb = np.load(EMB)
    idx = list(range(len(chunks)))
    if doc_filter:
        idx = [i for i in idx if doc_filter.lower() in chunks[i]["file"].lower()]
    if conops_only:
        idx = [i for i in idx if chunks[i]["conops"]]
    if not idx:
        return []
    bm = BM25Okapi([tokenize(chunks[i]["text"]) for i in idx])
    b = bm.get_scores(tokenize(query))
    b = b / (b.max() or 1)
    q = embedder().encode([query], normalize_embeddings=True)[0]
    d = emb[idx] @ q
    d = (d - d.min()) / ((d.max() - d.min()) or 1)
    score = 0.45 * b + 0.55 * d
    score = score * np.array([CONOPS_BOOST if chunks[i]["conops"] else 1.0 for i in idx])
    order = np.argsort(-score)[:k]
    return [(float(score[j]), chunks[idx[j]]) for j in order]


def cite(c):
    pg = f"p.{c['p1']}" if c["p1"] == c["p2"] else f"pp.{c['p1']}-{c['p2']}"
    return f"[{c['doc']}, {pg}]"


def cmd_search(args):
    hits = retrieve(args.query, args.k, args.doc, args.conops)
    if not hits:
        print("No results."); return
    for score, c in hits:
        print(f"\n### {cite(c)}  score={score:.2f}\nfile: {c['file']}")
        print(c["text"] if args.full else c["text"][:700] + (" ..." if len(c["text"]) > 700 else ""))


def cmd_docs(args):
    chunks = load_chunks()
    seen = {}
    for c in chunks:
        s = seen.setdefault(c["file"], {"doc": c["doc"], "chunks": 0, "pages": 0, "conops": c["conops"]})
        s["chunks"] += 1; s["pages"] = max(s["pages"], c["p2"])
    for f, s in sorted(seen.items()):
        flag = "  <== ConOps" if s["conops"] else ""
        print(f"{s['pages']:4d} pp {s['chunks']:4d} ch  {s['doc']:<32} {f[:80]}{flag}")


def cmd_page(args):
    matches = [p for p in REFS.rglob("*.pdf") if args.doc.lower() in p.name.lower()]  # page view is PDF only
    if not matches:
        sys.exit(f"No PDF matching '{args.doc}'")
    p = matches[0]
    out = subprocess.run(["pdftotext", "-layout", "-f", str(args.page), "-l", str(args.page), str(p), "-"],
                         capture_output=True, text=True).stdout
    print(f"=== {p.name}  page {args.page}\n{out}")


def cmd_ask(args):
    hits = retrieve(args.question, args.k, args.doc, args.conops)
    ctx = "\n\n".join(f"{cite(c)}\n{c['text']}" for _, c in hits)
    prompt = (
        "You are a research assistant for a NASA aerial wildfire / ACERO airspace management study. "
        "Answer the question using ONLY the passages below. Cite every claim with the bracketed "
        "citation shown before its passage. If the passages do not answer it, say so. "
        "Do not use em dashes.\n\n"
        f"PASSAGES:\n{ctx}\n\nQUESTION: {args.question}")
    r = subprocess.run(["claude", "-p", prompt], text=True)
    if r.returncode != 0:
        print("\n(claude CLI not available; here are the passages)\n" + ctx)


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(prog="rag", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    a = sp.add_parser("index"); a.add_argument("--rebuild", action="store_true"); a.set_defaults(fn=cmd_index)
    for name, fn in (("search", cmd_search), ("ask", cmd_ask)):
        a = sp.add_parser(name)
        a.add_argument("query" if name == "search" else "question")
        a.add_argument("-k", type=int, default=8 if name == "search" else 12)
        a.add_argument("--doc", help="only search files whose name contains this")
        a.add_argument("--conops", action="store_true", help="only the ConOps")
        a.add_argument("--full", action="store_true", help="print whole chunks")
        a.set_defaults(fn=fn)
    a = sp.add_parser("docs"); a.set_defaults(fn=cmd_docs)
    a = sp.add_parser("page"); a.add_argument("doc"); a.add_argument("page", type=int); a.set_defaults(fn=cmd_page)
    args = ap.parse_args()
    args.fn(args)


if __name__ == "__main__":
    main()
