# nasa_rag

Local RAG over ~/Desktop/NASA_Aerial_Wildfire_Project/References (43 PDFs).
Hybrid BM25 + bge-small embeddings, ConOps boosted. No API keys.

    nasarag index              # refresh (incremental, skips duplicates)
    nasarag search "query"     # cited passages
    nasarag ask "question"     # standalone answer through the claude CLI

In Claude Code: `/nasa-rag <question>` or `/nasa-rag build a sheet of ...`.
Index lives in ./index (gitignored, rebuildable).
