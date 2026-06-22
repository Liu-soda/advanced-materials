# Advanced Materials

**AI-driven materials science experiment platform.**

Upload research papers → AI reads them → generates experiment reports with detailed procedures, equipment lists, and safety protocols.

Built on **Paper-Orchestrator v3.0** multi-agent architecture at Huazhong University of Science & Technology.

## Features

- 📄 Upload PDF papers (Chinese/English) and get experiment reports
- 🤖 DeepSeek v4-pro powered analysis and synthesis
- 📝 Export to Word + PDF
- 💬 Interactive refinement via chat
- 🔍 SEM/XRD image analysis
- 📚 Knowledge base for permanent paper storage
- 🎨 Dark sci-fi web interface

## Quick Start

```bash
# Install dependencies
pip install flask pymupdf fpdf2 python-docx openai

# Start the server
python app_web.py

# Open in browser
# http://localhost:5000
# Default key: siclab2026
```

## Agent Architecture

```
Paper Input → Orchestrator → [paper-reader ×4 parallel] → Synthesizer → Quality Gate
```

## OpenCode Skill

Also available as an OpenCode Skill:

```bash
opencode skill install advanced-materials
```

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | HTML/CSS/JS (vanilla, no framework) |
| Backend | Flask 3.x |
| AI | DeepSeek v4-pro (OpenAI-compatible) |
| PDF Engine | PyMuPDF (fitz) — ~87ms/page |
| Documents | python-docx + fpdf2 |
| Auth | Session-based with configurable key |

## License

MIT © Advanced Materials Lab @ HUST
