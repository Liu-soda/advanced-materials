---
name: advanced-materials
description: >-
  AI-driven materials science experiment platform. Upload research papers
  (Chinese/English PDFs), and the Paper-Orchestrator multi-agent system
  reads, analyzes, and generates experiment reports with specific procedures,
  equipment lists, and safety protocols. Supports any material system.
license: MIT
compatibility: opencode
metadata:
  version: "3.0"
  author: "Advanced Materials Lab @ HUST"
  architecture: "Paper-Orchestrator v3.0"
  requires: "deepseek-v4-pro API key"
---

## Advanced Materials — AI Experiment Report Generator

An AI-powered materials science platform that reads academic papers and
generates customized experiment reports. Built on the Paper-Orchestrator
multi-agent architecture.

### What it does

- Upload any materials science PDFs (Chinese or English)
- AI reads ALL papers (PyMuPDF engine, ~87ms per paper)
- Generates experiment reports with: equipment lists, step-by-step procedures,
  parameter tables, safety protocols, and expected performance data
- DeepSeek v4-pro powered synthesis and refinement
- Interactive refinement: chat with AI to modify the report
- Export to Word and PDF
- Knowledge base for permanent paper storage
- SEM/XRD image analysis via vision AI

### When to use

Use this skill when:
- Analyzing literature for experiment planning
- Generating detailed experiment protocols from papers
- Building a lab knowledge base
- Need AI-assisted literature review with practical output

### Installation

1. Copy the `agents/` and `prompts/` to your `.opencode/` directory
2. Set your DeepSeek API key in `~/.local/share/opencode/auth.json`
3. Restart OpenCode or reload

### Web App (optional)

For the graphical interface:
```bash
cd skills/advanced-materials/webapp
pip install flask pymupdf fpdf2 python-docx openai
python app_web.py
# Open http://localhost:5000
# Default key: siclab2026
```

### Agent Architecture

```
Paper Input → Orchestrator → [reader-01..04 parallel] → Synthesizer → Quality Gate
                                                                ↓
                                                          Refinement Agent
                                                                ↓
                                                          Word/PDF Output
```

### Commands

- `/screen-papers <folder>` — Quick paper screening
- Chat with the AI assistant for materials science questions

### Configuration

Set these environment variables:
- `SIC_LAB_KEY` — Web app access key (default: siclab2026)
- `DEEPSEEK_API_KEY` — DeepSeek API key (auto-loaded from OpenCode auth)
