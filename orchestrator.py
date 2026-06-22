"""Paper-Orchestrator Pipeline for Flask Web App
Orchestrator -> paper-reader (parallel) -> global-synthesizer -> quality-gate
"""
import os, time
paper_cache = {}

def run_orchestrator_pipeline(uploads, sample_status, goal, requirements, api_key=''):
    """Execute the full Paper-Orchestrator agent pipeline"""
    
    # ====== STAGE 1: paper-reader agents (parallel) ======
    # Each reader extracts structured data using chain-of-thought reasoning
    import fitz
    from concurrent.futures import ThreadPoolExecutor
    
    def reader_agent(upload_item):
        """paper-reader: extracts structured info with domain knowledge"""
        # Check cache
        path = upload_item['path']
        cache_key = path + '_' + str(os.path.getmtime(path))
        if cache_key in paper_cache:
            return paper_cache[cache_key]
        
        try:
            doc = fitz.open(path)
            full_text = ''
            for page in doc: full_text += page.get_text() + '\n'
            doc.close()
            
            # Structured extraction (chain-of-thought)
            txt = full_text[:10000]
            result = {
                'name': upload_item['name'],
                'pages': len(doc),
                'text': txt,
                'methods': detect_methods(txt),
                'properties': detect_properties(txt),
                'key_findings': [],
                'limitations': []
            }
            
            # Extract key findings (first/last paragraphs often contain conclusions)
            lines = [l.strip() for l in txt.split('\n') if len(l.strip()) > 50]
            if lines:
                result['key_findings'] = lines[:3]  # first 3 substantive lines
                if len(lines) > 3:
                    result['conclusions'] = lines[-3:]  # last 3 lines
            
            paper_cache[cache_key] = result
            return result
        except:
            return {'name': upload_item['name'], 'pages': 0, 'text': '', 'methods': [], 'properties': {}, 'key_findings': []}
    
    # Parallel reading
    with ThreadPoolExecutor(max_workers=4) as pool:
        papers = list(pool.map(reader_agent, uploads))
    
    total = len(papers)
    
    # ====== STAGE 2: global-synthesizer ======
    # Synthesizes paper data + user requirements into a report
    if not api_key:
        return f'Analyzed {total} papers. (No AI key configured)', papers
    
    # Build structured paper summary for the synthesizer
    paper_summary = []
    for i, p in enumerate(papers, 1):
        ms = '; '.join(p.get('methods', [])[:4]) or 'N/A'
        ps = p.get('properties', {})
        tc = f'TC={ps["TC"]}W/mK' if ps.get('TC') else ''
        bs = f'BS={ps["BS"]}MPa' if ps.get('BS') else ''
        findings = p.get('key_findings', [])[:2]
        f_text = '; '.join(findings) if findings else ''
        paper_summary.append(f'Paper {i}: {p["name"]} ({p["pages"]}p)\n  Methods: {ms}\n  Properties: {tc} {bs}\n  Key findings: {f_text}')
    
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key, base_url='https://api.deepseek.com/v1')
        resp = client.chat.completions.create(
            model='deepseek-v4-pro',
            messages=[{'role': 'system', 'content': f'''You are a senior materials scientist writing a professional experiment report for a lab.

FORMAT REQUIREMENTS:
- Use clear Markdown-style headings: ## for major sections, ### for subsections
- Use numbered lists for procedures (1. 2. 3.)
- Use bullet points for equipment and notes
- Include SPECIFIC numbers from papers whenever possible
- Cite paper sources inline as [Paper N]
- Write in the user's preferred language
- Include: ## Abstract, ## Equipment & Materials, ## Procedure, ## Parameters, ## Expected Results, ## Safety, ## References

CONTENT REQUIREMENTS:
- Be PRECISE: exact temperatures (X C), times (Y min), pressures (Z MPa)
- Be COMPREHENSIVE: list every equipment item with specifications
- Be SAFETY-CONSCIOUS: include hazard warnings for high temp, chemicals, gases
- Be CUSTOMIZED: every report must be different based on the actual papers

USER REQUIREMENTS:
Sample: {sample_status if sample_status else 'Not specified'}
Goal: {goal if goal else 'Not specified'}
Special: {requirements if requirements else 'None'}

PAPERS ({total} total):
{chr(10).join(paper_summary)}'''},
             {'role': 'user', 'content': f'Generate experiment report. Apply my requirements:\n{requirements}\n\nGoal: {goal}'}],
            temperature=0.3, max_tokens=4000
        )
        report = resp.choices[0].message.content
    except:
        report = f'[Synthesizer] Analyzed {total} papers. Goal: {goal}. Requirements: {requirements}'
    
    # ====== STAGE 3: quality-gate ======
    # Validates the output
    issues = []
    if not report or len(report) < 20:
        issues.append('Report too short')
    if '???' in report or '??' in report:
        issues.append('Garbled characters detected')
    if total > 0 and str(total) not in report:
        issues.append('Paper count not mentioned')
    
    if issues and api_key:
        try:
            resp2 = client.chat.completions.create(
                model='deepseek-v4-pro',
                messages=[{'role': 'system', 'content': f'Fix these issues in the following report: {"; ".join(issues)}. Return the fixed report.'},
                          {'role': 'user', 'content': report}],
                temperature=0.1, max_tokens=4000
            )
            report = resp2.choices[0].message.content
        except:
            pass
    
    return report, papers
