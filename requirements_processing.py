import pandas as pd
from anthropic import Anthropic
import os
import time
from datetime import datetime
import re

API_KEY = ""
INPUT_FILE = "inpuc_vague_requirements_500.xlsx"
OUTPUT_FILE = f"output_vague_requirements_500.xlsx"
MODEL = "claude-sonnet"
MAX_TOKENS = 20000
MAX_RETRIES = 3
COMPLETE_INCOSE_RULES = """

INCOSE_RULES = """
R1 – Structured Statements
- Use consistent pattern: [WHEN condition], [ENTITY] shall [ACTION] [OBJECT] [PERFORMANCE ± tolerance]
- Example: "When processing user queries, the Database_System shall return search results within 2.0 ± 0.5 seconds"
R2 – Active Voice
- Place responsible entity at beginning: "The Security_Module shall encrypt..." not "Data shall be encrypted..."
R3 – Appropriate Subject-Verb
- System requirements have system as subject, not users
- Good: "The Authentication_System shall prompt..." not "The user shall enter..."
R4 – Defined Terms
- All technical terms must be in glossary and used consistently
- Maintain terminology consistency across all artifacts
R5 – Definite Articles
- Use "the" for specific entities: "the Database_System" not "a database system"
R6 – Common Units of Measure
- Use consistent units throughout (no mixing metric/imperial)
R7 – Vague Terms
- Replace subjective terms with measurable criteria
- Avoid: "adequate", "reasonable", "user-friendly", "fast", "robust", "flexible"
R8 – Escape Clauses
- Eliminate: "where possible", "as appropriate", "if necessary", "to the extent possible"
R9 – Open-Ended Clauses
- Avoid: "including but not limited to", "etc.", "and so on"
- Explicitly list all items or create separate requirements
R10 – Superfluous Infinitives
- Remove "shall be able to" → use direct "shall"
- Remove "shall be capable of" → use direct "shall"
R11 – Separate Clauses
- Each condition or qualification in its own clause for clarity
R12 – Correct Grammar
- Ensure grammatically correct statements, critical for international teams
R13 – Correct Spelling
- Check spelling, watch for correctly spelled wrong words
R14 – Correct Punctuation
- Use proper punctuation to clarify clause relationships
R15 – Logical Expressions
- Use explicit notation: [X AND Y], [X OR Y] instead of ambiguous constructions
R16 – Use of "Not"
- Avoid negative requirements ("shall not fail")
- Use positive formulations: "shall have ≥99.9% availability"
R17 – Use of Oblique Symbol
- Don't use "/" - it can mean "and", "or", "per", or alternatives
- Use explicit language instead
R18 – Single Thought Sentence
- One requirement = one capability/action
- Split compound requirements: "validate AND log AND notify" → 3 requirements
- Exception: Semantically linked parameters for ONE capability stay together
R19 – Combinators
- Words "and", "or", "then" often indicate multiple thoughts → split
R20 – Purpose Phrases
- Avoid "in order to", "so that" in requirement text
- Put explanations in rationale attributes
R21 – Parentheses
- Avoid parenthetical information in requirements
- Move supplementary info to rationale
R22 – Enumeration
- Don't list multiple items in one requirement
- Create separate requirement for each enumerated item
R23 – Supporting Diagrams
- Reference diagrams/models for complex behaviors
- Don't try to capture everything in text
R24 – Pronouns
- Avoid "it", "they", "this", "that"
- Repeat nouns for self-contained statements
R25 – Headings
- Requirements must be complete without depending on headings
- Each requirement understandable in isolation
R26 – Absolutes
- Avoid "100%", "always", "never", "all" unless truly absolute
- Use realistic values: "≥99.9%" instead of "100%"
R27 – Explicit Conditions
- State all applicable conditions directly
- "When transmitting over public networks..." not just "shall encrypt"
R28 – Multiple Conditions
- Clarify AND vs OR when multiple conditions apply
- "[Condition_A AND Condition_B]" or "[Condition_A OR Condition_B]"
R29 – Classification
- Classify by type: functional, performance, interface, safety, security
- Enables gap analysis and conflict detection
R30 – Unique Expression
- Each requirement appears exactly once
- No duplication with different wording
R31 – Solution Free
- Describe WHAT (capabilities), not HOW (implementation)
- Avoid: "MySQL", "REST API", "Python" unless truly constrained
R32 – Universal Qualification
- Use "each" not "all", "any", "both"
- Clarifies: applies to every individual item, not collection as whole
R33 – Range of Values
- Provide tolerance ranges: "2.0 ± 0.3 seconds"
- Formats: X ± Y, X +Y/-Z, ≥X, ≤X, X to Y
R34 – Measurable Performance
- Replace subjective terms with specific measurable criteria
- "fast" → "within 2.0 ± 0.5 seconds"
R35 – Temporal Dependencies
- Replace vague terms: "eventually", "soon", "before"
- Use specific time constraints: "within 5.0 ± 1.0 minutes"
R36 – Consistent Terms and Units
- Identical terminology in requirements, design, tests, manuals
- Maintain and enforce project glossary
R37 – Acronyms
- Use same acronym consistently throughout
- Don't mix "GPS" and "Global Positioning System" randomly
R38 – Abbreviations
- Avoid unless necessary and clearly defined
- Many have multiple meanings depending on context
R39 – Style Guide
- Follow organization-wide standards for patterns, attributes, formatting
R40 – Decimal Format
- Consistent decimal notation and significant digits
- Don't mix "5.0" and "5.00" randomly
R41 – Related Requirements
- Group related requirements logically
- Helps identify gaps and conflicts
R42 – Structured Sets
- Use consistent organizational templates
- Ensure all requirement types considered: functional, performance, interface, safety, security

Quality Characteristics:
1. SINGULAR (R18): One capability per requirement
2. UNAMBIGUOUS (R2, R3, R7): Clear, active voice, no vague terms
3. COMPLETE (R1, R27): All necessary elements present
4. FEASIBLE (R26): Realistic, achievable within constraints
5. VERIFIABLE (R33, R34): Measurable criteria with tolerances
6. APPROPRIATE (R31): Level-appropriate, solution-free
7. CONSISTENT (R4, R36, R37): Terminology maintained throughout

ALL placeholders in format [PLACEHOLDER_NAME] from original requirement MUST be preserved exactly in transformed requirement.
"""

PROMPTS = {
    "analyze": """Analyze if requirement must SPLIT (R18).
REQ: {req}
{rules}
JSON only: {{"should_split": bool, "num": 1-10, "capabilities": [...], "placeholders": [...]}}""",
    "improve": """Transform to INCOSE compliance.
REQ: {req}
{rules}
JSON only: {{"type": "Functional|Performance|Interface|Safety|Security", "requirement": "...", "verification": "Test|Inspection|Analysis|Demonstration", "placeholders": [...], "rules": [...], "vague_removed": ["old → new"], "tolerances": ["metric: val ± tol"], "summary": "..."}}""",
    "split": """Split into {num} atomic requirements.
REQ: {req}
CAPABILITIES: {caps}
{rules}
JSON array only: [{{"id": "1", "type": "...", "requirement": "...", "verification": "...", "placeholders": [...], "rules": [...], "vague_removed": [...], "tolerances": [...], "summary": "..."}}]"""
}
def extract_placeholders(text):
    return re.findall(r'\[([^\]]+)\]', text or "")

def parse_json(text):
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0]
    elif "```" in text:
        text = text.split("```")[1].split("```")[0]
    return json.loads(text.strip())

def call_api(client, prompt):
    for attempt in range(MAX_RETRIES):
        try:
            resp = client.messages.create(
                model=MODEL, max_tokens=MAX_TOKENS,
                messages=[{"role": "user", "content": prompt}]
            )
            return parse_json(resp.content[0].text)
        except Exception as e:
            if attempt < MAX_RETRIES - 1:
                time.sleep(2)
            else:
                raise

def fmt(items):
    return "; ".join(str(i) for i in items) if isinstance(items, list) and items else ""

def process(client, req, idx, total):
    print(f"[{idx}/{total}] {req[:50]}...")
    try:
        analysis = call_api(client, PROMPTS["analyze"].format(req=req, rules=INCOSE_RULES))
        should_split = analysis.get("should_split", False)
        num = analysis.get("num", 1)
        caps = analysis.get("capabilities", [])
    except:
        should_split, num, caps = False, 1, []
    time.sleep(0.5)
    try:
        if should_split and num > 1:
            print(f"  → Split: {num}")
            results = call_api(client, PROMPTS["split"].format(
                req=req, num=num, caps=", ".join(caps), rules=INCOSE_RULES
            ))
        else:
            print(f"  → Improve")
            results = [call_api(client, PROMPTS["improve"].format(req=req, rules=INCOSE_RULES))]
    except Exception as e:
        results = [{"type": "ERROR", "requirement": str(e), "verification": "N/A"}]
    return results

def main():
    if not API_KEY:
        raise ValueError("Set API_KEY")
    client = Anthropic(api_key=API_KEY)
    df = pd.read_excel(INPUT_FILE)
    df['customer_req'] = df.iloc[:, 0]
    df = df[df['customer_req'].notna()]
    all_results = []
    start = time.time()
    for idx, row in df.iterrows():
        req = row['customer_req']
        cat = f'REQ_{idx+1:03d}'
        try:
            results = process(client, req, idx + 1, len(df))
            consolidated = results[0].get('requirement', '') if len(results) == 1 else f"System shall meet {len(results)} requirements."
            detailed = consolidated if len(results) == 1 else "\n".join(f"{i+1}. {r.get('requirement', '')}" for i, r in enumerate(results))
            
            for i, r in enumerate(results):
                all_results.append({
                    'Category': cat,
                    'Customer_Req': req,
                    'Ambiguities_Identified': r.get('summary', ''),
                    'Improvements_Made': fmt(r.get('rules', [])),
                    'Vague_Terms_Removed': fmt(r.get('vague_removed', [])),
                    'Tolerances_Added': fmt(r.get('tolerances', [])),
                    'Consolidated_Requirement': consolidated if i == 0 else '',
                    'Detailed_Requirement': detailed if i == 0 else '',
                    'Sub_Requirement_Text': r.get('requirement', ''),
                    'Verification_Method': r.get('verification', '')
                })
            time.sleep(0.5)
            if (idx + 1) % 10 == 0:
                elapsed = time.time() - start
                print(f"  {idx+1}/{len(df)} - ~{(elapsed/(idx+1))*(len(df)-idx-1)/60:.1f}m left")         
        except Exception as e:
            all_results.append({
                'Category': cat, 'Customer_Req': req, 'Sub_Requirement_Text': f'ERROR: {e}',
                'Ambiguities_Identified': '', 'Improvements_Made': '', 'Vague_Terms_Removed': '',
                'Tolerances_Added': '', 'Consolidated_Requirement': '', 'Detailed_Requirement': '',
                'Verification_Method': ''
            })
    df_out = pd.DataFrame(all_results)[[
        'Category', 'Customer_Req', 'Ambiguities_Identified', 'Improvements_Made',
        'Vague_Terms_Removed', 'Tolerances_Added', 'Consolidated_Requirement',
        'Detailed_Requirement', 'Sub_Requirement_Text', 'Verification_Method'
    ]]    
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as w:
        df_out.to_excel(w, index=False, sheet_name='Requirements')
        ws = w.sheets['Requirements']
        for i, width in enumerate([15, 60, 40, 50, 40, 40, 50, 70, 70, 20]):
            ws.column_dimensions[chr(65 + i)].width = width
    print(f"\nDone: {len(df)} → {len(df_out)} in {(time.time()-start)/60:.1f}m → {OUTPUT_FILE}")
if __name__ == "__main__":
    main()
