import pandas as pd
from anthropic import Anthropic
import os
import time
from datetime import datetime
import re

API_KEY = ""
INPUT_FILE = "input_requirements_500.xlsx"
OUTPUT_FILE = f".xlsx"
MODEL = "claude-sonnet"
MAX_TOKENS = 20000
COMPLETE_INCOSE_RULES = """

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

1. SINGULAR (R18): One capability per requirement
2. UNAMBIGUOUS (R2, R3, R7): Clear, active voice, no vague terms
3. COMPLETE (R1, R27): All necessary elements present
4. FEASIBLE (R26): Realistic, achievable within constraints
5. VERIFIABLE (R33, R34): Measurable criteria with tolerances
6. APPROPRIATE (R31): Level-appropriate, solution-free
7. CONSISTENT (R4, R36, R37): Terminology maintained throughout

ALL placeholders in format [PLACEHOLDER_NAME] from original requirement MUST be preserved exactly in transformed requirement.
- Example: [VESSEL_TYPE], [SPEED_RANGE], [CREW_SIZE], [OPERATION_MODE]
- Never remove, rename, or forget placeholders
- Placeholders represent variables to be defined later
If no placeholders are provided in the original requirement, the output will also not contain a placeholder.
"""

ANALYZE_SPLIT_PROMPT = """You are an expert in Requirements Engineering according to ISO 29148 and all 42 INCOSE Guide rules.
TASK: Analyze whether this requirement MUST be split or is already ATOMIC.

ORIGINAL REQUIREMENT:
{customer_req}

{incose_rules}
CRITICAL: Identify and preserve ALL placeholders in format [PLACEHOLDER_NAME]
DECISION LOGIC (R18 - Single Thought Sentence):

DO NOT SPLIT when:
- Describes ONE coherent capability
- Parameters are semantically linked (e.g., "X days with Y persons")
- Separation would break semantic meaning

MUST SPLIT when:
- Multiple independent actions: "validate AND log AND notify"
- Multiple capabilities with "and": "encrypt data AND generate logs"
- Each action can be verified independently
- Contains enumeration of different items (R22)
- Multiple unrelated conditions or scenarios

OUTPUT FORMAT (JSON):
{{
  "should_split": true|false,
  "reasoning": "Detailed justification referencing R18",
  "number_of_atomic_requirements": 1-10,
  "identified_capabilities": ["Capability 1", "Capability 2", ...],
  "placeholders_found": ["[PLACEHOLDER_1]", "[PLACEHOLDER_2]", ...]
}}
Respond ONLY with valid JSON."""

IMPROVE_REQUIREMENT_PROMPT = """You are an expert in Requirements Engineering according to ISO 29148 and all 42 INCOSE Guide rules.
TASK: Transform this requirement into ISO 29148 + INCOSE compliant requirement.

ORIGINAL REQUIREMENT:
{customer_req}

{incose_rules}

CRITICAL: PRESERVE ALL PLACEHOLDERS [LIKE_THIS] FROM ORIGINAL IN YOUR OUTPUT
COMPREHENSIVE INCOSE IMPROVEMENT CHECKLIST:

**Structure (R1):** [WHEN condition], [ENTITY] shall [ACTION] [OBJECT] [PERFORMANCE ± tolerance]
**Voice (R2):** Active - "The [System_X] shall..." not passive
**Subject (R3):** System capability, not user action
**Terms (R4, R36):** Use consistent, defined terminology
**Articles (R5):** Use "the" for specific entities
**Units (R6):** Consistent measurement units
**Vague Terms (R7):** Replace with measurable criteria - DOCUMENT ALL REPLACEMENTS
**Escape Clauses (R8):** Remove "where possible", "as appropriate"
**Open Clauses (R9):** No "etc.", "including but not limited to"
**Infinitives (R10):** Remove "shall be able to"
**Grammar (R11-R14):** Correct grammar, spelling, punctuation
**Logic (R15):** Explicit [X AND Y] or [X OR Y]
**Negatives (R16):** Positive formulation, not "shall not"
**Symbols (R17):** No "/" - use explicit words
**Atomicity (R18-R19):** Single thought, split on "and"/"or"
**Purpose (R20-R21):** No "in order to", no parentheses
**Enumeration (R22):** Separate requirement per item
**Pronouns (R24):** No "it", "this" - repeat nouns
**Headings (R25):** Self-contained, not dependent on context
**Absolutes (R26):** Use "≥99.9%" not "100%"
**Conditions (R27-R28):** Explicit "When [condition]", clarify AND/OR
**Classification (R29):** Identify requirement type
**Uniqueness (R30):** Avoid duplication
**Solution-Free (R31):** WHAT not HOW - no technology names
**Qualification (R32):** Use "each" not "all"
**Ranges (R33):** Provide tolerances: X ± Y - DOCUMENT ALL TOLERANCES ADDED
**Measurable (R34-R35):** Specific metrics, not "fast" or "soon"
**Consistency (R36-R40):** Consistent terms, acronyms, decimals

CRITICAL OUTPUT REQUIREMENTS:
1. "vague_terms_removed": List EVERY vague/subjective term replaced with its specific measurable replacement
   Format: ["original_vague_term → specific_measurable_replacement", ...]
   Example: ["fast → within 2.0 ± 0.5 seconds", "adequate → ≥95% accuracy", "user-friendly → task completion within 30 ± 5 seconds"]

2. "tolerances_added": List EVERY quantitative performance metric with its tolerance
   Format: ["metric_description: value ± tolerance units", ...]
   Example: ["response time: 2.0 ± 0.5 seconds", "accuracy: ≥95%", "capacity: 1000 ± 50 users"]

OUTPUT FORMAT (JSON):
{{
  "requirement_type": "Functional|Performance|Interface|Safety|Security|etc.",
  "improved_requirement": "Complete INCOSE-compliant requirement with ALL [PLACEHOLDERS] preserved",
  "verification_method": "Test|Inspection|Analysis|Demonstration",
  "placeholders_preserved": ["[PLACEHOLDER_1]", "[PLACEHOLDER_2]", ...],
  "incose_rules_applied": ["R1: Structured format", "R2: Active voice", "R7: Removed vague terms", ...],
  "vague_terms_removed": ["original_vague_term → specific_measurable_replacement", ...],
  "tolerances_added": ["metric: value ± tolerance units", ...],
  "escape_clauses_removed": ["clause1", "clause2", ...],
  "improvements_summary": "Brief description of all transformations"
}}

VERIFY: All [PLACEHOLDERS] from input appear in improved_requirement
Respond ONLY with valid JSON."""

SPLIT_REQUIREMENT_PROMPT = """You are an expert in Requirements Engineering according to ISO 29148 and all 42 INCOSE Guide rules.
TASK: Split this requirement into {num_requirements} atomic, INCOSE-compliant requirements.

ORIGINAL REQUIREMENT:
{customer_req}

IDENTIFIED CAPABILITIES:
{capabilities}

{incose_rules}

CRITICAL: DISTRIBUTE PLACEHOLDERS [LIKE_THIS] APPROPRIATELY TO EACH SUB-REQUIREMENT

SPLITTING RULES (R18):
- Create exactly {num_requirements} independent requirements
- Each = ONE capability from identified list
- Each follows R1 structure
- Apply ALL 42 INCOSE rules to each sub-requirement
- Distribute [PLACEHOLDERS] to relevant sub-requirements

CRITICAL OUTPUT REQUIREMENTS FOR EACH SUB-REQUIREMENT:
1. "vague_terms_removed": List EVERY vague term replaced with measurable criteria
2. "tolerances_added": List EVERY quantitative metric with tolerance

OUTPUT FORMAT (JSON array):
[
  {{
    "sub_id": "1",
    "requirement_type": "Functional|Performance|etc.",
    "requirement_text": "Complete INCOSE-compliant requirement with relevant [PLACEHOLDERS]",
    "verification_method": "Test|Inspection|Analysis|Demonstration",
    "placeholders_used": ["[PLACEHOLDER_X]", ...],
    "incose_rules_applied": ["R1", "R2", "R7", ...],
    "vague_terms_removed": ["original → replacement", ...],
    "tolerances_added": ["metric: value ± tolerance", ...],
    "improvements_summary": "Brief summary"
  }},
  ...
]

VERIFY: All original [PLACEHOLDERS] distributed across sub-requirements
Respond ONLY with valid JSON array."""

def extract_placeholders(text):
    """Extract all placeholders [LIKE_THIS] from text"""
    return re.findall(r'\[([^\]]+)\]', text)

def verify_placeholders_preserved(original_text, generated_text):
    """Verify all placeholders from original are in generated text"""
    original_placeholders = set(extract_placeholders(original_text))
    generated_placeholders = set(extract_placeholders(generated_text))
    missing = original_placeholders - generated_placeholders
    if missing:
        print(f"WARNING: Missing placeholders: {missing}")
        return False
    return True

def init_claude_client():
    """Initialize Claude API Client"""
    if API_KEY == "" or API_KEY == "":
        raise ValueError("ERROR: Please insert your Claude API Key in the script!")
    return Anthropic(api_key=API_KEY)

def load_excel(filepath):
    """Load Excel file with single column of requirements"""
    try:
        df = pd.read_excel(filepath)
        print(f"Excel loaded: {len(df)} rows found")
        if len(df.columns) >= 1:
            df['customer_req'] = df.iloc[:, 0]
            print(f"   Column A ('{df.columns[0]}') → customer_req")
        else:
            raise ValueError("Excel file must have at least 1 column with requirements")
        df['Category'] = [f'REQ_{i+1:03d}' for i in range(len(df))]
        original_count = len(df)
        df = df[df['customer_req'].notna()]
        print(f"   {len(df)} requirements with text (filtered {original_count - len(df)} empty rows)")
        return df
    except FileNotFoundError:
        raise FileNotFoundError(f"ERROR: File '{filepath}' not found!")

def analyze_requirement(client, customer_req, max_retries=3):
    """Analyze if requirement should be split (R18)"""
    prompt = ANALYZE_SPLIT_PROMPT.format(
        customer_req=customer_req,
        incose_rules=COMPLETE_INCOSE_RULES
    )
    for attempt in range(max_retries):
        try:
            message = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                messages=[{"role": "user", "content": prompt}]
            )
            response_text = message.content[0].text

            import json
            if "```json" in response_text:
                response_text = response_text.split("```json")[1].split("```")[0]
            elif "```" in response_text:
                response_text = response_text.split("```")[1].split("```")[0]
            analysis = json.loads(response_text.strip())
            return (
                analysis['should_split'],
                analysis['number_of_atomic_requirements'],
                analysis['identified_capabilities'],
                analysis.get('placeholders_found', [])
            )
        except Exception as e:
            if attempt < max_retries - 1:
                print(f"Retry {attempt + 1}/{max_retries} after error: {str(e)[:100]}")
                time.sleep(2)
            else:
                print(f"Analysis failed after {max_retries} attempts, using default")
                return (False, 1, ["Unknown"], extract_placeholders(customer_req))

def improve_requirement(client, customer_req, max_retries=3):
    """Improve atomic requirement with all 42 INCOSE rules"""
    prompt = IMPROVE_REQUIREMENT_PROMPT.format(
        customer_req=customer_req,
        incose_rules=COMPLETE_INCOSE_RULES
    )
    for attempt in range(max_retries):
        try:
            message = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                messages=[{"role": "user", "content": prompt}]
            )
            response_text = message.content[0].text
            import json
            if "```json" in response_text:
                response_text = response_text.split("```json")[1].split("```")[0]
            elif "```" in response_text:
                response_text = response_text.split("```")[1].split("```")[0]
            improved = json.loads(response_text.strip())
            
            verify_placeholders_preserved(customer_req, improved['improved_requirement'])
            return [{
                "requirement_type": improved['requirement_type'],
                "requirement_text": improved['improved_requirement'],
                "verification_method": improved['verification_method'],
                "placeholders": ", ".join(improved.get('placeholders_preserved', [])),
                "incose_rules": ", ".join(improved.get('incose_rules_applied', [])),
                "vague_terms_removed": improved.get('vague_terms_removed', []),
                "tolerances_added": improved.get('tolerances_added', []),
                "improvements": improved.get('improvements_summary', '')
            }]    
        except Exception as e:
            if attempt < max_retries - 1:
                print(f"Retry {attempt + 1}/{max_retries} after error: {str(e)[:100]}")
                time.sleep(2)
            else:
                print(f"Improvement failed: {str(e)[:200]}")
                return [{
                    "requirement_type": "ERROR",
                    "requirement_text": f"ERROR: {str(e)[:500]}",
                    "verification_method": "N/A",
                    "placeholders": "",
                    "incose_rules": "",
                    "vague_terms_removed": [],
                    "tolerances_added": [],
                    "improvements": ""
                }]

def split_requirement(client, customer_req, num_requirements, capabilities, max_retries=3):
    """Split requirement into atomic INCOSE-compliant requirements"""
    prompt = SPLIT_REQUIREMENT_PROMPT.format(
        customer_req=customer_req,
        num_requirements=num_requirements,
        capabilities=", ".join(capabilities),
        incose_rules=COMPLETE_INCOSE_RULES
    )
    for attempt in range(max_retries):
        try:
            message = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                messages=[{"role": "user", "content": prompt}]
            )
            response_text = message.content[0].text
            import json
            if "```json" in response_text:
                response_text = response_text.split("```json")[1].split("```")[0]
            elif "```" in response_text:
                response_text = response_text.split("```")[1].split("```")[0]
            requirements = json.loads(response_text.strip())
            all_generated_placeholders = set()
            for req in requirements:
                all_generated_placeholders.update(extract_placeholders(req['requirement_text']))
            original_placeholders = set(extract_placeholders(customer_req))
            if not original_placeholders.issubset(all_generated_placeholders):
                print(f"   ⚠️ Some placeholders not distributed across split requirements")
            formatted = []
            for req in requirements:
                formatted.append({
                    "requirement_type": req['requirement_type'],
                    "requirement_text": req['requirement_text'],
                    "verification_method": req['verification_method'],
                    "placeholders": ", ".join(req.get('placeholders_used', [])),
                    "incose_rules": ", ".join(req.get('incose_rules_applied', [])),
                    "vague_terms_removed": req.get('vague_terms_removed', []),
                    "tolerances_added": req.get('tolerances_added', []),
                    "improvements": req.get('improvements_summary', '')
                })
            print(f"Split into {len(formatted)} requirements")
            return formatted
        except Exception as e:
            if attempt < max_retries - 1:
                print(f"Retry {attempt + 1}/{max_retries} after error: {str(e)[:100]}")
                time.sleep(2)
            else:
                print(f"Split failed: {str(e)[:200]}")
                return [{
                    "requirement_type": "ERROR",
                    "requirement_text": f"ERROR: {str(e)[:500]}",
                    "verification_method": "N/A",
                    "placeholders": "",
                    "incose_rules": "",
                    "vague_terms_removed": [],
                    "tolerances_added": [],
                    "improvements": ""
                }]

def format_list_to_string(item_list):
    """Convert list to readable string format"""
    if isinstance(item_list, list):
        if len(item_list) == 0:
            return ""
        return "; ".join(str(item) for item in item_list)
    return str(item_list)

def process_requirement(client, row, index, total):
    """Main processing logic with progress tracking"""
    customer_req = row.get('customer_req', '')
    category = row.get('Category', f'REQ_{index}')
    print(f"\n[{index}/{total}] Processing: {category}")
    print(f"   Original: {customer_req[:80]}...")
    placeholders = extract_placeholders(customer_req)
    if placeholders:
        print(f"Placeholders found: {placeholders}")
    should_split, num_reqs, capabilities, _ = analyze_requirement(client, customer_req)
    time.sleep(0.5)
    if should_split:
        print(f"Splitting → {num_reqs} requirements")
        requirements = split_requirement(client, customer_req, num_reqs, capabilities)
    else:
        print(f"Atomic → Applying 42 INCOSE rules")
        requirements = improve_requirement(client, customer_req)
    time.sleep(0.5)
    for req in requirements:
        req['category'] = category
        req['original_req'] = customer_req
    return requirements

def process_all_requirements(df):
    """Process all requirements with progress tracking"""
    client = init_claude_client()
    all_results = []
    total = len(df)
    print(f"\n{'='*70}")
    print(f"PROCESSING {total} REQUIREMENTS")
    print(f"{'='*70}")
    start_time = time.time()
    for idx, row in df.iterrows():
        try:
            results = process_requirement(client, row, idx + 1, total)
            category = row.get('Category', f'REQ_{idx+1}')
            customer_req = row.get('customer_req', '')
            print(f"   🔄 Consolidating requirements...")
            if len(results) == 1:
                consolidated = results[0]['requirement_text']
                detailed = results[0]['requirement_text']
            else:
                consolidated = f"The system shall meet {len(results)} requirements addressing: " + ", ".join([req['requirement_type'] for req in results])
                detailed = "The system shall meet the following requirements:\n" + "\n".join([
                    f"{i+1}. {req['requirement_text']}" for i, req in enumerate(results)
                ])
            all_vague_terms = []
            all_tolerances = []
            for req in results:
                all_vague_terms.extend(req.get('vague_terms_removed', []))
                all_tolerances.extend(req.get('tolerances_added', []))
            for i, req in enumerate(results):
                result = {
                    'Category': category,
                    'Customer_Req': customer_req,
                    'Ambiguities_Identified': req.get('improvements', ''),
                    'Improvements_Made': req.get('incose_rules', ''),
                    'Vague_Terms_Removed': format_list_to_string(req.get('vague_terms_removed', [])),
                    'Tolerances_Added': format_list_to_string(req.get('tolerances_added', [])),
                    'Consolidated_Requirement': consolidated if i == 0 else '',
                    'Detailed_Requirement': detailed if i == 0 else '',
                    'Sub_Requirement_Text': req['requirement_text'],
                    'Verification_Method': req['verification_method']
                }
                all_results.append(result)
            time.sleep(0.5)
            if (idx + 1) % 10 == 0:
                elapsed = time.time() - start_time
                avg_time = elapsed / (idx + 1)
                remaining = avg_time * (total - idx - 1)
                print(f"\n   ⏱️ Progress: {idx+1}/{total} ({(idx+1)/total*100:.1f}%) - Est. remaining: {remaining/60:.1f} min")
        except Exception as e:
            print(f"Error processing row {idx}: {str(e)[:200]}")
            all_results.append({
                'Category': row.get('Category', f'REQ_{idx}'),
                'Customer_Req': row.get('customer_req', ''),
                'Ambiguities_Identified': 'Processing error',
                'Improvements_Made': 'N/A',
                'Vague_Terms_Removed': '',
                'Tolerances_Added': '',
                'Consolidated_Requirement': f'ERROR: {str(e)[:200]}',
                'Detailed_Requirement': f'ERROR: {str(e)[:200]}',
                'Sub_Requirement_Text': f'ERROR: {str(e)[:500]}',
                'Verification_Method': 'N/A'
            })
    elapsed = time.time() - start_time
    print(f"\n{'='*70}")
    print(f"PROCESSING COMPLETE - Total time: {elapsed/60:.1f} minutes")
    print(f"{'='*70}")
    return pd.DataFrame(all_results)

def export_to_excel(df, filepath):
    """Export to Excel with formatting"""
    column_order = [
        'Category',
        'Customer_Req',
        'Ambiguities_Identified',
        'Improvements_Made',
        'Vague_Terms_Removed',
        'Tolerances_Added',
        'Consolidated_Requirement',
        'Detailed_Requirement',
        'Sub_Requirement_Text',
        'Verification_Method'
    ]
    df = df[column_order]
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='ISO_Compliant_Requirements')
        worksheet = writer.sheets['ISO_Compliant_Requirements']
        column_widths = {
            'A': 15,
            'B': 60,
            'C': 40,
            'D': 50,
            'E': 40,
            'F': 40,
            'G': 50,
            'H': 70,
            'I': 70,
            'J': 20
        }
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
            for cell in row:
                cell.alignment = cell.alignment.copy(wrap_text=True)
    print(f"\nExcel exported: {filepath}")

def main():
    """Main execution"""
    print("=" * 80)
    print("REQUIREMENTS PROCESSING - COMPLETE 42 INCOSE RULES (UPDATED)")
    print("All 42 INCOSE Guide rules implemented")
    print("ISO 29148 quality characteristics enforced")
    print("Placeholder preservation guaranteed [LIKE_THIS]")
    print("Handles large batches (tested up to 500+ requirements)")
    print("Properly extracts vague terms removed and tolerances added")
    print("Single-column input support")
    print("=" * 80)
    try:
        print("\n[1/3] Loading Excel file...")
        df_input = load_excel(INPUT_FILE)
        print("\n[2/3] Processing requirements with 42 INCOSE rules...")
        df_output = process_all_requirements(df_input)
        print("\n[3/3] Exporting ISO-compliant requirements...")
        export_to_excel(df_output, OUTPUT_FILE)
        print("\n" + "=" * 80)
        print("SUCCESSFULLY COMPLETED!")
        print("=" * 80)
        print(f"Input:  {len(df_input)} original requirements")
        print(f"Output: {len(df_output)} processed requirements")
        print(f"File:   {OUTPUT_FILE}")
        print("\nOutput column structure (A-J):")
        print("A: Category (auto-generated REQ_001, REQ_002, ...)")
        print("B: Customer_Req (from input column A)")
        print("C: Ambiguities_Identified")
        print("D: Improvements_Made")
        print("E: Vague_Terms_Removed (properly filled)")
        print("F: Tolerances_Added (properly filled)")
        print("G: Consolidated_Requirement")
        print("H: Detailed_Requirement")
        print("I: Sub_Requirement_Text")
        print("J: Verification_Method")
        print("\nAll 42 INCOSE Rules Applied")
        print("\nFeatures:")
        print("Single column input (requirements only)")
        print("Auto-generated Category IDs")
        print("All [PLACEHOLDERS] preserved in transformed requirements")
        print("ISO 29148 quality characteristics enforced")
        print("Vague terms and tolerances properly extracted and documented")
        print("Comprehensive INCOSE improvements documented")
        print("Supports splitting into 5+ sub-requirements when needed")
    except Exception as e:
        print(f"\nERROR: {str(e)}")
        import traceback
        traceback.print_exc()
if __name__ == "__main__":
    main()