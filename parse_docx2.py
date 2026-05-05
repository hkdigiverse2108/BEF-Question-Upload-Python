"""
parse_docx.py - V3 Robust
"""
import docx
import json
import re
import os

# Use a very specific marker that won't conflict with anything
IMG_START = "[[IMG_START]]"
IMG_END = "[[IMG_END]]"

def extract_field(text: str) -> dict:
    pattern = re.escape(IMG_START) + r"\s*(.*?)\s*" + re.escape(IMG_END)
    m = re.search(pattern, text, re.I)
    
    # Extract only the first image path found
    image_path = m.group(1).strip() if m else ""
    
    # Remove ALL instances of the image marker from the text
    clean_text = re.sub(pattern, '', text, flags=re.I).strip()
    
    return {
        "text": clean_text,
        "image": image_path
    }

def parse_block_regex(text: str) -> dict | None:
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if len(lines) < 4: return None

    q_marker = r'(?:\(\d+\)|\d+\s*[\.\)\:])'
    pattern = re.escape(IMG_START) + r'.*?' + re.escape(IMG_END) + r'\s*' + q_marker + r'|^' + q_marker
    
    q_idx = -1
    for idx, line in enumerate(lines):
        if re.match(pattern, line):
            q_idx = idx
            break
    
    if q_idx == -1: return None
    
    meta = lines[:q_idx]
    subtopic = meta[0] if len(meta) > 0 else "Unknown"
    category = meta[1].lower() if len(meta) > 1 else "aptitude"
    # Normalize question_type: lower case and replace spaces with hyphens (e.g., "Normal Csat" -> "normal-csat")
    question_type = meta[2].lower().strip().replace(" ", "-") if len(meta) > 2 else "normal"

    i = q_idx
    question_lines = []
    
    # Robust question number extraction: check first 4 lines of the block
    q_number = "0"
    marker_pattern = r'^(?:' + re.escape(IMG_START) + r'.*?' + re.escape(IMG_END) + r'\s*)?(?:\((\d+)\)|(\d+)\s*[\.\)\:])'
    for k in range(q_idx, min(q_idx + 5, len(lines))):
        m_num = re.match(marker_pattern, lines[k])
        if m_num:
            q_number = m_num.group(1) or m_num.group(2)
            break
    
    first_q_line = re.sub(r'^(?:' + re.escape(IMG_START) + r'.*?' + re.escape(IMG_END) + r'\s*)?(?:\(\d+\)|\d+\s*[\.\)\:])\s*', '', lines[q_idx]).strip()
    # If the first line was just the number or metadata, we might need to skip or handle properly
    # But usually, it's safer to just start collecting.
    question_lines.append(first_q_line)
    i += 1

    is_options = lambda s: bool(re.match(r'^(?:\([a-dA-Dअ-दए-डीक-घ]\)|[A-D]\s*[:\)])', s, re.I))
    is_pair = lambda s: bool(re.match(r'^(?:Pair\s*\d+\s*:|.*?\s*(?:=|—|–|-)\s*.*)', s, re.I))
    is_stmt = lambda s: bool(re.match(r'^(\d+\s*[\.\)]|\(\d+\))', s, re.I))
    # Note: Hindi doc uses Devanagari visarga 'ः' instead of Latin ':' in 'Answerः' / 'Solutionः'
    is_ans = lambda s: bool(re.match(r'^\s*(Answer|उत्तर)\s*[:\u0903]', s, re.I))
    is_sol = lambda s: bool(re.match(r'^\s*(Solution|व्याख्या|समाधान)(?:[:\u0903]|\b)', s, re.I))

    # Find the Answer line index to infer hidden options
    ans_idx = -1
    for k in range(q_idx, len(lines)):
        if is_ans(lines[k]):
            ans_idx = k
            break

    def is_options_func(idx, s):
        if re.match(r'^(?:\([a-dA-Dअ-दए-डीक-घ]\)|[A-D]\s*[:\)])', s, re.I): return True
        if ans_idx != -1 and ans_idx - 4 <= idx < ans_idx: return True
        return False

    is_options = lambda idx, s: is_options_func(idx, s)
    is_options_compat = lambda s: bool(re.match(r'^(?:\([a-dA-Dअ-दए-डीक-घ]\)|[A-D]\s*[:\)])', s, re.I))

    # 1. Question text starts at q_idx
    # We want to separate the main question text from statements/pairs.
    
    # Trigger patterns
    stmt_trigger = r'consider\b.*?statements|निम्नलिखित.*?कथन|दिए गए कथनों|विचार करें|Statement'
    pair_trigger = r'consider\b.*?pairs|निम्नलिखित.*?युग्म|जोड़ों पर विचार|Pair'
    lq_trigger = r'(?:How many|Which of the|Which one|उपरोक्त|Choose the correct|Select the correct|दिए गए|नीचे दिए गए|कितने जोड़े|किन कथनों|सही उत्तर|सही है|सही हैं|मेल खाते हैं).*?(?:statements|pairs|answer|code|कथन|युग्म|सही है|सही हैं|मेल खाते हैं|कूट|चुनें|चुनिए)'

    i = q_idx
    question_lines = []
    
    # First line (with number)
    first_q_line = re.sub(r'^(?:' + re.escape(IMG_START) + r'.*?' + re.escape(IMG_END) + r'\s*)?(?:\(\d+\)|\d+\.)\s*', '', lines[i]).strip()
    question_lines.append(first_q_line)
    i += 1

    # Logic to collect question text until statements or options
    is_stmt_mode = "statement" in question_type
    is_pair_mode = "pair" in question_type

    if re.search(pair_trigger, first_q_line, re.I):
        is_pair_mode = True
        is_stmt_mode = False
        question_type = "pair"
    elif re.search(stmt_trigger, first_q_line, re.I):
        is_stmt_mode = True
        is_pair_mode = False
        question_type = "statement"

    while i < len(lines):
        if is_options(i, lines[i]) or is_ans(lines[i]): break
        
        # If we see a statement-like number (1., 2.), and we are in statement mode,
        # it might be the start of statements. 
        # But we also look for the "Consider the following..." header.
        if is_stmt_mode and (is_stmt(lines[i]) or re.search(stmt_trigger, question_lines[-1], re.I)):
             if not re.search(stmt_trigger, lines[i], re.I): # Don't break if this line IS the trigger
                break
        if is_pair_mode and (is_pair(lines[i]) or re.search(pair_trigger, question_lines[-1], re.I)):
            if not re.search(pair_trigger, lines[i], re.I):
                break

        question_lines.append(lines[i])
        i += 1

    statements = []
    pairs = []
    lastQuestion = ""

    if is_stmt_mode:
        while i < len(lines) and not is_options(i, lines[i]) and not is_ans(lines[i]) and not is_sol(lines[i]) and not re.search(lq_trigger, lines[i], re.I):
            # Accept numbered (1. / 1)) OR un-numbered statement lines
            # Strip markers like 1. or A. or (1)
            text = re.sub(r'^(\d+|[A-Za-z])[\.\)]\s*|^\(\d+\)\s*|^\([A-Za-z]\)\s*', '', lines[i]).strip()
            if text: statements.append(extract_field(text))
            i += 1
        lq_parts = []
        while i < len(lines) and not is_options(i, lines[i]) and not is_ans(lines[i]):
            lq_parts.append(lines[i])
            i += 1
        lastQuestion = extract_field(' '.join(lq_parts).strip())
    elif is_pair_mode:
        while i < len(lines) and not is_options(i, lines[i]) and not is_ans(lines[i]) and not re.search(lq_trigger, lines[i], re.I):
            # Check if it's a pair line (contains = or dash)
            if re.search(r'(=|—|–|-)', lines[i]):
                parts = re.split(r'\s*(?:=|—|–|-)\s*', lines[i], maxsplit=1)
                row = []
                for p in parts:
                    # Strip markers like 1. or A. or (1)
                    clean_p = re.sub(r'^(\d+|[A-Za-z])[\.\)]\s*|^\(\d+\)\s*|^\([A-Za-z]\)\s*', '', p.strip()).strip()
                    row.append(extract_field(clean_p))
                if len(row) >= 2: pairs.append(row)
            i += 1
        lq_parts = []
        while i < len(lines) and not is_options(i, lines[i]) and not is_ans(lines[i]):
            lq_parts.append(lines[i])
            i += 1
        lastQuestion = extract_field(' '.join(lq_parts).strip())
    else:
        # Normal type: question text already collected until options
        pass

    question_field = extract_field('\n'.join(question_lines))

    # Options
    options = {}
    options_images = {}
    opt_map = {'a': 'A', 'b': 'B', 'c': 'C', 'd': 'D', 'अ': 'A', 'ब': 'B', 'स': 'C', 'द': 'D', 'ए': 'A', 'बी': 'B', 'सी': 'C', 'डी': 'D', 'क': 'A', 'ख': 'B', 'ग': 'C', 'घ': 'D'}
    
    while i < len(lines) and not is_ans(lines[i]):
        line = lines[i]
        # Robust marker: Check if line starts with marker or has marker after substantial space
        # And ensure we don't match inside IMG tags by checking the left context
        matches = []
        # Find all potential matches including Hindi markers like (क) or (ए) or (अ)
        pattern = r'(?:\(([a-dA-Dअ-दए-डीक-घ])\)|\b([a-dA-D])\s*:)'
        for m in re.finditer(pattern, line, re.I):
            # Check if this match is inside an IMG tag
            pre = line[:m.start()]
            if pre.count(IMG_START) > pre.count(IMG_END):
                continue # Inside tag
            matches.append(m)
        
        if matches:
            for idx_m, m in enumerate(matches):
                letter = (m.group(1) or m.group(2)).lower()
                start_text = m.end()
                end_text = matches[idx_m+1].start() if idx_m+1 < len(matches) else len(line)
                text_opt = line[start_text:end_text].strip()
                field = extract_field(text_opt)
                options[opt_map[letter]] = field["text"]
                options_images[opt_map[letter]] = field["image"]
            i += 1
        elif ans_idx != -1 and ans_idx - 4 <= i < ans_idx:
            # Fallback for MS Word numbered options
            opt_keys = ['A', 'B', 'C', 'D']
            letter = opt_keys[len(options)] if len(options) < 4 else 'D'
            ext = extract_field(line.strip())
            if letter in options:
                options[letter] += " " + ext["text"]
            else:
                options[letter] = ext["text"]
            if ext["image"]: options_images[letter] = ext["image"]
            i += 1
        elif options and not is_stmt(line) and not is_pair(line) and not is_ans(line):
            last_l = sorted(options.keys())[-1]
            ext = extract_field(line)
            options[last_l] += " " + ext["text"]
            if ext["image"]: options_images[last_l] = ext["image"]
            i += 1
        else:
            break

    # Answer  (stop when we hit the Solution header)
    answer = ''
    while i < len(lines) and not is_sol(lines[i]):
        am = re.search(r'Answer\s*:?\s*\(?([a-dA-D])\)?', lines[i], re.I)
        if am:
            answer = am.group(1).lower()
        i += 1

    # Solution  – header may be "Solution:" OR bare "SOLUTION"
    solution_lines = []
    while i < len(lines):
        if is_sol(lines[i]):
            # Capture any text after the header word on the same line
            after = re.sub(r'^\s*Solution\s*:?\s*', '', lines[i], flags=re.I).strip()
            if after:
                solution_lines.append(after)
            i += 1
            while i < len(lines):
                solution_lines.append(lines[i])
                i += 1
            break
        i += 1
    solution_field = extract_field('\n'.join(solution_lines))

    return {
        "number": q_number,
        "subtopic": subtopic,
        "category": category if category != "aptitude" else "concept",
        "question_type": question_type,
        "question": question_field["text"],
        "question_image": question_field["image"],
        "statements": statements,
        "pairs": pairs,
        "lastQuestion": lastQuestion if isinstance(lastQuestion, dict) else {"text": lastQuestion, "image": ""},
        "options": options,
        "options_images": options_images,
        "answer": answer,
        "solution": solution_field["text"],
        "solution_image": solution_field["image"]
    }

def normalise(q: dict) -> dict | None:
    if not q["question"] and not q["question_image"]: return None
    if len(q["options"]) < 4: return None
    if not q["answer"]:
        print(f"Warning: Question {q.get('number', 'Unknown')} is missing an 'Answer:' line. Proceeding anyway.")
    return q

def parse_docx_file(file_path: str) -> list:
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    doc = docx.Document(file_path)
    img_dir = "data/images"
    os.makedirs(img_dir, exist_ok=True)
    img_map = {r_id: rel.target_part for r_id, rel in doc.part.rels.items() if 'image' in rel.reltype}

    def iter_block_items(parent):
        if isinstance(parent, docx.document.Document):
            parent_elm = parent.element.body
        else:
            parent_elm = parent._element

        for child in parent_elm.iterchildren():
            if isinstance(child, docx.oxml.text.paragraph.CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, docx.oxml.table.CT_Tbl):
                yield Table(child, parent)

    def get_runs_text(element):
        from docx.text.run import Run
        text = ""
        for r_node in element.xpath('.//w:r'):
            run = Run(r_node, None)
            if run.text:
                text += run.text
            xml = run._element.xml
            if '<w:drawing>' in xml:
                m = re.search(r'r:embed=\"(rId\d+)\"', xml)
                if m and m.group(1) in img_map:
                    img_part = img_map[m.group(1)]
                    fname = img_part.partname.split('/')[-1]
                    out_path = os.path.join(img_dir, fname)
                    with open(out_path, 'wb') as f: f.write(img_part.blob)
                    clean_p = os.path.abspath(out_path).replace('\\', '/')
                    text += f" {IMG_START}{clean_p}{IMG_END} "
        return text

    lines = []
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            # Handle Word numbering: If paragraph has numbering, prepend a marker
            p_text = get_runs_text(item._element)
            
            # Check for numbering in XML
            if item._element.xpath('.//w:numPr'):
                # Try to get ilvl (indent level)
                ilvls = item._element.xpath('.//w:ilvl/@w:val')
                lvl = int(ilvls[0]) if ilvls else 0
                # We'll just prepend a simple (1) for all levels to trigger detection
                p_text = f"(1) {p_text}"

            for sl in p_text.split('\n'):
                if sl.strip(): lines.append(sl.strip())
        else: # Table
            for row in item.rows:
                row_text = [get_runs_text(cell._element).strip() for cell in row.cells]
                row_text = [t for t in row_text if t]
                if len(row_text) > 1:
                    lines.append("=".join(row_text))
                elif len(row_text) == 1:
                    lines.append(row_text[0])

    # Clean up accidental "ANSWER:" replacements from Hindi context
    for i_line in range(len(lines)):
        lines[i_line] = re.sub(
            r'ANSWER:\s*([-\u0900-\u097F]+)',
            lambda m: ('उत्तर' + m.group(1)) if m.group(1).startswith('-') or m.group(1) == 'काशी' else ('उत्तर ' + m.group(1)),
            lines[i_line]
        )

    # Hybrid block detection
    block_starts = []
    
    # 1. Triplet-based detection
    # Use EXACT line matching so question text that merely contains "concept" or
    # "statement" does not falsely fire the detector.
    triplet_starts = set()

    # 'Satement' (missing second 't') is a known typo in the English source doc
    type_exact    = {"normal", "statement", "satement", "pair",
                     "normal ", "statement ", "satement ", "pair ",
                     "सामान्य", "कथन", "जोड़ी"}
    concept_exact = {"concept", "concept ", "कॉन्सेप्ट", "कॉन्सेप्ट "}

    for j in range(1, len(lines)):
        l_j      = lines[j].lower().strip()
        l_j_prev = lines[j-1].lower().strip()

        is_type    = l_j      in type_exact
        is_concept = l_j_prev in concept_exact

        if is_type and is_concept:
            triplet_starts.add(max(0, j - 2))
    
    # 2. Sequential/Contextual detection
    # Any line matching the question marker that follows a solution block
    q_marker_re = r'(?:' + re.escape(IMG_START) + r'.*?' + re.escape(IMG_END) + r'\s*)?(?:\((\d+)\)|(\d+)\s*[\.\)\:])'
    
    last_q_val = 0
    in_solution = False
    
    for i, line in enumerate(lines):
        # Triplet always wins
        if i in triplet_starts:
            block_starts.append(i)
            # Find the number in this block to keep sequence
            for k in range(i, min(i+10, len(lines))):
                m = re.match(q_marker_re, lines[k])
                if m and m.group(1):
                    last_q_val = int(m.group(1))
                    break
            in_solution = False
            continue
            
        # Check for solution marker
        if re.search(r'\banswer\b|\bsolution\b|विकल्प|सही है', line, re.I):
            in_solution = True
            
        # Check for question number fallback
        m = re.match(q_marker_re, line)
        if m:
            q_val_str = m.group(1) or m.group(2)
            if q_val_str:
                q_val = int(q_val_str)
                # It's a new question if it's the next in sequence or follows a solution
                if (q_val == last_q_val + 1) or (in_solution and q_val > last_q_val):
                    # Only add if not already captured by a nearby triplet
                    is_new = True
                    for s in block_starts:
                        if abs(s - i) < 10:
                            is_new = False
                            break
                    if is_new:
                        # Skip q_val == 1 unless this looks like a genuine Q1 line
                        # (not a numPr-prefixed solution statement like "(1) Foo...").
                        # Genuine question markers are at the start of a block and
                        # typically preceded by blank/meta lines, whereas solution
                        # (1) lines live deep inside a block.
                        if q_val == 1:
                            # Only accept if this is the very first line in a new block
                            # i.e. nothing was detected yet, or check prev line context
                            prev_line = lines[i-1].strip() if i > 0 else ""
                            looks_like_new = not re.search(
                                r'(Statement|Correct|Incorrect|सही|गलत|\bAnswer\b|\bSolution\b|व्याख्या)',
                                prev_line, re.I)
                            if not looks_like_new:
                                continue
                        block_starts.append(i)
                        last_q_val = q_val
                        in_solution = False

    # Filter and refine: each block must contain a question number
    final_block_starts = sorted(list(set(block_starts)))
    
    questions = []
    for idx in range(len(final_block_starts)):
        start = final_block_starts[idx]
        end = final_block_starts[idx+1] if idx+1 < len(final_block_starts) else len(lines)
        q = parse_block_regex('\n'.join(lines[start:end]))
        if q:
            nq = normalise(q)
            if nq: questions.append(nq)
            
    print(f"Parsed {len(questions)} questions from {os.path.basename(file_path)}")
    return questions

def merge_questions(en_qs: list, hi_qs: list) -> list:
    merged = []
    # Map by question number for stable alignment
    en_map = {q["number"]: q for q in en_qs}
    hi_map = {q["number"]: q for q in hi_qs}
    
    # Get all unique numbers and sort them
    all_nums = sorted(list(set(en_map.keys()) | set(hi_map.keys())), key=lambda x: int(x) if x.isdigit() else 999)
    
    for num in all_nums:
        en_q = en_map.get(num)
        hi_q = hi_map.get(num)
        
        # Definitive metadata from English if available, else Hindi
        ref = en_q or hi_q
        subtopic = ref.get("subtopic", "Unknown")
        category = ref.get("category", "concept")
        if category == "aptitude": category = "concept"
        q_type = ref.get("question_type", "normal")

        def build_side(q_data, definitive_meta):
            if not q_data:
                return {
                    "number": num,
                    "subtopic": definitive_meta["subtopic"],
                    "category": definitive_meta["category"],
                    "question_type": definitive_meta["type"],
                    "question": "", "question_image": "",
                    "statements": [], "pairs": [], "lastQuestion": {"text": "", "image": ""},
                    "options": {"A":"","B":"","C":"","D":""},
                    "options_images": {"A":"","B":"","C":"","D":""},
                    "answer": "", "solution": "", "solution_image": ""
                }
            # Sync metadata
            q_data["subtopic"] = definitive_meta["subtopic"]
            q_data["category"] = definitive_meta["category"]
            q_data["question_type"] = definitive_meta["type"]
            return q_data

        meta_bundle = {"subtopic": subtopic, "category": category, "type": q_type}
        
        en_side = build_side(en_q, meta_bundle)
        if q_type == "pair" and not en_side["lastQuestion"]["text"].strip():
            en_side["lastQuestion"]["text"] = "How many of the above pairs are correctly matched?"
        elif q_type == "statement" and not en_side["lastQuestion"]["text"].strip():
            en_side["lastQuestion"]["text"] = "Which of the statements given above is/are correct?"

        hi_side = build_side(hi_q, meta_bundle)
        if q_type == "pair" and not hi_side["lastQuestion"]["text"].strip():
            hi_side["lastQuestion"]["text"] = "ऊपर दिए गए जोड़ों में से कितने सही मैच हुए हैं?"
        elif q_type == "statement" and not hi_side["lastQuestion"]["text"].strip():
            hi_side["lastQuestion"]["text"] = "ऊपर दिए गए कथनों में से कौन सा/से सही है/हैं?"

        merged.append({
            "number": num,
            "subtopic": subtopic,
            "category": category,
            "question_type": q_type,
            "english": en_side,
            "hindi": hi_side
        })
    return merged

def main():
    # Ensure the script runs from the project root
    script_path = os.path.abspath(__file__)
    project_root = os.path.dirname(os.path.dirname(script_path))
    os.chdir(project_root)
    
    os.makedirs("data", exist_ok=True)
    en = parse_docx_file("data/english.docx")
    hi = parse_docx_file("data/hindi.docx")
    merged = merge_questions(en, hi)
    with open("data/parsed_questions.json", "w", encoding="utf-8") as f:
        json.dump(merged, f, indent=2, ensure_ascii=False)
    print(f"DONE: Parsed {len(merged)} questions.")

if __name__ == "__main__":
    main()
