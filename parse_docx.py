"""
parse_docx.py - V3 Robust
"""
import docx
import json
import re
import os
import time
from deep_translator import GoogleTranslator

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

def extract_all_images(text: str) -> list:
    """Return all image paths found in the text, in document order."""
    pattern = re.escape(IMG_START) + r"\s*(.*?)\s*" + re.escape(IMG_END)
    return [m.group(1).strip() for m in re.finditer(pattern, text, re.I)]

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
    
    # Fallback: If no explicit marker found, but we have metadata (at least 3 lines),
    # the 4th line is likely the question start.
    if q_idx == -1 and len(lines) >= 3:
        q_idx = min(3, len(lines) - 1)
    
    if q_idx == -1: return None
    
    meta = lines[:q_idx]
    subtopic = meta[0] if len(meta) > 0 else "Unknown"
    category = meta[1].lower() if len(meta) > 1 else "aptitude"
    
    # Normalize question_type: lower case, " - " to "-", spaces to hyphens
    # Handle all variants of dashes (hyphen, en-dash, em-dash) and non-breaking spaces (\u00A0)
    raw_type = meta[2].lower().strip() if len(meta) > 2 else "normal"
    question_type = re.sub(r'[\s\u00A0]*[–—\-][\s\u00A0]*', '-', raw_type)
    question_type = re.sub(r'[\s\u00A0]+', '-', question_type)

    # --- CSAT FIX: Capture any image lines that sit between the type-metadata
    # line (q_idx - 1) and the actual question-number marker line (q_idx).
    # These images ARE the question body for Normal-csat questions.
    pre_question_image = ""
    pre_question_image_hindi = ""
    if len(meta) >= 3:
        # Collect ALL image tokens from the metadata block (between type and question number).
        # First image  = English question image.
        # Second image = Hindi question image (for CSAT image-only questions).
        img_pattern = re.escape(IMG_START) + r'\s*(.*?)\s*' + re.escape(IMG_END)
        for ml in meta:
            for m_img in re.finditer(img_pattern, ml, re.I):
                if not pre_question_image:
                    pre_question_image = m_img.group(1).strip()
                elif not pre_question_image_hindi:
                    pre_question_image_hindi = m_img.group(1).strip()

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

    is_options = lambda s: bool(re.match(r'^(?:\([a-dA-Dअ-दए-डीक-घ]\)|[A-D]\s*[:\.\)])', s, re.I))
    is_stmt = lambda s: IMG_START not in s and bool(re.match(r'^\s*(\d+\s*[\.\)]|\(\d+\))', s, re.I))
    # Ignore image markers in pair detection to avoid matching file paths
    is_pair = lambda s: IMG_START not in s and bool(re.match(r'^(?:Pair\s*\d+\s*:|.*?\s*(?:=|—|–|-)\s*.*)', s, re.I))
    is_img = lambda s: IMG_START in s
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
        if re.match(r'^(?:\([a-dA-Dअ-दए-डीक-घ]\)|[A-D]\s*[:\.\)])', s, re.I): return True
        if ans_idx != -1 and ans_idx - 4 <= idx < ans_idx: return True
        return False

    is_options = lambda idx, s: is_options_func(idx, s)
    is_options_compat = lambda s: bool(re.match(r'^(?:\([a-dA-Dअ-दए-डीक-घ]\)|[A-D]\s*[:\.\)])', s, re.I))

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
        current_stmt_lines = []
        stmt_raw_texts = []
        
        while i < len(lines) and not is_options(i, lines[i]) and not is_ans(lines[i]) and not is_sol(lines[i]) and not re.search(lq_trigger, lines[i], re.I):
            line = lines[i].strip()
            if not line:
                i += 1
                continue
                
            temp_line = re.sub(r'^' + re.escape(IMG_START) + r'.*?' + re.escape(IMG_END) + r'\s*', '', line).strip()
            is_new_stmt_marker = bool(re.match(r'^(\d+|[A-Za-z])[\.\)]|^\(\d+\)|^\([A-Za-z]\)', temp_line))
            
            if is_new_stmt_marker:
                if current_stmt_lines:
                    stmt_raw_texts.append('\n'.join(current_stmt_lines))
                current_stmt_lines = []
                
                # Strip the marker to get the text of the line
                cleaned_line = re.sub(r'^(\d+|[A-Za-z])[\.\)]\s*|^\(\d+\)\s*|^\([A-Za-z]\)\s*', '', temp_line).strip()
                
                # Restore leading image if present
                m_img = re.match(r'^(' + re.escape(IMG_START) + r'.*?' + re.escape(IMG_END) + r'\s*)', line)
                if m_img:
                    current_stmt_lines.append(m_img.group(1).strip())
                if cleaned_line:
                    current_stmt_lines.append(cleaned_line)
            else:
                current_stmt_lines.append(line)
            i += 1
            
        if current_stmt_lines:
            stmt_raw_texts.append('\n'.join(current_stmt_lines))
            
        for raw_text in stmt_raw_texts:
            f = extract_field(raw_text)
            all_imgs = extract_all_images(raw_text)
            hi_img = all_imgs[1] if len(all_imgs) > 1 else ""
            statements.append({
                "text": f["text"],
                "image": f["image"],
                "image_hindi": hi_img
            })
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
                    f = extract_field(clean_p)
                    all_imgs = extract_all_images(clean_p)
                    hi_img = all_imgs[1] if len(all_imgs) > 1 else ""
                    row.append({
                        "text": f["text"],
                        "image": f["image"],
                        "image_hindi": hi_img
                    })
                if len(row) >= 2: pairs.append(row)
            i += 1
        lq_parts = []
        while i < len(lines) and not is_options(i, lines[i]) and not is_ans(lines[i]):
            lq_parts.append(lines[i])
            i += 1
        lastQuestion = extract_field(' '.join(lq_parts).strip())
    else:
        # Normal type (including normal-csat): question text already collected until options.
        # For normal-csat the question body is typically an image; nothing extra to do here.
        pass

    q_all_text = '\n'.join(question_lines)
    question_field = extract_field(q_all_text)
    q_all_images = extract_all_images(q_all_text)

    # CSAT FIX: Images in the metadata block (before the question number) are the
    # actual question body.  Promote them to question_image / question_image_hindi.
    if pre_question_image and not question_field.get("image", ""):
        question_field["image"] = pre_question_image
        # Second metadata image → Hindi question image
        question_image_hindi = pre_question_image_hindi
    else:
        # If images were inline in the question text, second one is the Hindi version
        question_image_hindi = q_all_images[1] if len(q_all_images) > 1 else ""

    # Options
    options = {}
    options_images = {}
    opt_map = {'a': 'A', 'b': 'B', 'c': 'C', 'd': 'D', 'अ': 'A', 'ब': 'B', 'स': 'C', 'द': 'D', 'ए': 'A', 'बी': 'B', 'सी': 'C', 'डी': 'D', 'क': 'A', 'ख': 'B', 'ग': 'C', 'घ': 'D'}
    
    has_markers = False
    while i < len(lines) and not is_ans(lines[i]):
        line = lines[i]
        # Robust marker: Check if line starts with marker or has marker after substantial space
        # And ensure we don't match inside IMG tags by checking the left context
        matches = []
        # Find all potential matches including Hindi markers like (क) or (ए) or (अ)
        pattern = r'(?:\(([a-dA-Dअ-दए-डीक-घ])\)|\b([a-dA-Dअ-दए-डीक-घ])\s*[:\.])'
        for m in re.finditer(pattern, line, re.I):
            # Check if this match is inside an IMG tag
            pre = line[:m.start()]
            if pre.count(IMG_START) > pre.count(IMG_END):
                continue # Inside tag
            matches.append(m)
        
        if matches:
            has_markers = True
            for idx_m, m in enumerate(matches):
                letter = (m.group(1) or m.group(2)).lower()
                start_text = m.end()
                end_text = matches[idx_m+1].start() if idx_m+1 < len(matches) else len(line)
                text_opt = line[start_text:end_text].strip()
                field = extract_field(text_opt)
                options[opt_map[letter]] = field["text"]
                if field["image"]:
                    options_images[opt_map[letter]] = field["image"]
            i += 1
        elif not has_markers and ans_idx != -1 and ans_idx - 4 <= i < ans_idx:
            # Fallback for MS Word numbered options
            if re.search(re.escape(IMG_START), line):
                pass
            else:
                opt_keys = ['A', 'B', 'C', 'D']
                letter = opt_keys[len(options)] if len(options) < 4 else 'D'
                ext = extract_field(line.strip())
                if letter in options:
                    options[letter] += " " + ext["text"]
                else:
                    options[letter] = ext["text"]
                if ext["image"]: options_images[letter] = ext["image"]
                i += 1
                continue
        elif options and (is_img(line) or (not is_stmt(line) and not is_pair(line) and not is_ans(line))):
            last_l = sorted(options.keys())[-1]
            ext = extract_field(line)
            options[last_l] += " " + ext["text"]
            if ext["image"]:
                if not options_images.get(last_l):
                    options_images[last_l] = ext["image"]
                elif not options_images.get(last_l + "_hindi"):
                    options_images[last_l + "_hindi"] = ext["image"]
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
    sol_all_text = '\n'.join(solution_lines)
    solution_field = extract_field(sol_all_text)
    sol_all_images = extract_all_images(sol_all_text)
    # Second solution image (if present) is the Hindi version of the solution image
    solution_image_hindi = sol_all_images[1] if len(sol_all_images) > 1 else ""

    return {
        "number": q_number,
        "subtopic": subtopic,
        "category": category if category != "aptitude" else "concept",
        "question_type": question_type,
        "question": question_field["text"],
        "question_image": question_field["image"],
        "question_image_hindi": question_image_hindi,
        "statements": statements,
        "pairs": pairs,
        "lastQuestion": lastQuestion if isinstance(lastQuestion, dict) else {"text": lastQuestion, "image": ""},
        "options": options,
        "options_images": options_images,
        "answer": answer,
        "solution": solution_field["text"],
        "solution_image": solution_field["image"],
        "solution_image_hindi": solution_image_hindi
    }

def normalise(q: dict) -> dict | None:
    q_type = q["question_type"].lower()
    is_csat = "csat" in q_type

    # 1. Basic requirements for all types
    # For CSAT: question can be image-only (text may be empty if image is present)
    if not q["question"].strip() and not q.get("question_image", "").strip():
        return None

    if len(q["options"]) < 4: return None
    # Ensure all 4 options have text OR an image
    for k, val in q["options"].items():
        has_text = bool(str(val).strip())
        has_image = bool(q["options_images"].get(k, "").strip())
        if not has_text and not has_image:
            return None

    if not q["answer"].strip(): return None

    # For CSAT: solution can be image-only (text may be empty if image is present)
    if not q["solution"].strip() and not q.get("solution_image", "").strip():
        return None

    # 2. Type-specific requirements
    if "statement" in q_type and not is_csat:
        # Pure statement type: must have statements with text
        if not q["statements"]: return None
        if not all(s.get("text", "").strip() for s in q["statements"]): return None

        lq = q.get("lastQuestion", {})
        if not lq.get("text", "").strip():
            q["lastQuestion"] = {"text": "Which of the statements is correct?", "image": ""}

    elif "statement" in q_type and is_csat:
        # Statement-csat: question body is an image; statements list may be empty
        lq = q.get("lastQuestion", {})
        if not lq.get("text", "").strip() and not lq.get("image", "").strip():
            q["lastQuestion"] = {"text": "", "image": ""}

    elif "pair" in q_type and not is_csat:
        # Pure pair type: must have pairs with text in both columns
        if not q["pairs"]: return None
        for row in q["pairs"]:
            if not all(cell.get("text", "").strip() for cell in row): return None

        lq = q.get("lastQuestion", {})
        if not lq.get("text", "").strip():
            q["lastQuestion"] = {"text": "How many pairs are correctly matched?", "image": ""}

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
            # Check for numbering in XML or Style
            has_num = False
            try:
                if item._element.xpath('.//w:numPr'):
                    has_num = True
            except:
                pass
            
            if not has_num:
                style_name = item.style.name.lower()
                if "list" in style_name or "number" in style_name:
                    has_num = True

            if has_num:
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
    
    # Pre-process lines to remove hidden Unicode whitespace/artifacts
    clean_lines = []
    for l in lines:
        # Replace non-breaking spaces and other artifacts with standard space
        cl = l.replace('\u00A0', ' ').replace('\u200b', '').strip()
        clean_lines.append(cl)

    # 1. Fuzzy Triplet-based detection
    triplet_starts = set()

    # Normalize labels for ultra-fuzzy matching (ignores dashes and spaces)
    def norm_label(s):
        s = s.lower().strip()
        # Remove all dashes, spaces, and non-alphanumeric characters for comparison
        # (keeping Hindi characters)
        s = re.sub(r'[^a-z0-9\u0900-\u097F]', '', s)
        return s

    # These are normalized keys (no dashes, no spaces)
    type_exact = {"normal", "statement", "pair", "normalcsat", "statementcsat", "paircsat", "सामान्य", "कथन", "जोड़ी", "सामान्यcsat", "कथनcsat", "जोड़ीcsat"}
    concept_exact = {"concept", "aptitude", "कॉन्सेप्ट"}

    for j in range(len(clean_lines)):
        l_j = norm_label(clean_lines[j])
        
        # If this line looks like a question type
        if any(t in l_j for t in type_exact):
            # Look back up to 3 lines for a concept marker
            found_concept = False
            for look_back in range(1, 4):
                if j - look_back >= 0:
                    prev_l = norm_label(clean_lines[j - look_back])
                    if any(c in prev_l for c in concept_exact):
                        # Found a triplet! Block starts at the line before concept (subtopic)
                        triplet_starts.add(max(0, j - look_back - 1))
                        found_concept = True
                        break
            
            # If it's a CSAT type, we are even more aggressive
            if "csat" in l_j and not found_concept:
                triplet_starts.add(max(0, j - 2)) # Assume subtopic/concept are there
    
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
                if m:
                    val = m.group(1) or m.group(2)
                    if val:
                        last_q_val = int(val)
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
                                r'(Statement|Correct|Incorrect|सही|गलत|\bAnswer\b|\bSolution\b|व्याख्या|Fundamental|Rights)',
                                prev_line, re.I)
                            if not looks_like_new:
                                continue
                        block_starts.append(i)
                        last_q_val = q_val
                        in_solution = False

    # Filter and refine: each block must contain a question number
    final_block_starts = sorted(list(set(block_starts)))
    
    questions = []
    failed_blocks = []
    
    for idx in range(len(final_block_starts)):
        start = final_block_starts[idx]
        end = final_block_starts[idx+1] if idx+1 < len(final_block_starts) else len(lines)
        block_text = '\n'.join(lines[start:end])
        
        q = parse_block_regex(block_text)
        if q:
            nq = normalise(q)
            if nq:
                nq["status"] = "success"
                questions.append(nq)
            else:
                failed_blocks.append({
                    "reason": "Normalization failed (e.g., less than 4 options)",
                    "text": block_text,
                    "q_number": q.get("number", "Unknown"),
                    "options_count": len(q["options"]),
                    "options": q["options"]
                })
        else:
            failed_blocks.append({
                "reason": "Regex parsing failed",
                "text": block_text
            })
            
    print(f"Parsed {len(questions)} questions, Failed {len(failed_blocks)} blocks.")
    return questions, failed_blocks

def translate_to_hindi(text: str) -> str:
    if not text or not str(text).strip():
        return ""
    try:
        # Using GoogleTranslator for reliable free translation
        return GoogleTranslator(source='en', target='hi').translate(text)
    except Exception as e:
        print(f"Translation Error: {e}")
        return ""

def merge_questions(en_qs: list, parsing_failures: list = None, output_file: str = None, incremental_save: bool = False) -> tuple[list, list]:
    merged = []
    failed_translations = []
    total = len(en_qs)
    
    parsing_failures = parsing_failures or []
    
    for idx, en_q in enumerate(en_qs):
        num = en_q["number"]
        print(f"[{idx+1}/{total}] Translating Question {num}...")
        
        # Build Hindi side by translating English fields.
        # Rule:
        #   - If text exists  → translate via deep_translator (images same for both sides)
        #   - If no text, image-only → English side gets image 1, Hindi side gets image 2
        has_q_text  = bool(en_q["question"].strip())
        has_sol_text = bool(en_q["solution"].strip())

        # Question image routing
        q_img_hindi = en_q.get("question_image_hindi", "").strip()
        hi_question_image = q_img_hindi if (not has_q_text and q_img_hindi) else en_q["question_image"]

        # Solution image routing
        sol_img_hindi = en_q.get("solution_image_hindi", "").strip()
        hi_solution_image = sol_img_hindi if (not has_sol_text and sol_img_hindi) else en_q["solution_image"]

        # Prepare for options routing
        hi_options_images = {}
        # Pre-calculate options translation to get correct routing
        translated_options = {}
        for k, v in en_q["options"].items():
            translated_options[k] = translate_to_hindi(v)
            
            # Option image routing
            opt_img_en = en_q["options_images"].get(k, "")
            opt_img_hi = en_q["options_images"].get(k + "_hindi", "")
            
            # If no text, use Hindi image for Hindi side if available
            has_opt_text = bool(v.strip())
            hi_options_images[k] = opt_img_hi if (not has_opt_text and opt_img_hi) else opt_img_en

        hi_side = {
            "number": num,
            "subtopic": en_q["subtopic"],
            "category": en_q["category"],
            "question_type": en_q["question_type"],
            "question": translate_to_hindi(en_q["question"]),
            "question_image": hi_question_image,
            "statements": [],
            "pairs": [],
            "lastQuestion": {"text": "", "image": ""},
            "options": translated_options,
            "options_images": hi_options_images,
            "answer": en_q["answer"],
            "solution": translate_to_hindi(en_q["solution"]),
            "solution_image": hi_solution_image
        }

        # Translate Statements
        for stat in en_q["statements"]:
            stat_has_text = bool(stat.get("text", "").strip())
            stat_img_hi = stat.get("image_hindi", "").strip()
            hi_stat_img = stat_img_hi if (not stat_has_text and stat_img_hi) else stat.get("image", "")

            hi_side["statements"].append({
                "text": translate_to_hindi(stat.get("text", "")),
                "image": hi_stat_img
            })

        # Translate Pairs
        for row in en_q["pairs"]:
            hi_row = []
            for p_item in row:
                p_has_text = bool(p_item.get("text", "").strip())
                p_img_hi = p_item.get("image_hindi", "").strip()
                hi_p_img = p_img_hi if (not p_has_text and p_img_hi) else p_item.get("image", "")

                hi_row.append({
                    "text": translate_to_hindi(p_item.get("text", "")),
                    "image": hi_p_img
                })
            hi_side["pairs"].append(hi_row)

        # Translate Last Question
        if en_q["lastQuestion"]["text"]:
            hi_side["lastQuestion"]["text"] = translate_to_hindi(en_q["lastQuestion"]["text"])
            hi_side["lastQuestion"]["image"] = en_q["lastQuestion"]["image"]

        # Mark as partial if translation failed.
        # For image-only questions: no translated text is OK as long as we have an image.
        q_ok  = bool(hi_side["question"]) or bool(hi_side["question_image"])
        opts_ok = all(bool(opt) for opt in hi_side["options"].values()) or any(bool(img) for img in hi_side["options_images"].values())
        if not q_ok or not opts_ok:
            en_q["status"] = "partial_translation"
            failed_translations.append({
                "q_number": num,
                "reason": "One or more fields failed to translate"
            })

        # Strip internal routing-only fields before storing in the final output.
        # These were only used to route the correct image to the Hindi side.
        en_q.pop("question_image_hindi", None)
        en_q.pop("solution_image_hindi", None)
        for stat in en_q.get("statements", []):
            stat.pop("image_hindi", None)
        for row in en_q.get("pairs", []):
            for p_item in row:
                p_item.pop("image_hindi", None)
        
        # Cleanup options_images _hindi keys
        for k in list(en_q.get("options_images", {}).keys()):
            if k.endswith("_hindi"):
                en_q["options_images"].pop(k)

        merged.append({
            "number": num,
            "subtopic": en_q["subtopic"],
            "category": en_q["category"],
            "question_type": en_q["question_type"],
            "english": en_q,
            "hindi": hi_side
        })
        
        # Save incrementally if requested
        if incremental_save and output_file:
            with open(output_file, "w", encoding="utf-8") as f:
                json.dump({
                    "processed": merged, 
                    "parsing_failures": parsing_failures,
                    "failed_translations": failed_translations
                }, f, indent=2, ensure_ascii=False)
            
        # Small sleep to avoid rate limiting
        time.sleep(0.5)
            
    return merged, failed_translations

def process_document(file_path: str, output_file: str = "parsed_questions.json") -> dict:
    """Helper to parse and translate a document in one go."""
    en_qs, parsing_failures = parse_docx_file(file_path)
    
    # For API, we want to return both successful and failed ones
    merged, translation_failures = merge_questions(en_qs, parsing_failures=parsing_failures, output_file=output_file, incremental_save=True)
    
    return {
        "processed": merged,
        "parsing_failures": parsing_failures,
        "translation_failures": translation_failures
    }

def main():
    # Ensure the script runs from the project root
    script_path = os.path.abspath(__file__)
    project_root = os.path.dirname(script_path)
    os.chdir(project_root)
    
    os.makedirs("data", exist_ok=True)
    process_document("english.docx", "parsed_questions.json")
    print(f"DONE: Document processed.")

if __name__ == "__main__":
    main()
