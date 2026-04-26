from docx import Document
from docx.oxml.ns import qn
import os
import re
import zipfile
import random
import string
import xml.sax.saxutils as saxutils
import lxml.etree as ET

def parse_docx_to_data(input_docx, job_dir):
    """
    Parses a DOCX file and extracts questions, options, and answers.
    It combines logic for both standard text paragraphs and table-based layouts.
    Returns a dictionary suitable for JSON serialization (Preview).
    """
    doc = Document(input_docx)

    assets_dir = os.path.join(job_dir, "assets")
    os.makedirs(assets_dir, exist_ok=True)

    image_counter = 0

    def save_image(doc, rId):
        nonlocal image_counter
        image_part = doc.part.related_parts[rId]
        filename = f"img_{image_counter}.jpg"
        with open(os.path.join(assets_dir, filename), "wb") as f:
            f.write(image_part.blob)
        image_counter += 1
        return filename

    # Regex Patterns
    question_regex = re.compile(r'^(\d+)(?:\s+(.*))?$', re.I)
    option_regex = re.compile(r'^[A-D][\.\)]\s*(.*)')

    list_counters = {}
    
    def get_roman(num):
        val = [10, 9, 5, 4, 1]
        syb = ["x", "ix", "v", "iv", "i"]
        roman_num = ''
        i = 0
        while num > 0:
            for _ in range(num // val[i]):
                roman_num += syb[i]
                num -= val[i]
            i += 1
        return roman_num

    def parse_paragraph(child):
        tokens = []
        for run in child.xpath("./w:r"):
            text = "".join(node.text or "" for node in run.xpath(".//w:t"))
            if not text:
                continue
            is_sup = run.xpath(".//w:vertAlign[@w:val='superscript']")
            is_sub = run.xpath(".//w:vertAlign[@w:val='subscript']")
            is_bold = run.xpath(".//w:b")
            is_italic = run.xpath(".//w:i")
            if is_sup: text = f"<sup>{text}</sup>"
            if is_sub: text = f"<sub>{text}</sub>"
            if is_bold: text = f"<strong>{text}</strong>"
            if is_italic: text = f"<em>{text}</em>"
            tokens.append(text)
            
        paragraph_text = "".join(tokens).strip()

        alignment = child.xpath(".//w:jc/@w:val")
        align_val = alignment[0] if alignment else None

        if paragraph_text:
            yield ["text", paragraph_text, align_val]

        for node in child.iter():
            if node.tag.endswith('blip'):
                rId = node.get(qn('r:embed'))
                if rId:
                    yield ["image", rId, align_val]

    def parse_table(child):
        table_data = []
        for row in child.xpath("./w:tr"):
            row_data = []
            for cell in row.xpath("./w:tc"):
                cell_text = "".join(node.text or "" for node in cell.xpath(".//w:t")).strip()
                row_data.append(cell_text)
            table_data.append(row_data)
        return table_data

    def extract_numbering_map(docx_path):
        num_map = {}
        try:
            with zipfile.ZipFile(docx_path, 'r') as z:
                if "word/numbering.xml" not in z.namelist():
                    return num_map
                num_xml = z.read("word/numbering.xml")
            root = ET.fromstring(num_xml)
            namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            abstract_map = {}
            for abstractNum in root.xpath('//w:abstractNum', namespaces=namespaces):
                abs_id = abstractNum.get('{' + namespaces['w'] + '}abstractNumId')
                levels = {}
                for lvl in abstractNum.xpath('./w:lvl', namespaces=namespaces):
                    ilvl = lvl.get('{' + namespaces['w'] + '}ilvl')
                    numFmt = lvl.xpath('./w:numFmt', namespaces=namespaces)
                    fmt_val = numFmt[0].get('{' + namespaces['w'] + '}val') if numFmt else "decimal"
                    lvlText = lvl.xpath('./w:lvlText', namespaces=namespaces)
                    txt_val = lvlText[0].get('{' + namespaces['w'] + '}val') if lvlText else "%1."
                    levels[ilvl] = (fmt_val, txt_val)
                abstract_map[abs_id] = levels
                
            for num in root.xpath('//w:num', namespaces=namespaces):
                num_id = num.get('{' + namespaces['w'] + '}numId')
                abs_num_id = num.xpath('./w:abstractNumId', namespaces=namespaces)
                if abs_num_id:
                    abs_val = abs_num_id[0].get('{' + namespaces['w'] + '}val')
                    if abs_val in abstract_map:
                        num_map[num_id] = abstract_map[abs_val]
        except Exception:
            pass
        return num_map

    def format_number(counter, fmt_val, txt_val):
        valStr = str(counter)
        if fmt_val == "lowerLetter":
            valStr = chr(ord('a') + (counter - 1) % 26)
        elif fmt_val == "upperLetter":
            valStr = chr(ord('A') + (counter - 1) % 26)
        elif fmt_val == "lowerRoman":
            valStr = get_roman(counter).lower()
        elif fmt_val == "upperRoman":
            valStr = get_roman(counter).upper()
        return re.sub(r'%\d', valStr, txt_val)

    def iter_block_items(doc):
        body = doc.element.body
        
        num_map = extract_numbering_map(input_docx)
        
        def process_paragraph_tokens(child):
            tokens = list(parse_paragraph(child))
            numPr = child.xpath(".//w:numPr")
            if numPr:
                numId = numPr[0].xpath(".//w:numId")
                ilvl = numPr[0].xpath(".//w:ilvl")
                if numId and ilvl:
                    numId_val = numId[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
                    ilvl_val = ilvl[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
                    if numId_val and ilvl_val:
                        list_key = f"{numId_val}_{ilvl_val}"
                        list_counters[list_key] = list_counters.get(list_key, 0) + 1
                        counter = list_counters[list_key]
                        
                        if numId_val in num_map and ilvl_val in num_map[numId_val]:
                            fmt_val, txt_val = num_map[numId_val][ilvl_val]
                            formatted_num = format_number(counter, fmt_val, txt_val)
                        else:
                            roman = get_roman(counter).lower()
                            formatted_num = f"{roman}."
                            
                        if tokens and tokens[0][0] == "text":
                            tokens[0][1] = f"{formatted_num} " + tokens[0][1]
                        elif tokens:
                            tokens.insert(0, ["text", f"{formatted_num} ", tokens[0][2] if len(tokens[0])>2 else None])
                        else:
                            tokens.append(["text", formatted_num, None])
            return tokens

        for child in body:
            if child.tag.endswith('p'):
                for t in process_paragraph_tokens(child):
                    yield t
            elif child.tag.endswith('tbl'):
                # Detect if table is used for formatting/layout (e.g., column 1 is question number)
                is_layout = False
                for row in child.xpath("./w:tr"):
                    cells = row.xpath("./w:tc")
                    if len(cells) >= 2:
                        first_cell_text = "".join(node.text or "" for node in cells[0].xpath(".//w:t")).strip()
                        if re.match(r'^\d+[a-zA-Z]*[ivxlcdm]*\.?$', first_cell_text, re.I):
                            is_layout = True
                            break
                
                if is_layout:
                    # Unwrap the layout table into linear blocks so the standard parser can process it
                    current_prefix = ""
                    for row in child.xpath("./w:tr"):
                        cells = row.xpath("./w:tc")
                        if len(cells) >= 2:
                            first_cell_text = "".join(node.text or "" for node in cells[0].xpath(".//w:t")).strip()
                            
                            if first_cell_text:
                                digit_match = re.match(r'^(\d+)', first_cell_text)
                                if digit_match:
                                    current_prefix = digit_match.group(1)
                                    letter_part = first_cell_text[len(current_prefix):].strip().rstrip('.')
                                    
                                    yield ["text", current_prefix, None]
                                    
                                    if letter_part:
                                        yield ["text", f"({letter_part})", None]
                                        
                                elif current_prefix and re.match(r'^[a-zA-Z]+[ivxlcdm]*\.?$', first_cell_text, re.I):
                                    letter_part = first_cell_text.strip().rstrip('.')
                                    yield ["text", f"({letter_part})", None]
                                else:
                                    # if it doesn't match normal layouts but has text, just yield it
                                    yield ["text", first_cell_text, None]

                            # Clear counters for the new cell so numbering restarts at i.
                            list_counters.clear()
                            
                            # Yield contents of the second cell sequentially
                            for cell_child in cells[1].iterchildren():
                                if cell_child.tag.endswith('p'):
                                    for t in process_paragraph_tokens(cell_child):
                                        yield t
                                elif cell_child.tag.endswith('tbl'):
                                    yield ["table", parse_table(cell_child)]
                else:
                    yield ["table", parse_table(child)]

    mcq_questions = []
    structured_questions = []
    answers = {}
    
    current_qnum = None
    current_tokens = []
    options = []
    
    reading_answers = False
    current_answer_q = None
    current_answer_tokens = []

    for item in iter_block_items(doc):
        item_type = item[0]
        value = item[1]
        align_val = item[2] if len(item) > 2 else None

        if reading_answers:
            if item_type == "image" and current_answer_q is not None:
                filename = save_image(doc, value)
                current_answer_tokens.append(["image", filename])
                continue
            if item_type == "table":
                current_answer_tokens.append(["table", value])
                continue
            if item_type == "text":
                text = value.replace('\xa0',' ').strip()
                amatch = re.match(r'^(\d+[a-zA-Z]*[ivxlcdm]*)\.?\s*(.*)', text, re.I)
                if amatch:
                    if current_answer_q is not None:
                        if current_answer_q in answers:
                            answers[current_answer_q].extend(current_answer_tokens)
                        else:
                            answers[current_answer_q] = current_answer_tokens
                    current_answer_q = amatch.group(1)
                    cleaned = amatch.group(2).strip()
                    current_answer_tokens = []
                    if cleaned:
                        current_answer_tokens.append(["text", cleaned, align_val])
                else:
                    if current_answer_q is not None:
                        current_answer_tokens.append(["text", text, align_val])
            continue

        if item_type == "image":
            filename = save_image(doc, value)
            current_tokens.append(["image", filename, align_val])
            continue
        if item_type == "table":
            current_tokens.append(["table", value])
            continue

        if item_type == "text":
            text = value.replace('\xa0', ' ').replace('\t', ' ')
            text = re.sub(r'\s+', ' ', text).strip()

            if text.upper() == "ANSWERS":
                reading_answers = True
                continue
            if not text:
                continue

            qmatch = question_regex.match(text)
            if qmatch:
                if current_qnum is not None:
                    if options:
                        mcq_questions.append({"qnum": current_qnum, "tokens": current_tokens, "options": options})
                    else:
                        structured_questions.append({"qnum": current_qnum, "tokens": current_tokens})
                
                current_qnum = qmatch.group(1)
                current_tokens = []
                if qmatch.group(2):
                    current_tokens.append(["text", qmatch.group(2), align_val])
                options = []
                continue

            opt = option_regex.match(text)
            if opt:
                options.append(opt.group(1))
                continue

            current_tokens.append(["text", text, align_val])

    if current_answer_q is not None:
        if current_answer_q in answers:
            answers[current_answer_q].extend(current_answer_tokens)
        else:
            answers[current_answer_q] = current_answer_tokens
    if current_qnum is not None:
        if options:
            mcq_questions.append({"qnum": current_qnum, "tokens": current_tokens, "options": options})
        else:
            structured_questions.append({"qnum": current_qnum, "tokens": current_tokens})

    return {
        "mcq_questions": mcq_questions,
        "structured_questions": structured_questions,
        "answers": answers
    }


def generate_qti_from_data(data, job_dir):
    """
    Takes the JSON structured data and generates the QTI 2.1 package zip.
    """
    items_dir = os.path.join(job_dir, "items")
    os.makedirs(items_dir, exist_ok=True)

    mcq_questions = data.get("mcq_questions", [])
    structured_questions = data.get("structured_questions", [])
    answers = data.get("answers", {})

    def random_id(length=3):
        return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

    item_refs = []
    item_images = {}

    # -----------------------------
    # HELPER: convert tokens to html
    # -----------------------------
    def tokens_to_html(tokens):

        html = ""

        for t in tokens:
            ttype = t[0]
            val = t[1]

            if ttype == "text":
                align_val = t[2] if len(t) > 2 else None
                safe = saxutils.escape(val)
                safe = safe.replace("&lt;","<").replace("&gt;",">")

                if align_val == "center":
                    html += f"<div style='text-align:center'>{safe}</div>\n"
                elif align_val == "right":
                    html += f"<div style='text-align:right'>{safe}</div>\n"
                elif align_val == "both":
                    html += f"<div style='text-align:justify'>{safe}</div>\n"
                else:
                    html += f"<div>{safe}</div>\n"

            elif ttype == "image":
                align_val = t[2] if len(t) > 2 else None
                
                if align_val == "center":
                    html += f"<div style='text-align:center'>\n<img alt=\"diagram\" src=\"../assets/{val}\"/>\n</div>\n"
                elif align_val == "right":
                    html += f"<div style='text-align:right'>\n<img alt=\"diagram\" src=\"../assets/{val}\"/>\n</div>\n"
                else:
                    html += f"<div>\n<img alt=\"diagram\" src=\"../assets/{val}\"/>\n</div>\n"

            elif ttype == "table":

                html += "<table border='1'>\n"

                for row in val:

                    html += "<tr>"

                    for cell in row:

                        safe = saxutils.escape(cell)

                        html += f"<td>{safe}</td>"

                    html += "</tr>\n"

                html += "</table>\n"

        return html

    # -----------------------------
    # CONVERT ANSWER TOKENS TO STRING
    # -----------------------------
    def answer_tokens_to_string(tokens):

        result = ""

        for t in tokens:
            ttype = t[0]
            val = t[1]

            if ttype == "text":
                result += val + " "

            elif ttype == "image":
                result += f"[image:{val}] "

            elif ttype == "table":
                result += "[table] "

        return result.strip()

    letters = ["a","b","c","d","e","f"]

    # -----------------------------
    # MCQ QUESTIONS -> QTI
    # -----------------------------
    for q in mcq_questions:

        answer_text = "None"
        qnum_str = str(q["qnum"])

        if qnum_str in answers:
            answer_text = answer_tokens_to_string(answers[qnum_str])

        num_match = re.search(r'\d+', qnum_str)
        qnum_int = int(num_match.group()) if num_match else 0

        item_id = f"Q{qnum_int:03d}_{random_id()}"

        filename = os.path.join(items_dir, f"{item_id}.xml")

        stem_html = tokens_to_html(q["tokens"])

        title_text = ""

        for t in q["tokens"]:
            if t[0] == "text":
                title_text = t[1][:70]
                break

        safe_title = saxutils.escape(title_text)

        xml = f'''<assessmentItem xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
adaptive="false"
identifier="{item_id}"
timeDependent="false"
title="Q{qnum_str} {safe_title}"
xsi:schemaLocation="http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_v2p1.xsd">
<responseDeclaration identifier="RESPONSE" cardinality="single" baseType="identifier">
<correctResponse>
<value>{saxutils.escape(answer_text)}</value>
</correctResponse>
</responseDeclaration>
<outcomeDeclaration baseType="float" cardinality="single" identifier="SCORE"/>
<itemBody>
<choiceInteraction responseIdentifier="RESPONSE" shuffle="true" maxChoices="1">
<prompt>
<div>
{stem_html}
</div>
</prompt>
'''

        for i, opt in enumerate(q["options"]):

            if i >= len(letters):
                break

            safe = saxutils.escape(opt)

            xml += f'<simpleChoice identifier="{letters[i]}">{safe}</simpleChoice>\n'

        xml += '''
</choiceInteraction>
</itemBody>
<responseProcessing template="http://www.imsglobal.org/question/qti_v2p1/rptemplates/match_correct"/>
</assessmentItem>
'''

        with open(filename,"w",encoding="utf8") as f:
            f.write(xml)

        item_refs.append(item_id)

        imgs = [v for t in q["tokens"] if t[0]=="image" for v in [t[1]]]

        item_images[item_id] = imgs

    # -----------------------------
    # STRUCTURED QUESTIONS -> QTI
    # -----------------------------
    for q in structured_questions:
        answer_text = "None"
        qnum_str = str(q["qnum"])

        if qnum_str in answers:
            answer_text = answer_tokens_to_string(answers[qnum_str])

        num_match = re.search(r'\d+', qnum_str)
        qnum_int = int(num_match.group()) if num_match else 0

        item_id = f"Q{qnum_int:03d}_{random_id()}"
        filename = os.path.join(items_dir, f"{item_id}.xml")

        stem_html = tokens_to_html(q["tokens"])

        xml = f'''<assessmentItem xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
adaptive="false"
identifier="{item_id}"
timeDependent="false"
title="Q{qnum_str}"
xsi:schemaLocation="http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_v2p1.xsd">
<responseDeclaration identifier="RESPONSE" cardinality="single" baseType="string"/>
<correctResponse>
<value>{saxutils.escape(answer_text)}</value>
</correctResponse>
<outcomeDeclaration baseType="float" cardinality="single" identifier="SCORE"/>
<itemBody>
<prompt>
{stem_html}
</prompt>
<extendedTextInteraction responseIdentifier="RESPONSE" expectedLength="400"/>
</itemBody>
<responseProcessing template="http://www.imsglobal.org/question/qti_v2p1/rptemplates/match_correct"/>
</assessmentItem>
'''

        with open(filename, "w", encoding="utf-8") as f:
            f.write(xml)

        item_refs.append(item_id)

        imgs = [v for t in q["tokens"] if t[0]=="image" for v in [t[1]]]

        item_images[item_id] = imgs

    # -----------------------------
    # assessment_test.xml
    # -----------------------------
    assessment_xml = '''
<assessmentTest xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
identifier="TEST1" title="Converted Test" xsi:schemaLocation="http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_v2p1.xsd">
<testPart identifier="part1" navigationMode="linear" submissionMode="individual">
<assessmentSection identifier="section1" title="Converted Test" visible="true">
'''

    for item in item_refs:
        assessment_xml += f'<assessmentItemRef href="items/{item}.xml" identifier="{item}"/>\n'

    assessment_xml += '''
</assessmentSection>
</testPart>
</assessmentTest>
'''

    with open(os.path.join(job_dir, "assessment_test.xml"),"w") as f:
        f.write(assessment_xml)

    # -----------------------------
    # imsmanifest.xml
    # -----------------------------
    manifest = '''
<manifest
xmlns="http://www.imsglobal.org/xsd/imscp_v1p1"
xmlns:imsmd="http://www.imsglobal.org/xsd/imsmd_v1p2"
xmlns:imsqti="http://www.imsglobal.org/xsd/imsqti_metadata_v2p1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
identifier="MANIFEST1"
xsi:schemaLocation="
http://www.imsglobal.org/xsd/imscp_v1p1 http://www.imsglobal.org/xsd/imscp_v1p1.xsd
http://www.imsglobal.org/xsd/imsmd_v1p2 http://www.imsglobal.org/xsd/imsmd_v1p2p4.xsd
http://www.imsglobal.org/xsd/imsqti_metadata_v2p1 http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_metadata_v2p1.xsd">
<metadata>
<schema>QTIv2.1 Package</schema>
<schemaversion>1.0.0</schemaversion>
</metadata>
<organizations/>
<resources>
'''

    manifest += '''
<resource identifier="RES_TEST" type="imsqti_test_xmlv2p1" href="assessment_test.xml">
<metadata/>
<file href="assessment_test.xml"/>
</resource>
'''

    for item in item_refs:

        manifest += f'''
<resource identifier="RES_{item}" type="imsqti_item_xmlv2p1" href="items/{item}.xml">
<metadata/>
<file href="items/{item}.xml"/>
'''

        for img in item_images.get(item, []):
            manifest += f'    <file href="assets/{img}"/>\n'

        manifest += "</resource>\n"

    manifest += '''
</resources>
</manifest>
'''

    with open(os.path.join(job_dir, "imsmanifest.xml"), "w", encoding="utf-8") as f:
        f.write(manifest)

    # -----------------------------
    # ZIP QTI PACKAGE
    # -----------------------------
    zip_path = os.path.join(job_dir, "qti_package.zip")
    zipf = zipfile.ZipFile(zip_path, "w")

    for root, dirs, files in os.walk(job_dir):
        for file in files:
            if file.endswith((".zip",".docx")):
                continue
            path = os.path.join(root, file)
            arcname = os.path.relpath(path, job_dir)
            zipf.write(path, arcname)

    zipf.close()

    return zip_path
