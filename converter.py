from docx import Document
from docx.oxml.ns import qn
import os
import re

def convert_docx_to_qti(input_docx, job_dir):

    doc = Document(input_docx)

    assets_dir = os.path.join(job_dir, "assets")
    items_dir = os.path.join(job_dir, "items")
    os.makedirs(assets_dir, exist_ok=True)

    # -----------------------------
    # IMAGE HANDLING
    # -----------------------------
    image_counter = 0

    def save_image(doc, rId):
        nonlocal image_counter
        image_part = doc.part.related_parts[rId]
        filename = f"img_{image_counter}.jpg"
        with open(os.path.join(assets_dir, filename), "wb") as f:
            f.write(image_part.blob)
        image_counter += 1
        return filename

    # -----------------------------
    # ITERATE WORD CONTENT
    # -----------------------------
    def iter_block_items(doc):
        body = doc.element.body
        for child in body.iterchildren():

            # -----------------------------
            # PARAGRAPH
            # -----------------------------
            if child.tag.endswith('p'):

                tokens = []

                for run in child.xpath("./w:r"):

                    text = "".join(
                        node.text or ""
                        for node in run.xpath(".//w:t")
                    )

                    if not text:
                        continue

                    is_sup = run.xpath(".//w:vertAlign[@w:val='superscript']")
                    is_sub = run.xpath(".//w:vertAlign[@w:val='subscript']")
                    is_bold = run.xpath(".//w:b")
                    is_italic = run.xpath(".//w:i")

                    if is_sup:
                        text = f"<sup>{text}</sup>"
                    if is_sub:
                        text = f"<sub>{text}</sub>"
                    if is_bold:
                        text = f"<strong>{text}</strong>"
                    if is_italic:
                        text = f"<em>{text}</em>"

                    tokens.append(text)

                paragraph_text = "".join(tokens).strip()

                if paragraph_text:
                    yield ("text", paragraph_text)

                for node in child.iter():
                    if node.tag.endswith('blip'):
                        rId = node.get(qn('r:embed'))
                        yield ("image", rId)

            # -----------------------------
            # TABLE
            # -----------------------------
            if child.tag.endswith('tbl'):

                table_data = []

                for row in child.xpath(".//w:tr"):

                    row_data = []

                    for cell in row.xpath(".//w:tc"):

                        cell_text = "".join(
                            node.text or ""
                            for node in cell.xpath(".//w:t")
                        ).strip()

                        row_data.append(cell_text)

                    table_data.append(row_data)

                yield ("table", table_data)

    # -----------------------------
    # TABLE OPTION EXTRACTION
    # -----------------------------
    def extract_table_options(table):
        options = []
        for row in table:
            for cell in row:
                cell = cell.strip()
                if cell:
                    options.append(cell)
        return options

    # -----------------------------
    # REGEX PATTERNS (STRICT)
    # -----------------------------
    question_regex = re.compile(r'^(\d+)\s+(.*)')
    option_regex = re.compile(r'^[A-D][\.\)]\s*(.*)')
    part_regex = re.compile(r'^\(([a-z])\)\s*(.*)', re.I)
    subpart_regex = re.compile(r'^\(([ivxlcdm]+)\)\s*(.*)', re.I)
    marks_regex = re.compile(r'\[(\d+)\]')

    # -----------------------------
    # STORAGE
    # -----------------------------
    mcq_questions = []
    structured_questions = []

    current_qnum = None
    current_tokens = []
    options = []

    answers = {}
    reading_answers = False

    current_answer_q = None
    current_answer_tokens = []

    # -----------------------------
    # PARSE DOCUMENT
    # -----------------------------
    for item_type, value in iter_block_items(doc):

        if reading_answers:

            if item_type == "image" and current_answer_q is not None:
                filename = save_image(doc, value)
                current_answer_tokens.append(("image", filename))
                continue

            if item_type == "table":
                current_answer_tokens.append(("table", value))
                continue

            if item_type == "text":

                text = value.replace('\xa0',' ').strip()

                amatch = re.match(r'^(\d+)', text)

                if amatch:

                    if current_answer_q is not None:
                        answers[current_answer_q] = current_answer_tokens

                    current_answer_q = amatch.group(1)

                    cleaned = re.sub(r'^\d+\s*', '', text)

                    current_answer_tokens = []

                    if cleaned:
                        current_answer_tokens.append(("text", cleaned))

                else:

                    if current_answer_q is not None:
                        current_answer_tokens.append(("text", text))

            continue

        if item_type == "image":

            filename = save_image(doc, value)
            current_tokens.append(("image", filename))
            continue

        if item_type == "table":

            current_tokens.append(("table", value))
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
                        mcq_questions.append({
                            "qnum": current_qnum,
                            "tokens": current_tokens,
                            "options": options
                        })

                    else:
                        structured_questions.append({
                            "qnum": current_qnum,
                            "tokens": current_tokens
                        })

                current_qnum = qmatch.group(1)

                current_tokens = [
                    ("text", qmatch.group(2))
                ]

                options = []

                continue

            opt = option_regex.match(text)

            if opt:
                options.append(opt.group(1))
                continue

            current_tokens.append(("text", text))

    if current_answer_q is not None:
        answers[current_answer_q] = current_answer_tokens

    if current_qnum is not None:

        if options:
            mcq_questions.append({
                "qnum": current_qnum,
                "tokens": current_tokens,
                "options": options
            })
        else:
            structured_questions.append({
                "qnum": current_qnum,
                "tokens": current_tokens
            })

    # -----------------------------
    # QTI 2.1 OUTPUT GENERATION
    # -----------------------------
    import zipfile
    import random
    import string
    import xml.sax.saxutils as saxutils

    os.makedirs(items_dir, exist_ok=True)

    def random_id(length=3):
        return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

    item_refs = []
    item_images = {}

    # -----------------------------
    # HELPER: convert tokens to html
    # -----------------------------
    def tokens_to_html(tokens):

        html = ""

        for ttype, val in tokens:

            if ttype == "text":

                safe = saxutils.escape(val)
                safe = safe.replace("&lt;","<").replace("&gt;",">")

                html += f"<div>{safe}</div>\n"

            elif ttype == "image":

                html += f'''
<div>
<img alt="diagram" src="../assets/{val}"/>
</div>
'''

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

        for ttype, val in tokens:

            if ttype == "text":
                result += val + " "

            elif ttype == "image":
                result += f"[image:{val}] "

            elif ttype == "table":
                result += "[table] "

        return result.strip()

    letters = ["a","b","c","d","e","f"]

    # -----------------------------
    # MCQ QUESTIONS → QTI
    # -----------------------------
    for q in mcq_questions:

        answer_text = "None"

        if q["qnum"] in answers:
            answer_text = answer_tokens_to_string(answers[q["qnum"]])

        item_id = f"Q{int(q['qnum']):03d}_{random_id()}"

        filename = os.path.join(items_dir, f"{item_id}.xml")

        stem_html = tokens_to_html(q["tokens"])

        title_text = ""

        for ttype, val in q["tokens"]:
            if ttype == "text":
                title_text = val[:70]
                break

        safe_title = saxutils.escape(title_text)

        xml = f'''<assessmentItem xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
adaptive="false"
identifier="{item_id}"
timeDependent="false"
title="Q{int(q['qnum']):03d} {safe_title}"
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

        imgs = [v for t,v in q["tokens"] if t=="image"]

        item_images[item_id] = imgs

    # -----------------------------
    # STRUCTURED QUESTIONS → QTI
    # -----------------------------
    for q in structured_questions:
        answer_text = "None"

        if q["qnum"] in answers:
            answer_text = answer_tokens_to_string(answers[q["qnum"]])

        item_id = f"Q{int(q['qnum']):03d}_{random_id()}"
        filename = os.path.join(items_dir, f"{item_id}.xml")

        stem_html = tokens_to_html(q["tokens"])

        xml = f'''<assessmentItem xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
adaptive="false"
identifier="{item_id}"
timeDependent="false"
title="Q{int(q['qnum']):03d}"
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

        imgs = [v for t,v in q["tokens"] if t=="image"]

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

    print(f"QTI 2.1 package created: {zip_path}")
    return zip_path