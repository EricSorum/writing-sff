import docx
import html
import os

def is_list_paragraph(paragraph):
    # DOCX list paragraphs often have numbering or bullet styles
    if paragraph.style.name.lower().startswith('list') or paragraph._p.pPr is not None and paragraph._p.pPr.numPr is not None:
        return True
    return False

def get_list_type(paragraph):
    # Try to infer list type from style or numbering
    style = paragraph.style.name.lower()
    if 'roman' in style:
        return 'ol', 'I'  # Roman numerals
    elif 'alpha' in style:
        return 'ol', 'A'  # Capital letters
    elif 'number' in style:
        return 'ol', '1'  # Arabic numbers
    elif 'lowerroman' in style:
        return 'ol', 'i'  # Lowercase Roman numerals
    elif 'loweralpha' in style:
        return 'ol', 'a'  # Lowercase letters
    elif 'bullet' in style:
        return 'ul', None
    # Fallback: check numbering format
    numPr = paragraph._p.pPr.numPr if paragraph._p.pPr is not None else None
    if numPr is not None:
        # Try to guess from numbering format
        # This is a simplification; for more accuracy, use python-docx's numbering map
        return 'ol', '1'
    return None, None

def convert_docx_to_html(docx_path, output_html='index_lists.html'):
    doc = docx.Document(docx_path)
    html_content = ['<!DOCTYPE html>',
                   '<html>',
                   '<head>',
                   '<meta charset="utf-8">',
                   '<title>Writing Science Fiction and Fantasy</title>',
                   '<style>',
                   'body { font-family: Arial, sans-serif; line-height: 1.6; margin: 40px; }',
                   'h1, h2, h3 { color: #333; }',
                   'p { margin-bottom: 1em; }',
                   'ol[type="I"] { list-style-type: upper-roman; }',
                   'ol[type="A"] { list-style-type: upper-alpha; }',
                   'ol[type="1"] { list-style-type: decimal; }',
                   'ol[type="i"] { list-style-type: lower-roman; }',
                   'ol[type="a"] { list-style-type: lower-alpha; }',
                   '</style>',
                   '</head>',
                   '<body>']

    in_list = False
    list_type = None
    list_tag = None
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            if in_list:
                html_content.append(f'</{list_tag}>')
                in_list = False
            level = int(para.style.name[-1])
            html_content.append(f'<h{level}>{html.escape(para.text)}</h{level}>')
        elif is_list_paragraph(para):
            this_list_tag, this_list_type = get_list_type(para)
            if not in_list or this_list_tag != list_tag or this_list_type != list_type:
                if in_list:
                    html_content.append(f'</{list_tag}>')
                attrs = f' type="{this_list_type}"' if this_list_type else ''
                html_content.append(f'<{this_list_tag}{attrs}>')
                in_list = True
                list_tag = this_list_tag
                list_type = this_list_type
            # List item content
            formatted_text = ''
            for run in para.runs:
                text = html.escape(run.text)
                if run.bold:
                    text = f'<strong>{text}</strong>'
                if run.italic:
                    text = f'<em>{text}</em>'
                if run.underline:
                    text = f'<u>{text}</u>'
                formatted_text += text
            html_content.append(f'<li>{formatted_text}</li>')
        else:
            if in_list:
                html_content.append(f'</{list_tag}>')
                in_list = False
            formatted_text = ''
            for run in para.runs:
                text = html.escape(run.text)
                if run.bold:
                    text = f'<strong>{text}</strong>'
                if run.italic:
                    text = f'<em>{text}</em>'
                if run.underline:
                    text = f'<u>{text}</u>'
                formatted_text += text
            if formatted_text.strip():
                html_content.append(f'<p>{formatted_text}</p>')
    if in_list:
        html_content.append(f'</{list_tag}>')
    html_content.extend(['</body>', '</html>'])
    with open(output_html, 'w', encoding='utf-8') as f:
        f.write('\n'.join(html_content))
    return output_html

if __name__ == "__main__":
    docx_file = "Writing Science Fiction and Fantasy.docx"
    output_file = convert_docx_to_html(docx_file, output_html='index_lists.html')
    print(f"Conversion complete! HTML file saved as: {output_file}") 