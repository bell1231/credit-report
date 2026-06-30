import xml.etree.ElementTree as ET

path = r'C:/Users/Administrator/CodeBuddy/20260625211524/unpacked_ref/word/document.xml'
tree = ET.parse(path)
root = tree.getroot()

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

texts = []
for p in root.iter(f'{{{W}}}p'):
    t_text = ''
    for t in p.iter(f'{{{W}}}t'):
        if t.text:
            t_text += t.text
    if t_text.strip():
        texts.append(t_text.strip())

with open(r'C:/Users/Administrator/CodeBuddy/20260625211524/reference_text.txt', 'w', encoding='utf-8') as f:
    for line in texts:
        f.write(line + '\n')

print(f'Extracted {len(texts)} paragraphs')
