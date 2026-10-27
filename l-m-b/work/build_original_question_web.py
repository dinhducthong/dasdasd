from pathlib import Path
from shutil import copyfile
from html import escape
import re, zipfile, io
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn
from PIL import Image

SRC = Path(r'C:\Users\Administrator\Downloads\ĐỀ CƯƠNG ÔN TẬP GIỮA HK1 NH 2026-2027.docx')
OUTDIR = Path(r'D:\!Download\banh-trang-nuong-indie-upgrade\2026-10-27\l-m-b\outputs\on-luyen-de-cuong-goc')
ASSETS = OUTDIR / 'assets'
OUTDIR.mkdir(parents=True, exist_ok=True); ASSETS.mkdir(exist_ok=True)
copyfile(SRC, OUTDIR / 'De_cuong_goc.docx')
doc = Document(SRC)

# Đáp án tham khảo theo thứ tự 50 câu trắc nghiệm trong đề cương.
MCQ_ANSWERS = ['C','A','D','C','C','B','A','D','C','A', 'A','D','B','B','B','D','B','B', 'C','A','A','B','A','D','D','B','D','D','A', 'C','A','B','A','B','C','C', 'C','D','A','B','B', 'C','D','B','A','A','D','B','C','C']
# Đáp án tham khảo theo thứ tự các phát biểu a-d trong 11 bảng đúng/sai.
TF_VALUES = [
    True,True,True,False,  True,False,False,True,  False,False,True,False,
    True,True,False,False,  True,True,True,True,  True,True,True,True,
    True,True,True,True,  True,True,True,True,  False,True,True,False,
    False,True,True,True,  True,True,True,True,
]
TF_ANSWERS = {(question, row): TF_VALUES[(question-1)*4 + (row-1)] for question in range(1, 12) for row in range(1, 5)}

image_number = 0
def image_html(node, part):
    global image_number
    rid = node.get(qn('r:embed')) or node.get(qn('r:id'))
    if not rid or rid not in part.related_parts: return ''
    image_part = part.related_parts[rid]
    image_number += 1
    name = f'asset-{image_number:03d}.png'
    path = ASSETS / name
    try:
        im = Image.open(io.BytesIO(image_part.blob)); im.load()
        if im.mode not in ('RGB','RGBA'): im = im.convert('RGBA')
        im.save(path, 'PNG')
        return f'<img src="assets/{name}" alt="Công thức hoặc hình trong đề gốc">'
    except Exception:
        raw = ASSETS / f'asset-{image_number:03d}.bin'
        raw.write_bytes(image_part.blob)
        return '<span class="missing-object">[Đối tượng trong đề gốc]</span>'

def para_html(p):
    parts=[]
    # Read Word runs in document order, so every image remains attached to its A/B/C/D label.
    for run in p._p.iterchildren():
        if run.tag not in (qn('w:r'), qn('w:hyperlink')): continue
        for child in run.iter():
            if child.tag == qn('w:t'):
                parts.append(escape(child.text or ''))
            elif child.tag == qn('w:tab'):
                parts.append('    ')
            elif child.tag == qn('w:br'):
                parts.append('<br>')
            elif child.tag == qn('a:blip') or child.tag == '{urn:schemas-microsoft-com:vml}imagedata':
                parts.append(image_html(child, p.part))
    content=''.join(parts)
    if not content: return ''
    return f'<p>{content}</p>'

def table_html(table, tf_question_id=None):
    rows=[]
    for row_index, row in enumerate(table.rows):
        cell_html=[]
        for cell_index, cell in enumerate(row.cells):
            contents=escape(cell.text).replace(chr(10), '<br>')
            if tf_question_id is not None and row_index > 0 and cell_index in (1, 2):
                value='true' if cell_index == 1 else 'false'
                contents=f'<input type="radio" name="t{tf_question_id}-{row_index}" value="{value}" aria-label="{value}">'
            cell_html.append(f'<td>{contents}</td>')
        cells=''.join(cell_html)
        answer = '' if tf_question_id is None or row_index == 0 else str(TF_ANSWERS.get((tf_question_id, row_index), '')).lower()
        rows.append(f'<tr data-answer="{answer}">{cells}</tr>')
    return '<div class="table-wrap"><table>' + ''.join(rows) + '</table></div>'

items=[]
section=None
current=None
def append_content(html):
    if current is not None:
        current['content'].append(html)
    elif section:
        items.append({'kind':'note','section':section,'content':[html]})

for child in doc.element.body.iterchildren():
    if child.tag == qn('w:p'):
        p=Paragraph(child, doc); text=p.text.strip()
        if 'PHẦN 1. CÂU TRẮC NGHIỆM' in text:
            section='mcq'; current=None; continue
        if text.startswith('II. PHẦN TRẮC NGHIỆM ĐÚNG SAI'):
            section='tf'; current=None; continue
        if text.startswith('III. PHẦN TỰ LUẬN'):
            break
        if not section: continue
        if re.match(r'^Câu\s*(?:\d+(?:\.\d+)?|\d+)\s*[:.]', text):
            current={'kind':'question','section':section,'title':text,'content':[], 'id':len([x for x in items if x['kind']=='question' and x['section']==section])+1}
            items.append(current)
        elif re.match(r'^(Bài\s*\d+|PHẦN)', text):
            current=None; items.append({'kind':'heading','section':section,'content':[para_html(p)]})
        else:
            append_content(para_html(p))
    elif child.tag == qn('w:tbl') and section:
        append_content(table_html(Table(child, doc), current['id'] if section == 'tf' and current else None))

mcq=[x for x in items if x['kind']=='question' and x['section']=='mcq']
tf=[x for x in items if x['kind']=='question' and x['section']=='tf']

def render_item(item):
    content=''.join(item['content'])
    if item['kind']=='heading': return f'<h2>{content}</h2>'
    if item['kind']=='note': return content
    if item['section']=='mcq':
        choices=''.join(f'<label><input type="radio" name="m{item["id"]}" value="{x}"> {x}</label>' for x in 'ABCD')
        control=f'<div class="answer-choice" data-answer="{MCQ_ANSWERS[item["id"]-1]}" aria-label="Chọn đáp án cho {escape(item["title"])}">{choices}</div><p class="feedback" aria-live="polite"></p>'
    else:
        control=''
    return f'<article class="question"><h3>{escape(item["title"])}</h3><div class="source">{content}</div>{control}</article>'

mcq_html=''.join(render_item(x) for x in items if x['section']=='mcq')
tf_html=''.join(render_item(x) for x in items if x['section']=='tf')

html=f'''<!doctype html>
<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ôn luyện đề cương Vật lí 12</title><style>
:root{{--blue:#174ea6;--navy:#13233f;--paper:#f5f7fb;--line:#d9e1ef;--muted:#5e6b80;--good:#067647;--bad:#b42318}}*{{box-sizing:border-box}}body{{margin:0;background:var(--paper);color:var(--navy);font:16px/1.55 Arial,sans-serif}}header{{background:#174ea6;color:#fff;padding:32px 18px}}.wrap{{max-width:980px;margin:auto}}h1{{margin:0;font-size:clamp(1.5rem,4vw,2.2rem)}}header p{{margin:5px 0 0}}main{{padding:24px 18px 60px}}.notice,.question{{background:#fff;border:1px solid var(--line);border-radius:10px;padding:18px;margin:14px 0}}.notice{{border-left:5px solid var(--blue)}}.tabs{{display:flex;gap:8px;flex-wrap:wrap;margin:20px 0}}button{{font:inherit;padding:10px 14px;border:1px solid var(--line);background:#fff;border-radius:8px;cursor:pointer}}button.active,button.primary{{background:var(--blue);color:#fff;border-color:var(--blue)}}.panel{{display:none}}.panel.active{{display:block}}h2{{font-size:1.2rem;margin:26px 0 8px}}h3{{font-size:1rem;margin:0 0 10px}}.source p{{margin:5px 0;white-space:pre-wrap}}.source img{{display:inline-block;max-width:100%;vertical-align:middle;margin:2px 5px;image-rendering:auto}}.table-wrap{{overflow-x:auto;margin:8px 0}}table{{border-collapse:collapse;min-width:500px;width:100%}}td{{border:1px solid var(--line);padding:8px;vertical-align:top}}tr:first-child td{{font-weight:bold;background:#edf3ff}}.answer-choice{{display:flex;gap:14px;flex-wrap:wrap;margin-top:14px;padding-top:12px;border-top:1px solid var(--line)}}.answer-choice label{{cursor:pointer}}input{{accent-color:var(--blue);margin-right:5px}}.feedback{{font-weight:bold;margin:9px 0 0}}.is-correct{{background:#effcf5!important;color:var(--good)}}.is-incorrect{{background:#fff5f4!important;color:var(--bad)}}.actions{{display:flex;gap:10px;flex-wrap:wrap;margin:20px 0}}#status{{font-weight:bold}}.missing-object{{color:#a33;font-style:italic}}@media(max-width:560px){{main{{padding:16px 10px}}.question,.notice{{padding:14px}}}}
</style></head><body><header><div class="wrap"><h1>Ôn luyện đề cương Vật lí 12</h1><p>Nội dung trích từ đề cương gốc</p></div></header><main class="wrap">
<div class="notice"><strong>Cách làm:</strong> Nội dung câu hỏi, lựa chọn, bảng đúng/sai và hình/công thức được lấy trực tiếp từ đề gốc. Chọn đáp án của bạn; trang này ghi nhận số câu đã làm, không tự chấm vì đề gốc không kèm đáp án chính thức. <a href="De_cuong_goc.docx">Mở đề gốc</a>.</div>
<div class="tabs" role="tablist"><button class="active" data-target="mcq">Trắc nghiệm ({len(mcq)} câu)</button><button data-target="tf">Đúng sai ({len(tf)} câu)</button></div>
<section id="mcq" class="panel active"><div class="actions"><button class="primary" data-count="mcq">Kiểm tra số câu đã chọn</button><button data-reset="mcq">Làm lại phần này</button></div><p id="mcq-status" aria-live="polite"></p>{mcq_html}</section>
<section id="tf" class="panel"><div class="actions"><button class="primary" data-count="tf">Kiểm tra số câu đã chọn</button><button data-reset="tf">Làm lại phần này</button></div><p id="tf-status" aria-live="polite"></p>{tf_html}</section>
</main><script>
document.querySelectorAll('.tabs button').forEach(b=>b.onclick=()=>{{document.querySelectorAll('.tabs button').forEach(x=>x.classList.remove('active'));document.querySelectorAll('.panel').forEach(x=>x.classList.remove('active'));b.classList.add('active');document.getElementById(b.dataset.target).classList.add('active')}});
document.querySelectorAll('[data-count]').forEach(b=>b.onclick=()=>{{let s=b.dataset.count, total=new Set([...document.querySelectorAll('#'+s+' input')].map(x=>x.name)).size, answered=new Set([...document.querySelectorAll('#'+s+' input:checked')].map(x=>x.name)).size;document.getElementById(s+'-status').textContent='Bạn đã chọn '+answered+'/'+total+' đáp án.'}});
function setFeedback(el, right, answer){{el.classList.remove('is-correct','is-incorrect');el.classList.add(right?'is-correct':'is-incorrect');return right?'Chọn đúng.':'Chọn sai. Đáp án tham khảo: '+answer+'.'}}
document.querySelectorAll('#mcq input').forEach(input=>input.addEventListener('change',()=>{{let choice=input.closest('.answer-choice'), card=input.closest('.question'), correct=choice.dataset.answer, feedback=card.querySelector('.feedback');feedback.textContent=setFeedback(card,input.value===correct,correct)}}));
document.querySelectorAll('#tf input').forEach(input=>input.addEventListener('change',()=>{{let row=input.closest('tr'), correct=row.dataset.answer==='true', answer=correct?'Đúng':'Sai';let result=setFeedback(row,(input.value==='true')===correct,answer);let cell=row.cells[0];let note=cell.querySelector('.feedback');if(!note){{note=document.createElement('div');note.className='feedback';cell.append(note)}}note.textContent=result}}));
document.querySelectorAll('[data-reset]').forEach(b=>b.onclick=()=>{{let s=b.dataset.reset;document.querySelectorAll('#'+s+' input').forEach(x=>x.checked=false);document.querySelectorAll('#'+s+' .is-correct,#'+s+' .is-incorrect').forEach(x=>x.classList.remove('is-correct','is-incorrect'));document.querySelectorAll('#'+s+' .feedback').forEach(x=>x.textContent='');document.getElementById(s+'-status').textContent='Đã làm lại phần này.'}});
</script></body></html>'''
(OUTDIR/'index.html').write_text(html,encoding='utf-8')
print(f'MCQ: {len(mcq)}; TF: {len(tf)}; images converted: {image_number}')
