from PIL import Image
import zipfile, io

p = r'C:\Users\Administrator\Downloads\ĐỀ CƯƠNG ÔN TẬP GIỮA HK1 NH 2026-2027.docx'
with zipfile.ZipFile(p) as z:
    data = z.read('word/media/image2.wmf')
try:
    im = Image.open(io.BytesIO(data))
    print(im.format, im.size)
    im.load()
    print('loaded')
except Exception as exc:
    print(type(exc).__name__, str(exc))
