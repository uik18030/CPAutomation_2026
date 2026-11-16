import sys
import pdfplumber
import pypdfium2 as pdfium

for f in sys.argv[1:]:
    print(f)
    doc = pdfium.PdfDocument(f)
    with pdfplumber.open(f) as pdf:
        for i, p in enumerate(pdf.pages[:5]):
            q = doc[i]
            print(f"  p{i+1}: pdfplumber {p.width:.1f}x{p.height:.1f} "
                  f"mediabox={tuple(round(v, 1) for v in p.mediabox)} "
                  f"cropbox={tuple(round(v, 1) for v in p.cropbox)} rotate={p.rotation} | "
                  f"pdfium size={tuple(round(v, 1) for v in q.get_size())} "
                  f"cropbox={tuple(round(v, 1) for v in q.get_cropbox())} rotate={q.get_rotation()}")
            q.close()
    doc.close()
