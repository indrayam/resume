"""Move CAREER OVERVIEW + the first role (paras 21-24) from page 1 to the top of page 2."""
import re, sys, zipfile
src, dst = sys.argv[1], sys.argv[2]
zi = zipfile.ZipFile(src); d = zi.read('word/document.xml').decode()
b0 = d.find('<w:body>')
spans = [(m.start()+b0, m.end()+b0) for m in re.finditer(r'<w:p[ >].*?</w:p>', d[b0:], re.S)]
P = [d[a:b] for a, b in spans]
txt = lambda p: re.sub(r'<[^>]+>', '', p)
assert 'CAREER OVERVIEW' in txt(P[21]) and 'DevTools AI Agent' in txt(P[24])
assert '<w:br w:type="page"/>' in P[51] and 'Architected and delivered 12+' in txt(P[52])
block = ''.join(P[21:25])
# cut 21-24, then insert the block right after the page-break paragraph 51
out = d[:spans[21][0]] + d[spans[24][1]:spans[51][1]] + block + d[spans[51][1]:]
zo = zipfile.ZipFile(dst, 'w')
for it in zi.infolist():
    zo.writestr(it, out.encode() if it.filename == 'word/document.xml' else zi.read(it.filename))
zo.close()
