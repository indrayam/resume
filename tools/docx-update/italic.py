"""Italicize a word inside a run by splitting the run into before / word / after."""
import re, sys, zipfile
src, dst, word = sys.argv[1], sys.argv[2], sys.argv[3]
zi = zipfile.ZipFile(src); d = zi.read('word/document.xml').decode()
n = 0
def fix(m):
    global n
    r = m.group(0)
    t = re.search(r'(<w:t[^>]*>)([^<]*)(</w:t>)', r)
    if not t or word not in t.group(2) or '<w:i/>' in r: return r
    rpr = re.search(r'<w:rPr>.*?</w:rPr>', r, re.S)
    rpr = rpr.group(0) if rpr else '<w:rPr></w:rPr>'
    # schema order: rStyle, rFonts, b, bCs, then i, iCs
    m = re.match(r'<w:rPr>((?:<w:rStyle[^>]*/>)?(?:<w:rFonts[^>]*/>)?(?:<w:b/>)?(?:<w:bCs/>)?)', rpr)
    irpr = rpr[:m.end()] + '<w:i/><w:iCs/>' + rpr[m.end():]
    head = r[:r.find('<w:rPr>')] if '<w:rPr>' in r else re.match(r'<w:r[^>]*>', r).group(0)
    before, after = t.group(2).split(word, 1)
    mk = lambda rp, s: f'{head}{rp}<w:t xml:space="preserve">{s}</w:t></w:r>' if s else ''
    n += 1
    return mk(rpr, before) + mk(irpr, word) + mk(rpr, after)
out = re.sub(r'<w:r[ >](?:(?!</w:r>).)*</w:r>', fix, d, flags=re.S)
print('italicized', n)
zo = zipfile.ZipFile(dst, 'w')
for it in zi.infolist():
    zo.writestr(it, out.encode() if it.filename == 'word/document.xml' else zi.read(it.filename))
zo.close()
