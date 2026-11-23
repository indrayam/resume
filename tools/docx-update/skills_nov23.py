import sys,zipfile
src,dst=sys.argv[1],sys.argv[2]
z=zipfile.ZipFile(src); s=z.read('word/document.xml').decode()
old='<w:t xml:space="preserve">Git, Jenkins, Artifactory, Spinnaker </w:t>'
assert s.count(old)==1
s=s.replace(old,'<w:t xml:space="preserve">GitHub, GitHub Actions, Artifactory, SAST/DAST </w:t>')
i=s.find('<w:t>Public Cloud Technologies:</w:t>'); a=s.rfind('<w:p ',0,i)
b=s.find('</w:p>',s.find('<w:t xml:space="preserve"> Lambda)</w:t>'))+6
pair=s[a:b]
lam='<w:r><w:rPr><w:rFonts w:ascii="Montserrat" w:eastAsia="Calibri" w:hAnsi="Montserrat" w:cs="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve"> Lambda)</w:t></w:r>'
assert pair.count(lam)==1
new=(pair.replace(lam,'').replace('<w:t>AWS (EC2, EKS, S3,</w:t>','<w:t>Claude, Codex, Cursor, GrokBot (Outer Loop)</w:t>')
        .replace('<w:t>Public Cloud Technologies:</w:t>','<w:t>AI Coding Harness:</w:t>')
        .replace('w14:paraId="1E6209F2"','w14:paraId="1E6209F3"').replace('w14:paraId="429D0C01"','w14:paraId="429D0C02"'))
assert 'paraId="1E6209F3"' not in s and 'paraId="429D0C02"' not in s
IND='<w:ind w:left="2074" w:hanging="2074"/>'
def unindent(s,marker):
    k=s.find(marker); a=s.rfind('<w:p ',0,k); seg=s[a:k]; assert seg.count(IND)==1
    return s[:a]+seg.replace(IND,'')+s[k:]
s=s[:a]+new+s[a:]
s=unindent(s,'<w:t>Claude, Codex, Cursor, GrokBot (Outer Loop)</w:t>')
s=unindent(s,'<w:t xml:space="preserve">GitHub, GitHub Actions, Artifactory, SAST/DAST </w:t>')
out=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
for it in z.infolist():
    out.writestr(it, s.encode() if it.filename=='word/document.xml' else z.read(it.filename))
out.close()
