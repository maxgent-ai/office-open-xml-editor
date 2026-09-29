from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from pathlib import Path
A='http://schemas.openxmlformats.org/drawingml/2006/main';P='http://schemas.openxmlformats.org/presentationml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';PKG='http://schemas.openxmlformats.org/package/2006/relationships';CT='http://schemas.openxmlformats.org/package/2006/content-types'
pat='<a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="D21D54"/></a:fgClr><a:bgClr><a:srgbClr val="12CED4"/></a:bgClr></a:pattFill>'
outline='<a:pattFill prst="dnDiag"><a:fgClr><a:srgbClr val="00A650"/></a:fgClr><a:bgClr><a:srgbClr val="FF8800"/></a:bgClr></a:pattFill>'
def shape(i,x,y,w,h,text,para='',run='',field=False,scene='',list_style=''):
 t=f'<a:fld id="{{0FBB9679-570D-4509-9600-E07789CF24BB}}" type="datetime"><a:rPr sz="4800" b="1">{run}<a:latin typeface="Arial"/></a:rPr>' if field else f'<a:r><a:rPr sz="4800" b="1">{run}<a:latin typeface="Arial"/></a:rPr>'
 return f'''<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Control {i}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>{scene}</p:spPr><p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle>{list_style}</a:lstStyle><a:p>{para}{t}<a:t>{text}</a:t>{'</a:fld>' if field else '</a:r>'}</a:p></p:txBody></p:sp>'''
def slide(shapes):return f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>'
s1=slide(shape(2,600000,500000,5000000,1300000,'MMMMMMMM','',pat,field=True)+shape(3,600000,2500000,5000000,1300000,'MMMMMMMM','',pat)+shape(4,600000,4500000,5000000,1300000,'MMMMMMMM','',pat.replace('pct50','dnDiag'),field=True))
s2=slide(shape(2,600000,500000,8000000,1600000,'OUTLINE TEXT','',f'<a:ln w="50800">{outline}</a:ln><a:solidFill><a:srgbClr val="101010"/></a:solidFill>'))
sc='<a:scene3d><a:camera prst="perspectiveRelaxed"><a:rot lat="1800000" lon="1800000" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d extrusionH="457200" prstMaterial="plastic"/>'
pat3=pat.replace('pct50','dnDiag')
s3=slide(shape(2,600000,500000,5000000,1500000,'MMMMMMMM','',pat3,scene=sc)+shape(3,600000+5486400+50800,500000,5000000,1500000,'MMMMMMMM','',pat3,scene=sc))
parts={
 '[Content_Types].xml':f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
 '_rels/.rels':f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
 'ppt/presentation.xml':f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst>'+''.join(f'<p:sldId id="{255+i}" r:id="rId{i}"/>' for i in range(1,4))+'</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
 'ppt/_rels/presentation.xml.rels':f'<Relationships xmlns="{PKG}">'+''.join(f'<Relationship Id="rId{i}" Type="{R}/slide" Target="slides/slide{i}.xml"/>' for i in range(1,4))+'</Relationships>',
 'ppt/slides/slide1.xml':s1,'ppt/slides/slide2.xml':s2,'ppt/slides/slide3.xml':s3,
}
out=Path('tests/patterns/fixtures/pptx-field-outline-projection.pptx')
with ZipFile(out,'w',ZIP_DEFLATED) as z:
 for path,content in sorted(parts.items()):
  zi=ZipInfo(path,(2020,1,1,0,0,0));zi.compress_type=ZIP_DEFLATED;z.writestr(zi,content)
print(out,out.stat().st_size)
