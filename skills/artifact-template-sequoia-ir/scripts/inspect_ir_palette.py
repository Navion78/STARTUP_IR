"""Read vector/theme color candidates from a PDF or PPTX without changing it.

Usage: python inspect_ir_palette.py company.pdf --output palette.json
The report is evidence for visual review, not an automatic brand decision.
"""
from pathlib import Path
from collections import defaultdict
import argparse, json, zipfile, colorsys
from xml.etree import ElementTree as ET

def record(store,color,page,kind,area=0):
    color=color.upper()
    if len(color)!=7 or color in ('#FFFFFF','#000000'): return
    row=store[color];row['pages'].add(page);row[kind]+=1;row['area']+=max(area,0)

def scan_pdf(source,store):
    import pdfplumber
    def rgb(v):
        if isinstance(v,(list,tuple)) and len(v)==3:
            return '#'+''.join(f'{max(0,min(255,round(float(x)*255))):02X}' for x in v)
    with pdfplumber.open(source) as doc:
        for i,p in enumerate(doc.pages,1):
            for r in p.rects:
                c=rgb(r.get('non_stroking_color'))
                if c:record(store,c,i,'shapes',r['width']*r['height'])
            for ch in p.chars:
                c=rgb(ch.get('non_stroking_color'))
                if c:record(store,c,i,'textRuns')
    return 'PDF RGB vector fills/text; raster images and non-RGB colors are excluded.'

def scan_pptx(source,store):
    ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    with zipfile.ZipFile(source) as z:
        theme={}
        if 'ppt/theme/theme1.xml' in z.namelist():
            root=ET.fromstring(z.read('ppt/theme/theme1.xml'))
            scheme=root.find('.//a:clrScheme',ns)
            if scheme is not None:
                for item in scheme:
                    if len(item):theme[item.tag.rsplit('}',1)[-1]]=item[0].get('val') or item[0].get('lastClr')
        import re
        slides=sorted((n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)),key=lambda x:int(re.search(r'(\d+)\.xml$',x).group(1)))
        for i,name in enumerate(slides,1):
            root=ET.fromstring(z.read(name))
            for item in root.findall('.//a:solidFill',ns):
                c=item.find('a:srgbClr',ns)
                if c is not None:value=c.get('val')
                else:
                    c=item.find('a:schemeClr',ns);value=theme.get(c.get('val')) if c is not None else None
                if value and len(value)==6:record(store,'#'+value,i,'shapes')
    return 'PPTX slide solidFill occurrences; theme1 is resolved where possible; transforms, images, and chart parts require visual review.'

def main():
    ap=argparse.ArgumentParser();ap.add_argument('input');ap.add_argument('--output');args=ap.parse_args()
    source=Path(args.input).expanduser().resolve()
    if not source.is_file():ap.error('Input file does not exist.')
    store=defaultdict(lambda:{'pages':set(),'shapes':0,'textRuns':0,'area':0.0})
    if source.suffix.lower()=='.pdf':method=scan_pdf(source,store)
    elif source.suffix.lower()=='.pptx':method=scan_pptx(source,store)
    else:ap.error('Only PDF and PPTX inputs are supported. Use visual review for other formats.')
    rows=[]
    for c,r in store.items():
        channels=[int(c[i:i+2],16)/255 for i in (1,3,5)]
        rows.append({'hex':c,'pages':sorted(r['pages']),'pageCount':len(r['pages']),'shapeOccurrences':r['shapes'],'textOccurrences':r['textRuns'],'vectorArea':round(r['area'],1),'saturation':round(colorsys.rgb_to_hsv(*channels)[1],3)})
    rows.sort(key=lambda x:(x['vectorArea'],x['pageCount'],x['shapeOccurrences']+x['textOccurrences']),reverse=True)
    result={'source':source.name,'method':method,'automaticBrandDecision':False,'note':'Compare repeated headers, titles, tables and diagrams on rendered pages. Do not select a photograph/logo/warning color solely from this report.','candidates':rows}
    data=json.dumps(result,ensure_ascii=False,indent=2)
    if args.output:
        target=Path(args.output).expanduser().resolve()
        if target==source:ap.error('Output must differ from the source file.')
        target.parent.mkdir(parents=True,exist_ok=True);target.write_text(data,encoding='utf-8')
    else:print(data)

if __name__=='__main__':main()
