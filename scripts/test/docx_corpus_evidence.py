"""Independent ZIP/XML corpus oracle. Never imports OfficeMCP or the SuperDoc SDK.

Reads XML parts individually and hashes binary members in 1 MiB chunks. Large
media are not inflated into retained buffers. Scope is recorded in every report.
"""
import copy
import hashlib
import json
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
PR = '{http://schemas.openxmlformats.org/package/2006/relationships}'
W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
W15 = '{http://schemas.microsoft.com/office/word/2012/wordml}'

def digest_file(file):
    h = hashlib.sha256()
    with open(file, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()

def member_hash(z, name):
    h = hashlib.sha256()
    with z.open(name) as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()

def parse(z, name):
    b = z.read(name)
    if b'<!DOCTYPE' in b or b'<!ENTITY' in b:
        raise ValueError('Unexpected XML declaration: ' + name)
    return ET.fromstring(b)

def value(n, key, default=None):
    return n.get(W + key, default) if n is not None else default

def properties(n):
    if n is None:
        return {}
    out = {}
    for key in ['b', 'i', 'strike', 'dstrike', 'caps', 'smallCaps']:
        v = n.find(W + key)
        if v is not None:
            out[key] = value(v, 'val', '1') not in ['0', 'false', 'off']
    for key in ['u', 'vertAlign', 'highlight', 'color', 'sz', 'rStyle']:
        v = n.find(W + key)
        if v is not None:
            out[key] = value(v, 'val', 'single' if key == 'u' else None)
    fonts = n.find(W + 'rFonts')
    if fonts is not None:
        for key in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
            if W + key in fonts.attrib:
                out['font:' + key] = fonts.attrib[W + key]
    return out

def text(n):
    return ''.join(x.text or '' for x in n.iter(W + 't'))

def snapshot(file):
    with zipfile.ZipFile(file) as z:
        names = z.namelist()
        styles = parse(z, 'word/styles.xml') if 'word/styles.xml' in names else ET.Element('styles')
        style_map = {s.get(W+'styleId'): s for s in styles.findall(W+'style')}
        default_style = next((s.get(W+'styleId') for s in style_map.values() if s.get(W+'type') == 'paragraph' and s.get(W+'default') in ['1','true']), None)
        defaults = properties(styles.find(W+'docDefaults/'+W+'rPrDefault/'+W+'rPr'))
        def inherited(style_id, seen=None):
            seen = set() if seen is None else seen
            if not style_id or style_id in seen or style_id not in style_map:
                return {}
            seen.add(style_id)
            style = style_map[style_id]
            result = inherited(value(style.find(W+'basedOn'), 'val'), seen)
            result.update(properties(style.find(W+'rPr')))
            return result
        def paragraph(p):
            style_id = value(p.find(W+'pPr/'+W+'pStyle'), 'val', default_style)
            base = {**defaults, **inherited(style_id)}
            spans, position, visible = [], 0, []
            for r in p.iter(W+'r'):
                s = ''.join((c.text or '') if c.tag in [W+'t', W+'delText'] else '\t' if c.tag == W+'tab' else '\n' if c.tag in [W+'br', W+'cr'] else '' for c in r)
                if not s:
                    continue
                props = properties(r.find(W+'rPr'))
                effective = {**base, **inherited(props.get('rStyle')), **props}
                effective.pop('rStyle', None)
                # Missing and explicit false marks have the same meaning.
                effective = {k:v for k,v in effective.items() if v is not False and v is not None}
                key = json.dumps(effective, ensure_ascii=False, sort_keys=True)
                if spans and spans[-1][2] == key:
                    spans[-1][1] += len(s)
                else:
                    spans.append([position, len(s), key])
                visible.append(s); position += len(s)
            return {'text': text(p), 'runText': ''.join(visible), 'spans': spans}
        stories, controls, images, comments, table_records = {}, [], [], [], []
        binaries = sorted(member_hash(z,n) for n in names if not n.endswith('/') and not n.endswith(('.xml','.rels')))
        for name in sorted(n for n in names if n == 'word/document.xml' or (n.startswith(('word/header','word/footer','word/footnotes','word/endnotes')) and n.endswith('.xml'))):
            d = parse(z, name)
            if name.endswith(('footnotes.xml','endnotes.xml')):
                ps = [p for note in d if note.get(W+'id') not in ['-1','0'] for p in note.iter(W+'p')]
                if not ps:
                    continue
            else:
                ps = list(d.iter(W+'p'))
            stories[name] = [paragraph(p) for p in ps]
            for index,t in enumerate(d.iter(W+'tbl')):
                rows = [[text(c) for c in r.findall(W+'tc')] for r in t.findall(W+'tr')]
                table_records.append({'part':name,'index':index,'rows':rows,
                    'grid':[value(n,'w') for n in t.findall(W+'tblGrid/'+W+'gridCol')],
                    'merges':[(value(n,'val','continue'),n.tag.rsplit('}',1)[-1]) for n in t.iter() if n.tag in [W+'gridSpan', W+'vMerge']]})
            for sdt in d.iter(W+'sdt'):
                pr = sdt.find(W+'sdtPr')
                controls.append({'part':name,'tag':value(pr.find(W+'tag'),'val') if pr is not None else None,
                    'alias':value(pr.find(W+'alias'),'val') if pr is not None else None,'text':text(sdt)})
            rel_name = posixpath.join(posixpath.dirname(name),'_rels',posixpath.basename(name)+'.rels')
            rels = {r.get('Id'): r for r in parse(z,rel_name)} if rel_name in names else {}
            for pi,p in enumerate(ps):
                for drawing in p.iter(W+'drawing'):
                    blips = list(drawing.iter(A+'blip'))
                    refs=[]
                    for blip in blips:
                        rel=rels.get(blip.get(R+'embed'))
                        if rel is None:
                            raise ValueError('Missing image relationship')
                        target=rel.get('Target'); member=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(name),target))
                        refs.append(member_hash(z,member))
                    ext = next(drawing.iter(WP+'extent'),None); pr = next(drawing.iter(WP+'docPr'),None)
                    images.append({'part':name,'paragraph':pi,'hashes':refs,'size':[ext.get('cx'),ext.get('cy')] if ext is not None else None,
                        'alt':pr.get('descr','') if pr is not None else '', 'inline':drawing.find(WP+'inline') is not None})
            if name == 'word/document.xml' and 'word/comments.xml' in names:
                cdoc=parse(z,'word/comments.xml'); ex=parse(z,'word/commentsExtended.xml') if 'word/commentsExtended.xml' in names else ET.Element('empty')
                for c in cdoc:
                    cid=c.get(W+'id'); cp=list(c.iter(W+'p')); paraid=cp[-1].get(W14+'paraId') if cp else None
                    resolution=[n for n in ex if n.get(W15+'paraId')==paraid] if paraid else []
                    anchors=[i for i,p in enumerate(ps) if any(n.get(W+'id')==cid for n in p.iter(W+'commentRangeStart'))]
                    comments.append({'text':text(c),'anchors':anchors,'resolved':len(resolution)==1 and resolution[0].get(W15+'done') in ['1','true']})
        result={'stories':stories,'tables':table_records,'controls':controls,'images':images,'comments':comments,'binaries':binaries}
        return result

def inventory(file):
    s=snapshot(file); ps=s['stories']['word/document.xml']; texts=[p['text'] for p in ps]
    return {'file':file,'sha256':digest_file(file),'counts':{k:len(s[k]) for k in ['tables','controls','images','comments','binaries']},
        'paragraphs':len(ps),'stories':list(s['stories']), 'uniqueParagraphs':[{'index':i,'text':p['text']} for i,p in enumerate(ps) if 20<=len(p['text'])<=400 and texts.count(p['text'])==1][:12],
        'tables':[{'index':t['index'],'part':t['part'],'rows':t['rows'][:3],'merged':bool(t['merges'])}for t in s['tables'][:4]],'controls':s['controls'][:5],'comments':s['comments']}

def verify(source,saved,spec):
    before=snapshot(source); after=snapshot(saved); expected=copy.deepcopy(before)
    kind=spec['kind']
    if kind in ['text','table']:
        # Same-length replacements make unchanged formatting coverage exact.
        find,replacement=spec['find'],spec['replace']; assert len(find)==len(replacement)
        part=spec.get('part','word/document.xml'); index=spec['paragraphIndex']; p=expected['stories'][part][index]
        assert p['text'].count(find)==1 and p['runText'].count(find)==1
        p['text']=p['text'].replace(find,replacement); p['runText']=p['runText'].replace(find,replacement)
        if kind=='table':
            table=next(t for t in expected['tables'] if t['part']==part and t['index']==spec['tableIndex'])
            assert table['rows'][spec['row']][spec['column']]==find
            table['rows'][spec['row']][spec['column']]=replacement
        else:
            # A paragraph can be inside a table or control. Derive the expected
            # projected container text from the original XML, never the output.
            with zipfile.ZipFile(source) as z:
                document=parse(z,part); selected=list(document.iter(W+'p'))[index]
                selected_text_nodes=set(selected.iter(W+'t'))
                def projected(node):
                    all_text_nodes=list(node.iter(W+'t'))
                    indices=[i for i,t in enumerate(all_text_nodes) if t in selected_text_nodes]
                    if not indices:return text(node)
                    start=sum(len(t.text or '') for t in all_text_nodes[:indices[0]])
                    original=text(node); paragraph_text=text(selected)
                    return original[:start]+paragraph_text.replace(find,replacement)+original[start+len(paragraph_text):]
                for table_index,table_xml in enumerate(document.iter(W+'tbl')):
                    table=next(t for t in expected['tables'] if t['part']==part and t['index']==table_index)
                    for row_index,row_xml in enumerate(table_xml.findall(W+'tr')):
                        for cell_index,cell_xml in enumerate(row_xml.findall(W+'tc')):
                            if selected in list(cell_xml.iter(W+'p')):table['rows'][row_index][cell_index]=projected(cell_xml)
                controls=[c for c in expected['controls'] if c['part']==part]
                for control,control_xml in zip(controls,document.iter(W+'sdt')):
                    if selected in list(control_xml.iter(W+'p')):control['text']=projected(control_xml)
    elif kind=='control':
        control=next(c for c in expected['controls'] if c['tag']==spec['tag']); old=control['text']; control['text']=spec['replace']
        # Control replacement can recreate runs. Compare its text and all other
        # paragraphs strictly, and retain the original SDT tag and alias.
        p=expected['stories']['word/document.xml'][spec['paragraphIndex']]; assert p['text']==old
        assert len(p['spans'])==1, 'Control fixture must have uniform original formatting'
        p['text']=spec['replace']; p['runText']=spec['replace']; p['spans']=[[0,len(spec['replace']),p['spans'][0][2]]]
    elif kind=='image':
        expected['images'][spec['imageIndex']]['alt']=spec['replace']
    else:
        assert kind=='copy'
    mismatches=[]
    for domain in expected:
        if after[domain]!=expected[domain]:
            mismatches.append(domain)
    if mismatches:
        # Bounded diagnostics without inflating large binary media.
        detail={}
        for domain in mismatches:
            if domain=='stories':
                detail[domain]=[]
                for part in expected[domain]:
                    for i,(a,b) in enumerate(zip(expected[domain][part],after[domain].get(part,[]))):
                        if a!=b:
                            detail[domain].append({'part':part,'index':i,'expected':a,'actual':b})
                            if len(detail[domain])>=3:break
                    if len(detail[domain])>=3:break
            else:detail[domain]={'expectedCount':len(expected[domain]),'actualCount':len(after[domain])}
        return {'passed':False,'mismatches':mismatches,'detail':detail}
    return {'passed':True,'sourceHash':digest_file(source),'savedHash':digest_file(saved),
        'scope':'Ordered paragraph text/run formatting, tables/grid/merges, control tag/alias/value, drawing position/size/alt/binary reference, story text, comment anchors/resolution, all binary-member hashes. Not full styles/layout or arbitrary Office objects.'}

if __name__=='__main__':
    mode=sys.argv[1]
    if mode=='inventory':result=inventory(sys.argv[2])
    elif mode=='snapshot':result=snapshot(sys.argv[2])
    else:
        with open(sys.argv[4]) as f:spec=json.load(f)
        result=verify(sys.argv[2],sys.argv[3],spec)
    print(json.dumps(result,ensure_ascii=False))
