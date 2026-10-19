"""Select real bid excerpts without adding synthetic content. Preserve sources."""
import copy, hashlib, json, os, posixpath, sys, zipfile
import xml.etree.ElementTree as E
from docx_corpus_evidence import inventory, snapshot, digest_file, W, R, PR, text

def excerpt(source, output, start, end):
    with zipfile.ZipFile(source) as z:
        d=E.fromstring(z.read('word/document.xml'));body=d.find(W+'body');children=list(body)
        selected=[copy.deepcopy(n) for n in children[start:end] if n.tag!=W+'sectPr']
        section=body.find(W+'sectPr')
        body.clear();body.extend(selected)
        if section is not None:body.append(copy.deepcopy(section))
        modified={'word/document.xml':E.tostring(d,encoding='utf-8',xml_declaration=True)}
        refs={v for n in d.iter() for k,v in n.attrib.items() if k in [R+'id',R+'embed',R+'link']}
        name='word/_rels/document.xml.rels';rels=E.fromstring(z.read(name))
        # Drop only unreferenced media/embedding relationships. Keep styles,
        # numbering, themes and other document support parts regardless of IDs.
        for rel in list(rels):
            if rel.get('Id') not in refs and rel.get('Type','').rsplit('/',1)[-1] in ['image','oleObject','package']:
                rels.remove(rel)
        modified[name]=E.tostring(rels,encoding='utf-8',xml_declaration=True)
        names=set(z.namelist());keep=set();queue=['[Content_Types].xml','_rels/.rels']
        while queue:
            n=queue.pop()
            if n in keep or n not in names:continue
            keep.add(n)
            relfile=n if n.endswith('.rels') else posixpath.join(posixpath.dirname(n),'_rels',posixpath.basename(n)+'.rels')
            if relfile in names:
                keep.add(relfile);root=E.fromstring(modified.get(relfile,z.read(relfile)))
                base='' if relfile=='_rels/.rels' else posixpath.dirname(posixpath.dirname(relfile))
                for rel in root:
                    if rel.get('TargetMode')=='External':continue
                    target=rel.get('Target','');resolved=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base,target))
                    queue.append(resolved)
        types=E.fromstring(z.read('[Content_Types].xml'))
        for n in list(types):
            if n.tag.endswith('Override') and n.get('PartName','').lstrip('/') not in keep:types.remove(n)
        modified['[Content_Types].xml']=E.tostring(types,encoding='utf-8',xml_declaration=True)
        with zipfile.ZipFile(output,'x',compression=zipfile.ZIP_DEFLATED) as out:
            for name in sorted(keep):out.writestr(name,modified.get(name,z.read(name)))
    return {'source':source,'sourceHash':digest_file(source),'bodyElementRange':[start,end],
        'method':'Select original body elements and section properties; prune unreferenced media relationships and unreachable package parts. No new task text.'}

def case(id,file,classification,provenance):
    s=snapshot(file); inv=inventory(file);ps=s['stories']['word/document.xml']; jobs=[]
    # Exact real phrase, not a marker inserted for the test.
    phrase=next((p['text'] for p in ps if '2、供应商将按招标文件的规定履行合同责任和义务。'==p['text']),None)
    if phrase:
        find,replace='义务','职责';index=next(i for i,p in enumerate(ps)if p['text']==phrase)
    elif any('我方报价真实有效！' in p['text'] for p in ps):
        index=next(i for i,p in enumerate(ps)if '我方报价真实有效！' in p['text'])
        phrase=ps[index]['text'];find,replace='有效','准确'
    elif any('January 15, 2025' in p['text'] for p in ps):
        index=next(i for i,p in enumerate(ps)if 'January 15, 2025' in p['text'])
        phrase=ps[index]['text'];find,replace='2025','2026'
    if phrase:
        assert sum(p['text'].count(phrase) for p in ps)==1
        jobs.append({'id':'text-edit','spec':{'kind':'text','paragraphIndex':index,'find':find,'replace':replace},'select':{'type':'text','pattern':phrase},
        'instruction':f'In the unique paragraph containing {phrase!r}, replace {find!r} with {replace!r}. Preserve formatting and all other content.'})
    jobs.append({'id':'comment-read','spec':{'kind':'copy'},
        'instruction':'Read existing comments, report their text and open/resolved status (or honestly report none), then save an unchanged copy. Do not create, edit, reply to, delete or resolve comments.'})
    jobs.append({'id':'save-recovery','spec':{'kind':'copy'},'instruction':'Save an unchanged copy. Recover a completed save receipt by its operation ID if its response is unavailable; never create a second output.'})
    for t in s['tables']:
        if t['part']!='word/document.xml' or not t['rows'] or '单价（元）' not in t['rows'][0]:continue
        column=t['rows'][0].index('单价（元）');row=next(i for i,r in enumerate(t['rows'])if len(r)>column and r[column]=='10.50')
        # Map the selected physical cell paragraph to the independent story order.
        with zipfile.ZipFile(file) as z:
            d=E.fromstring(z.read('word/document.xml'));table=list(d.iter(W+'tbl'))[t['index']];cell=table.findall(W+'tr')[row].findall(W+'tc')[column];p=next(cell.iter(W+'p'));pindex=list(d.iter(W+'p')).index(p)
        key=t['rows'][row][0];header=t['rows'][0][column]
        jobs.append({'id':'table-edit','spec':{'kind':'table','tableIndex':t['index'],'row':row,'column':column,'paragraphIndex':pindex,'find':'10.50','replace':'10.55'},
            'select':{'type':'table','header':'单价（元）'},'keyColumn':t['rows'][0][0],'key':key,'columnHeader':header,'merged':bool(t['merges']),
            'instruction':f'In the table with column {header!r}, change the price for row key {key!r} from 10.50 to 10.55. Identify the row and column by their text and leave every other cell unchanged. If the merged table makes the by-key workflow unsupported, use a supported guarded cell action.'})
        break
    for control in s['controls']:
        if control['tag']=='company-name':
            ci=next(i for i,p in enumerate(ps)if p['text']==control['text'])
            jobs.append({'id':'template-control','spec':{'kind':'control','tag':control['tag'],'paragraphIndex':ci,'replace':'Meridian Labs'},
                'instruction':'Set the text content control tagged company-name to Meridian Labs. Keep the control tag and alias and all other content.'})
    return {'id':id,'file':os.path.abspath(file),'classification':classification,'sourceHash':digest_file(file),'bytes':os.path.getsize(file),'provenance':provenance,'inventory':inv,'jobs':jobs}

if __name__=='__main__':
    source,public_root,out=sys.argv[1:4];os.makedirs(out,exist_ok=True)
    business=os.path.join(out,'bid-business.docx');pricing=os.path.join(out,'bid-pricing.docx')
    bp=excerpt(source,business,0,46);pp=excerpt(source,pricing,46,81)
    corpus={'schema':'officemcp.docx.corpus.v1','original':os.path.abspath(source),'originalHash':digest_file(source),'documents':[
        case('bid-business',business,'real-derived-private',bp),case('bid-pricing',pricing,'real-derived-private',pp),
        case('bid-stress',source,'real-private',{'method':'Unmodified supplied real bid; also the 234 MB stress fixture'}),
        case('public-nda',os.path.join(public_root,'apps/docs/public/fixtures/sample-nda.docx'),'public-synthetic',{'method':'Upstream sanitized Word-style NDA; README identifies it as synthetic and redistributable'}),
        case('public-controls',os.path.join(public_root,'examples/content-controls/public/content-control.docx'),'public-synthetic',{'method':'Upstream Acme/Amazing Corp NDA content-control example; supplemental feature fixture'}),
        case('public-comments',os.path.join(public_root,'packages/superdoc/tests/cdn-smoke/fixtures/sample-review.docx'),'public-synthetic',{'method':'Upstream review memo with existing comments and tracked changes; read/preservation coverage only'})]}
    with open(os.path.join(out,'corpus.json'),'x')as f:json.dump(corpus,f,ensure_ascii=False,indent=2)
    print(json.dumps({'manifest':os.path.abspath(os.path.join(out,'corpus.json')),'documents':[{'id':d['id'],'bytes':d['bytes'],'classification':d['classification'],'jobs':[j['id']for j in d['jobs']]}for d in corpus['documents']]}))
