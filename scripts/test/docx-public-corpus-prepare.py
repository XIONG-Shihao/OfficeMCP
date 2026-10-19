"""Bind unmodified independent public documents to exact applicable edit jobs."""
import json, os, sys
from docx_corpus_evidence import inventory, snapshot, digest_file
sources=[
 ('uk-model-contract','https://assets.publishing.service.gov.uk/media/68af256b969253904d1558e5/MSC_-_Core_Terms_-_E_W_-_Word_v2.2A_2025.DOCX','https://www.gov.uk/government/publications/the-model-services-contract-core-terms-england-wales','public-operational-contract-template'),
 ('au-tax-board-report','https://taxboard.gov.au/sites/taxboard.gov.au/files/2026-08/2025-26-bot-annual-report.docx','https://taxboard.gov.au/publications-and-media/publications','public-real-report'),
 ('alameda-rfp','https://www.alamedaca.gov/files/assets/public/v/1/it/arpd-rfp.docx','https://www.alamedaca.gov/BUSINESS/Bid-on-City-Contracts/Request-for-Proposal-Recreation-Management-Software-System','public-real-procurement'),
 ('alameda-contract','https://www.alamedaca.gov/files/assets/public/v/1/it/service-provider-agreement.docx','https://www.alamedaca.gov/BUSINESS/Bid-on-City-Contracts/Request-for-Proposal-Recreation-Management-Software-System','public-operational-contract-template')]
root,out=sys.argv[1:3];docs=[]
for name,url,page,classification in sources:
 file=os.path.abspath(os.path.join(root,name+'.docx'));s=snapshot(file);ps=s['stories']['word/document.xml'];texts=[p['text']for p in ps]
 candidates=[]
 for find,replacement in [('Supplier','Provider'),('software','platform'),('Board','Panel'),('2026','2027'),('City','Town')]:
  for i,p in enumerate(ps):
   if 20<=len(p['text'])<=1200 and p['text'].count(find)==1 and p['runText'].count(find)==1 and texts.count(p['text'])==1:
    candidates.append((i,p['text'],find,replacement))
 assert candidates,name
 index,phrase,find,replacement=candidates[0]
 jobs=[{'id':'text-edit','spec':{'kind':'text','paragraphIndex':index,'find':find,'replace':replacement},'select':{'type':'text','pattern':phrase},'instruction':f'In the unique paragraph containing {phrase!r}, replace {find!r} with {replacement!r}. This is a private test output, not a revision of the public original. Preserve its formatting and all other content.'},
 {'id':'comment-read','spec':{'kind':'copy'},'instruction':'Read existing comments and report their text/status or honestly report none. Save an unchanged copy; comments are read-only.'},
 {'id':'save-recovery','spec':{'kind':'copy'},'instruction':'Save an unchanged copy, recover a lost completed Save receipt by the same operation ID after restart, and avoid duplicate outputs.'}]
 docs.append({'id':name,'file':file,'classification':classification,'sourceHash':digest_file(file),'bytes':os.path.getsize(file),'provenance':{'url':url,'sourcePage':page,'downloaded':'2026-10-08','distribution':'Local downloaded evaluation corpus; not bundled into npm runtime'},'inventory':inventory(file),'jobs':jobs})
with open(out,'x')as f:json.dump({'schema':'officemcp.docx.corpus.v1','documents':docs},f,ensure_ascii=False,indent=2)
print(json.dumps({'manifest':out,'documents':[{'id':d['id'],'bytes':d['bytes'],'paragraphs':d['inventory']['paragraphs'],'features':d['inventory']['counts']}for d in docs]}))
