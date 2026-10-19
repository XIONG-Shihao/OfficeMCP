/** Bounded JSON read payloads; each distribution ships this identical helper. */
function stable(value) {
  if(Array.isArray(value))return value.map(stable)
  if(value&&typeof value==='object')return Object.fromEntries(Object.keys(value).sort().map(key=>[key,stable(value[key])]))
  return value
}
export class DocxReadCache {
  constructor({maximumBytes=1024*1024,maximumEntryBytes=256*1024,maximumEntries=32}={}) {
    for(const n of [maximumBytes,maximumEntryBytes,maximumEntries])if(!Number.isSafeInteger(n)||n<1)throw Error('Invalid DOCX read cache limit')
    Object.assign(this,{maximumBytes,maximumEntryBytes,maximumEntries});this.entries=new Map();this.bytes=0;this.hits=0;this.misses=0
  }
  key(namespace,args){return JSON.stringify([namespace,stable(args)])}
  get(namespace,args) {
    const key=this.key(namespace,args),row=this.entries.get(key)
    if(!row){this.misses++;return undefined}
    this.entries.delete(key);this.entries.set(key,row);this.hits++
    return JSON.parse(row.payload)
  }
  set(namespace,args,value) {
    const key=this.key(namespace,args),payload=JSON.stringify(value)
    this.delete(key)
    if(payload===undefined)return false
    const bytes=Buffer.byteLength(key)+Buffer.byteLength(payload)
    if(bytes>this.maximumEntryBytes||bytes>this.maximumBytes)return false
    while(this.entries.size>=this.maximumEntries||this.bytes+bytes>this.maximumBytes)this.delete(this.entries.keys().next().value)
    this.entries.set(key,{namespace,payload,bytes});this.bytes+=bytes;return true
  }
  delete(key){const row=this.entries.get(key);if(row){this.bytes-=row.bytes;this.entries.delete(key)}}
  clear(namespace){if(namespace===undefined){this.entries.clear();this.bytes=0;return}for(const [key,row]of this.entries)if(row.namespace===namespace)this.delete(key)}
  stats(){return{entries:this.entries.size,bytes:this.bytes,hits:this.hits,misses:this.misses}}
}
