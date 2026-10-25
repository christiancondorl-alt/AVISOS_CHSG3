import { numberFor, fieldsFrom, fileKey, safeFileName } from './core.js';
const ROOT='https://graph.microsoft.com/v1.0';
export class GraphStore {
  constructor(config, token, fetchImpl=globalThis.fetch.bind(globalThis)) {this.c=config;this.token=token;this.fetch=fetchImpl;this.base=`/sites/${encodeURIComponent(config.siteId)}/lists/${encodeURIComponent(config.avisosListId)}`;}
  async request(path,options={}) {
    const url=path.startsWith('https:')?path:ROOT+path;
    if(!url.startsWith(ROOT+'/')) throw new Error('Destino de API no permitido.');
    const token=await this.token();
    const headers={Authorization:`Bearer ${token}`,...options.headers};
    if(options.body && typeof options.body==='string') headers['Content-Type']='application/json';
    const response=await this.fetch(url,{...options,headers});
    if(!response.ok) {
      let detail='';try{detail=(await response.json()).error?.message||'';}catch{}
      const err=new Error(`SharePoint (${response.status}): ${detail||response.statusText}`);err.status=response.status;throw err;
    }
    return response.status===204?null:response.json();
  }
  async settings() {
    const x=await this.request(`/sites/${encodeURIComponent(this.c.siteId)}/lists/${encodeURIComponent(this.c.configuracionListId)}/items/${encodeURIComponent(this.c.configuracionItemId)}?$expand=fields`);
    const base=x.fields.UltimoIdHistorico;
    if(base===undefined||base===null) throw new Error('Falta UltimoIdHistorico en la configuración central.');
    numberFor(Number(base),1);
    const catalog=JSON.parse(x.fields.CatalogoJson||'null');
    if(!catalog?.autores||!catalog?.taxonomia||!catalog?.prioridades||!catalog?.responsables) throw new Error('Catálogo central incompleto.');
    this.baseHistorica=Number(base);return catalog;
  }
  get(id) { return this.request(`${this.base}/items/${encodeURIComponent(id)}?$expand=fields`); }
  patch(id,fields,etag) {return this.request(`${this.base}/items/${encodeURIComponent(id)}/fields`,{method:'PATCH',headers:etag?{'If-Match':etag}:{},body:JSON.stringify(fields)});}
  async find(key) {
    if(!/^[0-9a-f-]{36}$/i.test(key)) throw new Error('Clave de registro inválida.');
    const filter=encodeURIComponent(`fields/ClaveRegistro eq '${key}'`);
    const data=await this.request(`${this.base}/items?$expand=fields&$filter=${filter}`);
    return data.value?.[0]||null;
  }
  async reserve(key,username) {
    // POST no se reintenta a ciegas: ClaveRegistro es única e indexada.
    let item=await this.find(key);
    if(!item) {
      try {item=await this.request(`${this.base}/items`,{method:'POST',body:JSON.stringify({fields:{Title:'Aviso reservado',ClaveRegistro:key,NumeroAviso:`R-${key}`,Estado:'Borrador',BaseHistorica:this.baseHistorica,RegistradoPor:username}})});}
      catch(err) {item=await this.find(key);if(!item)throw err;}
    }
    item=await this.get(item.id);
    const number=numberFor(Number(item.fields.BaseHistorica),item.id);
    if(item.fields.NumeroAviso&&item.fields.NumeroAviso!==`R-${key}`&&item.fields.NumeroAviso!==number) throw new Error('El número almacenado no coincide con la base del registro. Contacta al administrador.');
    if(!item.fields.NumeroAviso||item.fields.NumeroAviso===`R-${key}`) await this.patch(item.id,{NumeroAviso:number},item.eTag);
    return this.get(item.id);
  }
  async save(id,values,files,onProgress=()=>{}) {
    let item=await this.get(id),f=item.fields;
    if(f.Estado==='Registrado') return item;
    if(f.Estado!=='Borrador'&&f.Estado!=='PendienteAdjuntos') throw new Error('El aviso no admite guardado en su estado actual.');
    let manifest=f.ManifestJson?JSON.parse(f.ManifestJson):null;
    let uploaded=f.EvidenciasJson?JSON.parse(f.EvidenciasJson):[];
    if(!manifest) {
      manifest=files.map((x,i)=>({key:fileKey(x),name:x.name,size:x.size,remoteName:`${id}-${i+1}-${safeFileName(x.name)}`}));
      await this.patch(id,{...fieldsFrom(values),Estado:'PendienteAdjuntos',ManifestJson:JSON.stringify(manifest),EvidenciasJson:'[]'},item.eTag);
    }
    for(const expected of manifest) {
      if(uploaded.some(x=>x.key===expected.key)) continue;
      const file=files.find(x=>fileKey(x)===expected.key);
      if(!file) throw new Error(`Vuelve a seleccionar el archivo original pendiente: ${expected.name}. El aviso conserva su número.`);
      onProgress(`Subiendo ${expected.name}…`);
      const result=await this.request(`/drives/${encodeURIComponent(this.c.evidenciasDriveId)}/root:/${encodeURIComponent(expected.remoteName)}:/content`,{method:'PUT',headers:{'Content-Type':file.type||'application/octet-stream'},body:file});
      uploaded.push({key:expected.key,name:expected.name,id:result.id,url:result.webUrl});
      // El mismo nombre remoto hace que reintentar un PUT reemplace el archivo, no lo duplique.
      item=await this.get(id);
      if(item.fields.Estado==='Registrado')return item;
      if(item.fields.Estado!=='PendienteAdjuntos')throw new Error('El estado cambió durante la carga.');
      await this.patch(id,{EvidenciasJson:JSON.stringify(uploaded)},item.eTag);
    }
    item=await this.get(id);
    if(item.fields.Estado==='Registrado')return item;
    if(item.fields.Estado!=='PendienteAdjuntos')throw new Error('El estado cambió antes de finalizar.');
    await this.patch(id,{Estado:'Registrado'},item.eTag);
    return this.get(id);
  }
}
