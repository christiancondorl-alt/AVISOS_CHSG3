export const FIELD_IDS = ['fecha','area','autor','unidad','sistema','subsistema','responsable','otra','titulo','descripcion','consecuencia','prioridad'];
export function numberFor(base,id) {
  const n=Number(id);
  if(!Number.isSafeInteger(base)||base<0||!Number.isSafeInteger(n)||n<1||!Number.isSafeInteger(base+n)) throw new Error('Configuración de numeración inválida.');
  return String(base+n);
}
export function fileKey(f) { return `${f.name}|${f.size}|${f.lastModified}`; }
export function validateFiles(files) {
  if(files.length>2) throw new Error('Selecciona como máximo 2 archivos.');
  const allowed=/\.(jpg|jpeg|png|webp|gif|bmp|heic|heif|tif|tiff|doc|docx|xls|xlsx|xlsm|ppt|pptx|pdf|mp4|mov|avi|wmv|m4v|mp3|wav|m4a|ogg|aac)$/i;
  const seen=new Set();
  for(const f of files) {
    if(f.size>100*1024*1024) throw new Error(`${f.name}: supera 100 MB.`);
    if(!f.size) throw new Error(`${f.name}: el archivo está vacío.`);
    if(!allowed.test(f.name)) throw new Error(`${f.name}: formato no permitido.`);
    if(seen.has(fileKey(f))) throw new Error('El mismo archivo aparece dos veces.');
    seen.add(fileKey(f));
  }
}
export function validateValues(v,c) {
  for(const id of FIELD_IDS.filter(x=>x!=='otra')) if(!v[id]?.trim()) throw new Error('Completa todos los campos obligatorios.');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(v.fecha)||Number.isNaN(Date.parse(v.fecha+'T12:00:00Z'))) throw new Error('Fecha inválida.');
  if(!c.autores[v.area]?.includes(v.autor)) throw new Error('Selecciona un autor del área indicada.');
  if(!c.taxonomia[v.unidad]?.[v.sistema]?.includes(v.subsistema)) throw new Error('La ruta del equipo no es válida.');
  if(!c.responsables.includes(v.responsable)||!c.prioridades.includes(v.prioridad)) throw new Error('Responsable o prioridad inválidos.');
  if(v.responsable==='Otras'&&!v.otra?.trim()) throw new Error('Especifica el área responsable.');
  for(const k of ['titulo','consecuencia','otra']) if(v[k]?.length>255) throw new Error('Un campo corto supera 255 caracteres.');
  if(v.descripcion.length>12000) throw new Error('La descripción supera 12 000 caracteres.');
}
export function fieldsFrom(v) {
  return {Title:v.titulo.trim(),FechaRegistro:v.fecha+'T12:00:00Z',AreaEmisora:v.area,AutorAviso:v.autor,Unidad:v.unidad,Sistema:v.sistema,Subsistema:v.subsistema,AreaResponsable:v.responsable==='Otras'?v.otra.trim():v.responsable,Descripcion:v.descripcion.trim(),Consecuencia:v.consecuencia.trim(),Prioridad:v.prioridad};
}
export function safeFileName(name) {return name.replace(/["*:<>?\/\\|#%\x00-\x1f]/g,'_').replace(/[. ]+$/g,'').slice(0,140)||'archivo';}
