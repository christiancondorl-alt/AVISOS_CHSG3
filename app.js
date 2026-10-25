import { CONFIG } from './config.js';
import { DEMO_CATALOG } from './demo-catalog.js';
import { FIELD_IDS, validateFiles, validateValues } from './core.js';
import { GraphStore } from './graph.js';
const $=id=>document.getElementById(id), demo=CONFIG.mode==='demo';
let catalog,store,auth,account,reservation=null,key=null,busy=false,finished=false,locked=false,demoCounter=0;
let storageKey='';
function message(text,type=''){$('notice').textContent=text;$('notice').className='notice '+type;}
function select(id,values,value='') {
  const node=$(id);node.replaceChildren(new Option('Selecciona una opción',''));
  for(const option of values)node.add(new Option(option,option));
  node.value=values.includes(value)?value:'';node.disabled=!values.length;
}
function route() {$('route').textContent=[$('unidad').value,$('sistema').value,$('subsistema').value].filter(Boolean).join(' → ')||'La ruta del equipo aparecerá aquí.';}
function author(value=''){select('autor',catalog.autores[$('area').value]||[],value);}
function subsystem(value=''){select('subsistema',catalog.taxonomia[$('unidad').value]?.[$('sistema').value]||[],value);route();}
function system(value='',sub=''){select('sistema',Object.keys(catalog.taxonomia[$('unidad').value]||{}),value);subsystem(sub);}
function other(){const isOther=$('responsable').value==='Otras';$('other-wrap').hidden=!isOther;$('otra').required=isOther;}
function values(){return Object.fromEntries(FIELD_IDS.map(id=>[id,$(id).value]));}
function loadValues(v={}){
  for(const id of ['fecha','titulo','descripcion','consecuencia','otra']) if(v[id]!==undefined)$(id).value=v[id];
  select('area',Object.keys(catalog.autores),v.area);author(v.autor);
  select('unidad',Object.keys(catalog.taxonomia),v.unidad);system(v.sistema,v.subsistema);
  select('responsable',catalog.responsables,v.responsable);select('prioridad',catalog.prioridades,v.prioridad);other();
}
function today(){const parts=new Intl.DateTimeFormat('en-CA',{timeZone:'America/Lima',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date());const get=t=>parts.find(x=>x.type===t).value;return `${get('year')}-${get('month')}-${get('day')}`;}
function remember(){if(!demo&&storageKey)sessionStorage.setItem(storageKey,JSON.stringify({key,reservationId:reservation?.id,values:values()}));}
function controls(){
  $('reserve').disabled=busy||!!reservation||(!demo&&!account)||!catalog;
  $('fields').disabled=busy||!reservation||finished||locked;
  // Se permite volver a seleccionar evidencias durante un guardado incompleto.
  $('files').disabled=busy||!reservation||finished;
  if(locked&&!busy&&!finished){$('fields').disabled=false;for(const id of FIELD_IDS)$(id).disabled=true;}
  else if(!locked){for(const id of FIELD_IDS){const n=$(id);n.disabled=n.tagName==='SELECT'&&n.options.length<=1;}}
  $('save').disabled=busy||!reservation||finished;
  $('login').disabled=busy;$('logout').disabled=busy;
  $('reserve').textContent=reservation?'Número reservado':'Iniciar y reservar número';
  $('new').hidden=!finished;
  $('save').textContent=demo?'Probar registro (demostración)':locked?'Reintentar guardado':'Registrar aviso';
}
async function run(fn){if(busy)return;busy=true;controls();try{await fn();}catch(err){message(err.message||'No se pudo completar la operación.','error');}finally{busy=false;controls();}}
function links(item){
  const a=$('record-link');a.hidden=true;
  if(item.webUrl&&/^https:\/\/[^/]+\.sharepoint\.com\//i.test(item.webUrl)){a.href=item.webUrl;a.hidden=false;}
  $('result-files').replaceChildren();
  for(const file of JSON.parse(item.fields?.EvidenciasJson||'[]')){
    const li=document.createElement('li'),link=document.createElement('a');link.textContent=file.name;
    if(/^https:\/\/[^/]+\.sharepoint\.com\//i.test(file.url||'')){link.href=file.url;link.target='_blank';link.rel='noopener noreferrer';}
    li.append(link);$('result-files').append(li);
  }
}
function complete(item){finished=true;locked=false;$('result').hidden=false;$('result-title').textContent=demo?`Prueba completada · ${item.fields.NumeroAviso}`:`Aviso N.º ${item.fields.NumeroAviso} registrado`;$('result-text').textContent=demo?'Demostración: no se ha enviado ni guardado información en SharePoint.':'El aviso se guardó en SharePoint. No se ha asignado una OT.';links(item);message($('result-text').textContent,'success');if(storageKey)sessionStorage.removeItem(storageKey);}
function fromFields(f){
  const r=catalog.responsables.includes(f.AreaResponsable)?f.AreaResponsable:'Otras';
  return {fecha:f.FechaRegistro?.slice(0,10),area:f.AreaEmisora,autor:f.AutorAviso,unidad:f.Unidad,sistema:f.Sistema,subsistema:f.Subsistema,responsable:r,otra:r==='Otras'?f.AreaResponsable:'',titulo:f.Title,descripcion:f.Descripcion,consecuencia:f.Consecuencia,prioridad:f.Prioridad};
}
async function token(){try{return (await auth.acquireTokenSilent({account,scopes:CONFIG.scopes})).accessToken;}catch{throw new Error('La sesión requiere renovarse. Pulsa «Ingresar con Microsoft» y vuelve a intentar.');}}
async function connected(){
  storageKey=`sg3:${CONFIG.siteId}:${CONFIG.avisosListId}:${account.homeAccountId}`;
  reservation=null;finished=false;locked=false;catalog=null;$('aviso').value='Sin reservar';$('result').hidden=true;$('file-list').replaceChildren();$('form').reset();
  store=new GraphStore(CONFIG,token);catalog=await store.settings();loadValues({fecha:today()});
  $('account').textContent=account.username;$('logout').hidden=false;
  let saved;try{saved=JSON.parse(sessionStorage.getItem(storageKey)||'null');}catch{saved=null;}
  key=saved?.key||null;reservation=null;finished=false;locked=false;
  if(key){reservation=await store.find(key);if(reservation){reservation=await store.reserve(key,account.username);$('aviso').value=reservation.fields.NumeroAviso;locked=reservation.fields.Estado==='PendienteAdjuntos';loadValues(locked||reservation.fields.Estado==='Registrado'?fromFields(reservation.fields):saved.values);if(reservation.fields.Estado==='Registrado'){complete(reservation);return;}}}
  if(!reservation)$('aviso').value='Sin reservar';
  message(locked?'Este aviso tiene un guardado pendiente. Vuelve a seleccionar los archivos originales que falten y pulsa «Reintentar guardado».':reservation?'Se recuperó tu número reservado. Puedes continuar.':'Sesión conectada. Inicia el aviso para reservar su número en SharePoint.');
}
$('area').addEventListener('change',()=>{author();remember();});
$('unidad').addEventListener('change',()=>{system();remember();});
$('sistema').addEventListener('change',()=>{subsystem();remember();});
$('subsistema').addEventListener('change',route);$('responsable').addEventListener('change',other);
$('form').addEventListener('input',()=>{if(reservation&&!locked)remember();});
$('form').addEventListener('change',()=>{if(reservation&&!locked)remember();});
$('files').addEventListener('change',()=>{
  $('file-list').replaceChildren();try{const files=Array.from($('files').files);validateFiles(files);for(const f of files){const li=document.createElement('li');li.textContent=`${f.name} · ${(f.size/1024/1024).toFixed(2)} MB`;$('file-list').append(li);}}catch(e){$('files').value='';message(e.message,'error');}
});
$('reserve').addEventListener('click',()=>run(async()=>{
  if(demo){reservation={id:String(++demoCounter),fields:{NumeroAviso:`DEMO-${String(demoCounter).padStart(4,'0')}`}};message('Número de demostración. Esta prueba no reserva un número real.');}
  else{key=key||crypto.randomUUID();remember();message('Reservando número en SharePoint…');reservation=await store.reserve(key,account.username);message(`N.º de aviso ${reservation.fields.NumeroAviso} reservado. Completa los campos.`);}
  $('aviso').value=reservation.fields.NumeroAviso;remember();
}));
$('form').addEventListener('submit',event=>{
  event.preventDefault();if(!reservation||finished||busy)return;
  run(async()=>{
    const files=Array.from($('files').files),v=values();validateFiles(files);if(!locked)validateValues(v,catalog);remember();
    if(demo){complete(reservation);return;}
    message('Guardando aviso…');
    try{const item=await store.save(reservation.id,v,files,t=>message(t));complete(item);}
    catch(e){try{const item=await store.get(reservation.id);if(item.fields.Estado==='Registrado'){complete(item);return;}locked=item.fields.Estado==='PendienteAdjuntos';}catch{}throw e;}
  });
});
$('new').addEventListener('click',()=>{
  reservation=null;key=null;finished=false;locked=false;$('form').reset();loadValues({fecha:today()});$('aviso').value='Sin reservar';$('result').hidden=true;$('file-list').replaceChildren();controls();message(demo?'Demostración: inicia un aviso de prueba.':'Inicia un nuevo aviso para reservar el siguiente número.');
});
$('login').addEventListener('click',()=>run(async()=>{
  remember();const response=await auth.loginPopup({scopes:CONFIG.scopes,prompt:'select_account'});account=response.account;auth.setActiveAccount(account);await connected();
}));
$('logout').addEventListener('click',()=>run(async()=>{
  remember();await auth.logoutPopup({account,postLogoutRedirectUri:new URL('./',location.href).href});location.reload();
}));
async function boot(){
  if(demo){catalog=DEMO_CATALOG;loadValues({fecha:today()});$('mode').textContent='DEMOSTRACIÓN';$('account').textContent='Vista de prueba · autores ficticios';message('Modo demostración: explora el formulario. No se conecta ni guarda en SharePoint.');controls();return;}
  $('mode').textContent='SHAREPOINT';
  if(CONFIG.mode!=='sharepoint')throw new Error('Modo desconocido en config.js.');
  for(const k of ['tenantId','clientId','siteId','avisosListId','configuracionListId','configuracionItemId','evidenciasDriveId','redirectUri'])if(!CONFIG[k])throw new Error(`Completa «${k}» en config.js siguiendo LEEME.md.`);
  if(!/^[\da-f-]{36}$/i.test(CONFIG.tenantId)||!/^[\da-f-]{36}$/i.test(CONFIG.clientId))throw new Error('tenantId y clientId deben ser los identificadores GUID de Microsoft Entra.');
  if(new URL(CONFIG.redirectUri).origin!==location.origin)throw new Error('La URI de retorno debe pertenecer a este sitio.');
  const {PublicClientApplication}=await import('./vendor/msal.js');
  auth=new PublicClientApplication({auth:{clientId:CONFIG.clientId,authority:`https://login.microsoftonline.com/${CONFIG.tenantId}`,redirectUri:CONFIG.redirectUri},cache:{cacheLocation:'sessionStorage'}});
  await auth.initialize();$('login').hidden=false;
  message('Ingresa con tu cuenta corporativa para obtener el catálogo y reservar un número.');controls();
}
boot().catch(e=>{message(e.message,'error');controls();});
