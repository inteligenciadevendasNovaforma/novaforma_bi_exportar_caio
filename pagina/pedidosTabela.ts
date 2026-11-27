import {makeXlsx} from './pedidosTabelaXlsx.mjs';

export const MAX_PEDIDOS_BYTES=30*1024*1024;
const MAX_COMPRESSED_BYTES=MAX_PEDIDOS_BYTES+65536;
export const ERRO_PEDIDOS='Não foi possível preparar o arquivo. Gere a exportação novamente.';
export const ERRO_PEDIDOS_GRANDE='A exportação é muito grande. Reduza os filtros e tente novamente.';
export interface FolhaTabela {name:string;headers:string[];rows:(string|number|boolean|null)[][];numeric:boolean[];title?:string;}
export interface CargaTabela {tipo:'pedidosTabela';versao:1;titulo:string;periodos?:unknown;filtros?:unknown;S:FolhaTabela[];}
const fail=():never=>{throw new Error(ERRO_PEDIDOS);};
const text=(value:unknown):value is string=>typeof value==='string'&&value.length<=32767;
function validarCarga(value:unknown):CargaTabela {
 const c=value as CargaTabela;
 if(!c||c.tipo!=='pedidosTabela'||c.versao!==1||!text(c.titulo)||!Array.isArray(c.S)||!c.S.length||c.S.length>100)fail();
 for(const s of c.S){
  if(!s||!text(s.name)||!Array.isArray(s.headers)||!s.headers.length||s.headers.length>16384||!s.headers.every(text)||!Array.isArray(s.numeric)||s.numeric.length!==s.headers.length||!s.numeric.every(n=>typeof n==='boolean')||!Array.isArray(s.rows)||s.rows.length>1048574||(s.title!==undefined&&!text(s.title)))fail();
  for(const row of s.rows){if(!Array.isArray(row)||row.length!==s.headers.length)fail();for(const cell of row)if(cell!==null&&!(text(cell)||(typeof cell==='number'&&Number.isFinite(cell))||typeof cell==='boolean'))fail();}
 }
 return c;
}
export async function decodificarPedidosTabela(encoded:string,compressed=true):Promise<CargaTabela>{
 if(!encoded||encoded.length>Math.ceil(MAX_COMPRESSED_BYTES/3)*4)throw new Error(ERRO_PEDIDOS_GRANDE);
 if(!/^[A-Za-z0-9_-]+$/.test(encoded))fail();
 const raw=atob(encoded.replace(/-/g,'+').replace(/_/g,'/').padEnd(Math.ceil(encoded.length/4)*4,'='));
 const input=Uint8Array.from(raw,c=>c.charCodeAt(0));if(input.length>MAX_COMPRESSED_BYTES)throw new Error(ERRO_PEDIDOS_GRANDE);
 let bytes:Uint8Array;
 if(compressed){
  const reader=new Blob([input as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();const chunks:Uint8Array[]=[];let total=0;
  try{while(true){const part=await reader.read();if(part.done)break;total+=part.value.length;if(total>MAX_PEDIDOS_BYTES){await reader.cancel();throw new Error(ERRO_PEDIDOS_GRANDE);}chunks.push(part.value);}}finally{reader.releaseLock();}
  bytes=new Uint8Array(total);let offset=0;for(const part of chunks){bytes.set(part,offset);offset+=part.length;}
 }else bytes=input;
 if(bytes.length>MAX_PEDIDOS_BYTES)throw new Error(ERRO_PEDIDOS_GRANDE);
 return validarCarga(JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(bytes)));
}
function contexto(c:CargaTabela):(string|number|boolean|null)[][]{
 const rows:(string|number|boolean|null)[][]=[['Título',c.titulo]];
 const add=(label:string,value:unknown,depth=0):void=>{
  if(depth>12)fail();
  if(value===undefined)return;
  if(value===null||typeof value!=='object'){if(!text(label)||!(value===null||text(value)||typeof value==='number'||typeof value==='boolean'))fail();rows.push([label,value as string|number|boolean|null]);return;}
  for(const [key,v] of Object.entries(value))add(label+' / '+key,v,depth+1);
 };
 add('Períodos',c.periodos);add('Filtros',c.filtros);return rows;
}
export async function gerarXlsxTabela(value:unknown):Promise<Uint8Array>{
 const c=validarCarga(value),used=new Set<string>();
 const sheets=c.S.map((s,i)=>{
  const base=(s.name.replace(/[\\/?*\[\]:\x00-\x1f]/g,'_').replace(/^'+|'+$/g,'').trim()||'Planilha '+(i+1)).slice(0,31);let name=base,n=2;
  while(used.has(name.toLocaleLowerCase())){const suffix=' ('+(n++)+')';name=base.slice(0,31-suffix.length)+suffix;}used.add(name.toLocaleLowerCase());
  // Optional title means no extra row when absent, matching the native writer.
  return {...s,name};
 });
 if(!used.has('contexto')&&(c.periodos!==undefined||c.filtros!==undefined))sheets.push({name:'Contexto',headers:['Campo','Valor'],numeric:[false,false],rows:contexto(c),title:c.titulo});
 const bytes=await makeXlsx(sheets);if(bytes.length>MAX_PEDIDOS_BYTES)throw new Error(ERRO_PEDIDOS_GRANDE);return bytes;
}
export function mensagemPedidos(error:unknown):string{return error instanceof Error&&error.message===ERRO_PEDIDOS_GRANDE?ERRO_PEDIDOS_GRANDE:ERRO_PEDIDOS;}
