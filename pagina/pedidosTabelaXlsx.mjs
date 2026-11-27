import {buildXlsxFiles} from './pedidosTabelaLayout.mjs';
// Standalone ZIP: local headers, CRC32, central directory and EOCD.
// This new JSON path has no runtime library imports; legacy writers are intact.
const strToU8=value=>new TextEncoder().encode(value);
const crcTable=Uint32Array.from({length:256},(_,i)=>{let c=i;for(let j=0;j<8;j++)c=c&1?0xedb88320^(c>>>1):c>>>1;return c>>>0;});
const crc32=bytes=>{let c=0xffffffff;for(const b of bytes)c=crcTable[(c^b)&255]^(c>>>8);return (c^0xffffffff)>>>0;};
async function zipManual(files){
 const entries=[];
 for(const [name,raw] of Object.entries(files)){
  // Compress each XML entry with the browser API so complete large exports
  // remain practical without embedding a compression library in the new path.
  const stream=new Blob([raw]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  const data=new Uint8Array(await new Response(stream).arrayBuffer());
  entries.push({name:strToU8(name),data,size:raw.length,crc:crc32(raw)});
 }
 const localSize=entries.reduce((n,e)=>n+30+e.name.length+e.data.length,0),centralSize=entries.reduce((n,e)=>n+46+e.name.length,0),total=localSize+centralSize+22;
 if(total>30*1024*1024)throw new Error('A exportação é muito grande. Reduza os filtros e tente novamente.');
 const bytes=new Uint8Array(total),view=new DataView(bytes.buffer);let offset=0;
 const u16=(o,n)=>view.setUint16(o,n,true),u32=(o,n)=>view.setUint32(o,n,true);
 for(const e of entries){e.offset=offset;u32(offset,0x04034b50);u16(offset+4,20);u16(offset+6,0x800);u16(offset+8,8);u16(offset+12,33);u32(offset+14,e.crc);u32(offset+18,e.data.length);u32(offset+22,e.size);u16(offset+26,e.name.length);bytes.set(e.name,offset+30);bytes.set(e.data,offset+30+e.name.length);offset+=30+e.name.length+e.data.length;}
 const centralOffset=offset;
 for(const e of entries){u32(offset,0x02014b50);u16(offset+4,20);u16(offset+6,20);u16(offset+8,0x800);u16(offset+10,8);u16(offset+14,33);u32(offset+16,e.crc);u32(offset+20,e.data.length);u32(offset+24,e.size);u16(offset+28,e.name.length);u32(offset+42,e.offset);bytes.set(e.name,offset+46);offset+=46+e.name.length;}
 u32(offset,0x06054b50);u16(offset+8,entries.length);u16(offset+10,entries.length);u32(offset+12,centralSize);u32(offset+16,centralOffset);return bytes;
}
const numberTyped=v=>typeof v==="number"&&Number.isFinite(v)?v:null;
export async function makeXlsx(sheets){return zipManual(buildXlsxFiles(sheets,numberTyped));}
export function base64(bytes){let binary='';for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));return btoa(binary);}
