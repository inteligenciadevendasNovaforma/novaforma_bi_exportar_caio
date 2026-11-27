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
const xml=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const numberTyped=v=>typeof v==="number"&&Number.isFinite(v)?v:null;
const column=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
export async function makeXlsx(sheets){
 const files={},shared=[],indexes=new Map();let strings=0;
 const stringIndex=value=>{const s=String(value??'');strings++;if(!indexes.has(s)){indexes.set(s,shared.length);shared.push(s);}return indexes.get(s);};
 for(const [sheetIndex,sheet] of sheets.entries()){
  const offset=sheet.title?1:0,rows=[...(sheet.title?[[sheet.title]]:[]),sheet.headers,...sheet.rows];const content=rows.map((row,ri)=>'<row r="'+(ri+1)+'">'+row.map((value,ci)=>{const reference=column(ci)+(ri+1),n=ri>offset&&sheet.numeric?.[ci]?numberTyped(value):null;if(n!==null)return `<c r="${reference}" s="2"><v>${n}</v></c>`;return `<c r="${reference}" t="s"${ri<=offset?' s="1"':''}><v>${stringIndex(value)}</v></c>`;}).join('')+'</row>').join('');
  files[`xl/worksheets/sheet${sheetIndex+1}.xml`]=strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="'+(offset+1)+'" topLeftCell="A'+(offset+2)+'" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>'+sheet.headers.map((_,i)=>`<col min="${i+1}" max="${i+1}" width="${i===0?18:24}" customWidth="1"/>`).join('')+'</cols><sheetData>'+content+'</sheetData>'+(sheet.title&&sheet.headers.length>1?'<mergeCells count="1"><mergeCell ref="A1:'+column(sheet.headers.length-1)+'1"/></mergeCells>':'')+'</worksheet>');
 }
 files['xl/sharedStrings.xml']=strToU8('<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'+strings+'" uniqueCount="'+shared.length+'">'+shared.map(s=>'<si><t xml:space="preserve">'+xml(s)+'</t></si>').join('')+'</sst>');
 files['xl/styles.xml']=strToU8('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF052B5F"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
 files['xl/workbook.xml']=strToU8('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+sheets.map((s,i)=>`<sheet name="${xml(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')+'</sheets></workbook>');
 files['xl/_rels/workbook.xml.rels']=strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')+`<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId${sheets.length+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
 files['_rels/.rels']=strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
 files['[Content_Types].xml']=strToU8('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'+sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')+'</Types>');
 return zipManual(files);
}
export function base64(bytes){let binary='';for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));return btoa(binary);}
