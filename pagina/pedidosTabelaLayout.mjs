// Shared XML layout. The native writer embeds this exact function so both
// independently buildable packages keep identical cells, styles and KPI rows.
export function buildXlsxFiles(sheets,numberValue){
 const xml=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
 const column=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
 const utf8=v=>new TextEncoder().encode(v),files={},shared=[],indexes=new Map();let strings=0;
 const stringIndex=value=>{const s=String(value??'');strings++;if(!indexes.has(s)){indexes.set(s,shared.length);shared.push(s);}return indexes.get(s);};
 const cell=(value,ci,ri,style,numeric=false)=>{const ref=column(ci)+ri,n=numeric?numberValue(value):null;return n!==null?`<c r="${ref}" s="${style}"><v>${n}</v></c>`:`<c r="${ref}" s="${style}" t="s"><v>${stringIndex(value)}</v></c>`;};
 const totals=sheets.find(s=>String(s.name).trim().toLowerCase()==='totais do filtro');
 const kpis=(totals?.rows||[]).filter(r=>r.length>=2&&String(r[0]??'').trim()).map(r=>[r[0],r[1]]);
 for(const [sheetIndex,sheet] of sheets.entries()){
  const matrix=['pedidos','materiais'].includes(String(sheet.name).trim().toLowerCase()),width=Math.max(1,sheet.headers.length),offset=sheet.title?1:0,headerRow=offset+1,lastData=headerRow+sheet.rows.length,rows=[],merges=[];
  const row=(ri,content,height)=>`<row r="${ri}"${height?` ht="${height}" customHeight="1"`:''}>${content}</row>`;
  const merge=(ri,start,end)=>{if(end>start)merges.push(`${column(start)}${ri}:${column(end)}${ri}`);};
  if(sheet.title){rows.push(row(1,cell(sheet.title,0,1,1),26));merge(1,0,width-1);}
  rows.push(row(headerRow,sheet.headers.map((v,i)=>cell(v,i,headerRow,1)).join(''),30));
  for(const [i,values] of sheet.rows.entries()){
   const ri=headerRow+i+1,zebra=matrix&&i%2===1;
   rows.push(row(ri,values.map((value,ci)=>cell(value,ci,ri,sheet.numeric?.[ci]?(zebra?4:2):(zebra?3:0),!!sheet.numeric?.[ci])).join(''),matrix?18:undefined));
  }
  let lastRow=lastData;
  if(matrix&&totals){
   const heading=lastData+2;rows.push(row(lastData+1,'',8));rows.push(row(heading,cell('Totais do filtro',0,heading,5),22));merge(heading,0,width-1);
   const capacity=Math.max(1,Math.floor(width/2));let next=heading+1;
   for(let start=0;start<kpis.length;start+=capacity){
    const group=kpis.slice(start,start+capacity),labels=[],values=[];let labelHeight=28;
    for(const [i,[label,value]] of group.entries()){
     const left=Math.floor(i*width/group.length),right=Math.floor((i+1)*width/group.length)-1;
     labels.push(cell(label,left,next,6));values.push(cell(value,left,next+1,7,true));merge(next,left,right);merge(next+1,left,right);
     labelHeight=Math.max(labelHeight,Math.ceil(String(label).length/((right-left+1)*12))*12+8);
    }
    rows.push(row(next,labels.join(''),Math.min(72,labelHeight)));rows.push(row(next+1,values.join(''),26));next+=2;
   }
   if(!kpis.length){rows.push(row(next,cell('Sem indicadores de totais disponíveis.',0,next,8),30));merge(next,0,width-1);next++;}
   const note='Os totais consideram todas as páginas do filtro, inclusive ao exportar apenas a página atual.';
   rows.push(row(next,cell(note,0,next,8),Math.max(30,Math.ceil(note.length/(width*10))*12)));merge(next,0,width-1);lastRow=next;
  }
  if(lastRow>1048576)throw new Error('A exportação é muito grande. Reduza os filtros e tente novamente.');
  const cols=sheet.headers.map((h,i)=>{const size=i===0?18:sheet.numeric?.[i]?18:/cliente|material|vendedor|televend|munic/i.test(h)&&!/c[oó]d/i.test(h)?34:22;return `<col min="${i+1}" max="${i+1}" width="${size}" customWidth="1"/>`;}).join('');
  files[`xl/worksheets/sheet${sheetIndex+1}.xml`]=utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:'+column(width-1)+Math.max(1,lastRow)+'"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="'+headerRow+'" topLeftCell="A'+(headerRow+1)+'" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>'+cols+'</cols><sheetData>'+rows.join('')+'</sheetData>'+(matrix?'<autoFilter ref="A'+headerRow+':'+column(width-1)+lastData+'"/>':'')+(merges.length?'<mergeCells count="'+merges.length+'">'+merges.map(ref=>`<mergeCell ref="${ref}"/>`).join('')+'</mergeCells>':'')+'</worksheet>');
 }
 files['xl/sharedStrings.xml']=utf8('<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'+strings+'" uniqueCount="'+shared.length+'">'+shared.map(s=>'<si><t xml:space="preserve">'+xml(s)+'</t></si>').join('')+'</sst>');
 const xf=(font,fill,num=0,align='left',wrap=false)=>`<xf numFmtId="${num}" fontId="${font}" fillId="${fill}" borderId="0" xfId="0" applyFill="1" applyFont="1" applyAlignment="1"${num?' applyNumberFormat="1"':''}><alignment horizontal="${align}" vertical="center"${wrap?' wrapText="1"':''}/></xf>`;
 files['xl/styles.xml']=utf8('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts><fonts count="5"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FF052B5F"/><sz val="10"/><name val="Calibri"/></font><font><b/><color rgb="FF052B5F"/><sz val="13"/><name val="Calibri"/></font><font><color rgb="FF526574"/><sz val="9"/><name val="Calibri"/></font></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF052B5F"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEAF2F5"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2C8580"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9">'+[xf(0,0),xf(1,2,0,'left',true),xf(0,0,164,'right'),xf(0,3),xf(0,3,164,'right'),xf(1,4,0,'left',true),xf(2,3,0,'left',true),xf(3,3,164,'right'),xf(4,0,0,'left',true)].join('')+'</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
 files['xl/workbook.xml']=utf8('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+sheets.map((s,i)=>`<sheet name="${xml(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')+'</sheets></workbook>');
 files['xl/_rels/workbook.xml.rels']=utf8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')+`<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId${sheets.length+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
 files['_rels/.rels']=utf8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
 files['[Content_Types].xml']=utf8('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'+sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')+'</Types>');
 return files;
}
