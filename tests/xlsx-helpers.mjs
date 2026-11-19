import {unzipSync,strFromU8} from 'fflate';
import {createHash} from 'node:crypto';
export const decode=s=>s.replace(/&#x([a-f0-9]+);/gi,(_,n)=>String.fromCodePoint(parseInt(n,16))).replace(/&#(\d+);/g,(_,n)=>String.fromCodePoint(Number(n))).replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
export function attributes(s) {
  return Object.fromEntries([...s.matchAll(/([\w:]+)="([^"]*)"/g)].map(m=>[m[1],decode(m[2])]));
}
export function openXlsx(bytes) {
  const parts=unzipSync(bytes);
  const xml=p=>strFromU8(parts[p]);
  const rels=Object.fromEntries([...xml('xl/_rels/workbook.xml.rels').matchAll(/<Relationship\b([^>]*)\/?\s*>/g)].map(m=>{const a=attributes(m[1]);return [a.Id,a.Target]}));
  const sheetPaths=Object.fromEntries([...xml('xl/workbook.xml').matchAll(/<sheet\b([^>]*)\/?\s*>/g)].map(m=>{
    const a=attributes(m[1]),target=rels[a['r:id']];
    return [a.name,target.startsWith('/')?target.slice(1):`xl/${target}`];
  }));
  return {parts,xml,sheetPaths};
}
export function sheetFingerprint(book,name) {
  const strings=[...book.xml('xl/sharedStrings.xml').matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(m=>[...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(t=>decode(t[1])).join(''));
  const values=[];
  for(const m of book.xml(book.sheetPaths[name]).matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const a=attributes(m[1]),body=m[2]||'',formula=body.match(/<f\b[^>]*?(?:\/>|>([\s\S]*?)<\/f>)/);
    if(formula) { values.push([a.r,'formula',decode(formula[1]||'')]);continue; }
    const v=body.match(/<v\b[^>]*>([\s\S]*?)<\/v>/);
    if(!v)continue;
    let val=decode(v[1]),type='text';
    if(a.t==='s') val=strings[Number(val)];
    else if(a.t==='b') {val=val==='1';type='boolean';}
    else if(!a.t||a.t==='n') {val=Number(val);type='number';}
    if(val!=='')values.push([a.r,type,val]);
  }
  values.sort((a,b)=>a[0]<b[0]?-1:a[0]>b[0]?1:0);
  return createHash('sha256').update(JSON.stringify(values)).digest('hex');
}
