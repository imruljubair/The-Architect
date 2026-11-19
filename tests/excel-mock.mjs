// In-memory Office Scripts host for generator integration tests.
// Stores formula text; Excel itself remains responsible for formula calculation.
import assert from 'node:assert/strict';

function style() {
  const value = new Proxy({}, { get: (_, key) => {
    if (key === 'then') return undefined;
    if (key === 'getSeries') return () => [style()];
    if (key === 'getPoints') return () => [style()];
    if (/^(get|add)/.test(key)) return () => style();
    if (/^(set|clear|freeze|unfreeze)/.test(key)) return () => {};
    throw new Error(`Unsupported formatting method: ${String(key)}`);
  }});
  return value;
}
export function colName(c) {
  let s = '';
  for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s;
  return s;
}
function coordinate(s) {
  const m = s.replaceAll('$','').match(/^([A-Z]*)(\d*)$/);
  assert.ok(m, `Invalid address: ${s}`);
  return [m[2] ? Number(m[2]) - 1 : 0, [...m[1]].reduce((c,l) => c * 26 + l.charCodeAt(0)-64,0)-1];
}
class Range {
  constructor(sheet,r,c,nr=1,nc=1) {
    assert.ok([r,c,nr,nc].every(Number.isInteger) && r>=0 && c>=0 && nr>0 && nc>0, `Invalid range ${sheet.name}: ${[r,c,nr,nc]}`);
    Object.assign(this,{sheet,r,c,nr,nc});
  }
  each(fn) {
    for(let r=0;r<this.nr;r++) for(let c=0;c<this.nc;c++) fn(this.r+r,this.c+c,r,c);
  }
  setValue(value) { this.each((r,c)=>this.sheet.cells.set(`${r},${c}`,{value})); }
  setValues(values) { assert.equal(values.length,this.nr); this.each((r,c,i,j)=>this.sheet.cells.set(`${r},${c}`,{value:values[i][j]})); }
  setFormula(formula) { assert.equal(typeof formula,'string'); this.each((r,c)=>this.sheet.cells.set(`${r},${c}`,{formula})); }
  setFormulas(values) { assert.equal(values.length,this.nr); this.each((r,c,i,j)=>this.sheet.cells.set(`${r},${c}`,{formula:values[i][j]})); }
  getValues() { return Array.from({length:this.nr},(_,r)=>Array.from({length:this.nc},(_,c)=>this.sheet.cells.get(`${this.r+r},${this.c+c}`)?.value ?? '')); }
  getValue() { return this.getValues()[0][0]; }
  getFormula() { return this.sheet.cells.get(`${this.r},${this.c}`)?.formula || ''; }
  clear(apply) { if(apply !== 'formats') this.each((r,c)=>this.sheet.cells.delete(`${r},${c}`)); }
  getRowCount() { return this.nr; }
  getColumnCount() { return this.nc; }
  getRowIndex() { return this.r; }
  getColumnIndex() { return this.c; }
  getWorksheet() { return this.sheet; }
  getAddress() { return `${this.sheet.name}!${colName(this.c)}${this.r+1}:${colName(this.c+this.nc-1)}${this.r+this.nr}`; }
  getResizedRange(dr,dc) { return new Range(this.sheet,this.r,this.c,this.nr+dr,this.nc+dc); }
  getFormat() { return style(); }
  getDataValidation() { return style(); }
  addConditionalFormat() { return style(); }
  clearAllConditionalFormats() {}
  merge() {}
  unmerge() {}
  select() {}
  setControl(control) { this.sheet.controls.set(`${this.r},${this.c}`,control); }
  getControl() { return this.sheet.controls.get(`${this.r},${this.c}`) || {type:'empty'}; }
  getEntireColumn() { return this; }
  delete() { this.sheet.events.push('deleteColumns'); }
}
class NamedItem {
  constructor(owner,name,range) { Object.assign(this,{owner,name,range}); }
  getRange() { return this.range; }
  getName() { return this.name; }
  getFormula() { return `=${this.range.getAddress()}`; }
  delete() { this.owner.names.delete(this.name); }
}
class Sheet {
  constructor(book,name) { Object.assign(this,{book,name,cells:new Map(),names:new Map(),controls:new Map(),events:[],shapes:[]}); }
  getName() { return this.name; }
  setName(name) { this.book.sheets.delete(this.name); this.name=name; this.book.sheets.set(name,this); }
  getCell(r,c) { return new Range(this,r,c); }
  getRangeByIndexes(r,c,nr,nc) { return new Range(this,r,c,nr,nc); }
  getRange(address) {
    if (!address) return new Range(this,0,0,1000,100);
    if(this.names.has(address)) return this.names.get(address).getRange();
    const [a,b=a] = address.split(':');
    const [r,c] = coordinate(a), [er,ec] = coordinate(b);
    return new Range(this,r,Math.max(0,c),er-r+1,Math.max(0,ec)-Math.max(0,c)+1);
  }
  getUsedRange() {
    const keys=[...this.cells.keys()].map(k=>k.split(',').map(Number));
    return new Range(this,0,0,Math.max(1,...keys.map(k=>k[0]+1)),Math.max(1,...keys.map(k=>k[1]+1)));
  }
  getNames() { return [...this.names.values()]; }
  getNamedItem(name) { const n=this.names.get(name)||this.book.names.get(name); assert.ok(n, `Missing name ${this.name}!${name}`); return n; }
  addNamedItem(name,range) { const n=new NamedItem(this,name,range); this.names.set(name,n); return n; }
  delete() { this.book.sheets.delete(this.name); }
  copy() {
    const sheet=this.book.addWorksheet(`${this.name}_copy${this.book.sheets.size}`);
    sheet.cells=structuredClone(this.cells);
    for(const n of this.names.values()) sheet.addNamedItem(n.name,new Range(sheet,n.range.r,n.range.c,n.range.nr,n.range.nc));
    return sheet;
  }
  getShapes() { return this.shapes; }
  getFreezePanes() { return style(); }
  getCharts() { return []; }
  addChart() { this.events.push('chart'); return style(); }
  activate() { this.book.active=this.name; }
  setTabColor() {}
  setPosition() {}
  setShowGridlines() {}
  setShowHeadings() {}
}
export class Workbook {
  constructor(layers,epochs=1) {
    this.sheets=new Map(); this.names=new Map();
    const sheet=this.addWorksheet('Architect');
    this.addWorksheet('Helper'); this.addWorksheet('tinyDigits');
    sheet.getRangeByIndexes(4,1,layers.length,3).setValues(layers);
    sheet.getRange('G5').setValue(false); sheet.getRange('H5').setValue(epochs);
    // Pixel grid needed by Image input cases.
    sheet.getRange('P5:BK23').setValue(255);
  }
  getWorksheet(name) { return this.sheets.get(name); }
  getWorksheets() { return [...this.sheets.values()]; }
  addWorksheet(name) { assert.ok(!this.sheets.has(name),`Duplicate sheet ${name}`); const s=new Sheet(this,name);this.sheets.set(name,s);return s; }
  getNames() { return [...this.names.values()]; }
  getNamedItem(name) { const n=this.names.get(name); assert.ok(n,`Missing workbook name ${name}`); return n; }
  addNamedItem(name,range) { const n=new NamedItem(this,name,range);this.names.set(name,n);return n; }
  getApplication() { return style(); }
}
export const ExcelScript = new Proxy({}, {get:(_,type)=>new Proxy({}, {get:(_,name)=>name})});
export function formulas(book,sheet='Blueprint') {
  return [...book.getWorksheet(sheet).cells].filter(([,v])=>v.formula).sort(([a],[b])=>a.localeCompare(b));
}
