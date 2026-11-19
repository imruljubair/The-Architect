import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import ts from 'typescript';
import {Workbook,ExcelScript,formulas} from './excel-mock.mjs';

function load(name='main') {
 const code=fs.readFileSync(new URL(`../src/${name}.ts`,import.meta.url),'utf8');
 const js=ts.transpileModule(code,{compilerOptions:{target:ts.ScriptTarget.ES2020,module:ts.ModuleKind.None}}).outputText;
 const context=vm.createContext({ExcelScript,console:{log(){}}});
 vm.runInContext(js,context);
 return context.main;
}
const main=load();
const cases={
 classification:[['Rand',5,6],['Linear',4,''],['ReLU','',''],['Linear',3,''],['Softmax','',''],['CELoss','','']],
 regression:[['Rand',3,2],['Linear',4,''],['Tanh','',''],['Linear',2,''],['MSELoss','','']],
 normalization:[['Rand',3,2],['Linear',4,''],['LayerNorm','',''],['GELU','',''],['Linear',2,''],['MSELoss','','']],
 tinyDigits:[['tinyDigits','',2],['Flatten','',''],['Linear',10,''],['Softmax','',''],['CELoss','','']],
 inference:[['Rand',3,2],['Linear',2,''],['ReLU','','']],
};
for(const [name,layers] of Object.entries(cases)) {
 test(`${name}: forward generation ignores the former checkbox`,()=>{
  const book=new Workbook(layers,8);
  let deleted=0;
  const shapes=Array.from({length:3},()=>({delete(){deleted++;}}));
  book.getWorksheet('Architect').shapes=shapes;
  book.getWorksheet('Architect').getRange('G5').setValue(true);
  for(let n=0;n<2;n++) {
   main(book);
   assert.deepEqual([...book.sheets.keys()].sort(),['Architect','Blueprint','Helper','tinyDigits'].sort());
   const text=[...book.getWorksheet('Blueprint').cells.values()].map(v=>v.value??v.formula??'').join('\n');
   assert.match(text,/Forward/);
   assert.doesNotMatch(text,/Backprop|Weight Update|\.backward\(|optimizer/);
   assert.ok(formulas(book).some(([,v])=>v.formula.includes('MMULT(')));
   assert.equal(deleted,0);
   assert.equal(book.getWorksheet('Architect').shapes.length,3);
  }
 });
}
test('original randomize and clear logic preserve all three buttons',()=>{
 const book=new Workbook(cases.classification);
 let deleted=0;
 book.getWorksheet('Architect').shapes=Array.from({length:3},()=>({delete(){deleted++;}}));
 load('RandomInputs')(book);
 assert.ok(book.getWorksheet('Architect').getRange('B5').getValue());
 main(book);
 load('ClearButton')(book);
 assert.equal(book.getWorksheet('Architect').getRange('B5').getValue(),'');
 assert.equal(book.getWorksheet('Blueprint'),undefined);
 assert.equal(deleted,0);
});
test('all distributed Office Scripts match their source',()=>{
 for(const [name,file] of [['main','architect_Mar24_2026_improve'],['RandomInputs','RandomInputs'],['ClearButton','ClearButton']]) {
  assert.equal(JSON.parse(fs.readFileSync(new URL(`../${file}.osts`,import.meta.url))).body,fs.readFileSync(new URL(`../src/${name}.ts`,import.meta.url),'utf8'));
 }
});
for(const action of ['RandomInputs','ClearButton']) {
 test(`${action}: disabled backprop cell stays blank after a reset`,()=>{
  const book=new Workbook(cases.classification);
  const cell=book.getWorksheet('Architect').getRange('G5');
  cell.setControl({type:ExcelScript.CellControlType.empty});
  cell.setValue(false);
  load(action)(book);
  assert.equal(cell.getValue(),'');
  assert.equal(cell.getControl().type,ExcelScript.CellControlType.empty);
 });
 test(`${action}: reset does not require checkbox APIs`,()=>{
  const book=new Workbook(cases.classification);
  const cell=book.getWorksheet('Architect').getRange('G5');
  cell.getControl=()=>{throw new Error('Checkbox inspection is unavailable');};
  cell.setControl=()=>{throw new Error('Checkbox mutation is unavailable');};
  load(action)(book);
  assert.equal(cell.getValue(),'');
  assert.doesNotMatch(fs.readFileSync(new URL(`../src/${action}.ts`,import.meta.url),'utf8'),/\.getControl\(|\.setControl\(/);
 });
}
