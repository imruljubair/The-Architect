import fs from 'node:fs';
import test from 'node:test';
import assert from 'node:assert/strict';
import {openXlsx,sheetFingerprint,attributes,decode} from './xlsx-helpers.mjs';

const book=openXlsx(fs.readFileSync(new URL('../theArchitect_finalV2.xlsx',import.meta.url)));
test('all three original buttons retain script links',()=>{
 const drawing=book.xml('xl/drawings/drawing1.xml');
 const shapes=[...drawing.matchAll(/<xdr:twoCellAnchor>([\s\S]*?)<\/xdr:twoCellAnchor>/g)].map(m=>m[1]);
 assert.equal(shapes.length,3);
 assert.deepEqual(shapes.map(s=>decode(s.match(/<a:t>([\s\S]*?)<\/a:t>/)[1]).trim()),['🎲','↺','▶︎']);
 const ids=book.xml('customXml/item1.xml');
 for(const shape of shapes) {
  const link=shape.match(/<asl:scriptLink\b([^>]*)/);
  assert.ok(link,'Every button requires a script association');
  assert.ok(ids.includes(JSON.parse(attributes(link[1]).val).shareId));
 }
 // Randomize/Clear keep their original links; only Run is redirected.
 assert.ok(ids.includes('012622D2XT7CIHOUZNR5DISLEO6WQNEKXQ'));
 assert.ok(ids.includes('012622D2SMHIAULIEIP5D2O3PLF3HD4IG6'));
 assert.ok(!ids.includes('012622D2T36UWMTBOYABCLA6M3YOY2T66H'));
});
test('backpropagation checkbox is disabled in the saved workbook',()=>{
 const sheet=book.xml(book.sheetPaths.Architect);
 const cell=sheet.match(/<c\b([^>]*\br="G5"[^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/);
 assert.ok(cell);
 assert.ok(!cell[2],'G5 has no training value');
 const style=Number(attributes(cell[1]).s||0);
 const xfs=book.xml('xl/styles.xml').match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1];
 const xf=[...xfs.matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)][style][0];
 assert.ok(!xf.includes('xfComplement'),'G5 must not have a native checkbox control');
 assert.ok(!Object.keys(book.sheetPaths).some(name=>/Dashboard|_ep|inference/.test(name)));
 const source=fs.readFileSync(new URL('../src/main.ts',import.meta.url),'utf8');
 assert.doesNotMatch(source,/backpropFlag|numOfEpoch|\.backward\(|torch\.optim|Dashboard/);
});
test('original dataset and helper values/formulas are preserved',()=>{
 assert.equal(sheetFingerprint(book,'tinyDigits'),'ff5e63f841b9faf7bf948560a1f4c086d17de188344ec6bf15c2be39e4036240');
 assert.equal(sheetFingerprint(book,'Helper'),'5e62ab10d1a20492750a3d3c56a482e2cde318172fe32edfacfdb61618d489a1');
});
