import fs from 'node:fs';
const names={main:'architect_Mar24_2026_improve',RandomInputs:'RandomInputs',ClearButton:'ClearButton'};
for(const [source,target] of Object.entries(names)) {
  const file=new URL(`../${target}.osts`,import.meta.url);
  const script=JSON.parse(fs.readFileSync(file,'utf8'));
  script.body=fs.readFileSync(new URL(`../src/${source}.ts`,import.meta.url),'utf8');
  fs.writeFileSync(file,JSON.stringify(script,null,2)+'\n');
}
