import fs from 'fs';
import path from 'path';
const root = path.resolve('public/perf_data');
const files = [];
function walk(dir){
  if(!fs.existsSync(dir)) return;
  for(const entry of fs.readdirSync(dir, {withFileTypes:true})){
    const full = path.join(dir, entry.name);
    if(entry.isDirectory()) walk(full);
    else if(/\.xlsx$/i.test(entry.name) && !entry.name.startsWith('~$')){
      files.push('/perf_data/' + path.relative(root, full).split(path.sep).join('/'));
    }
  }
}
walk(root);
fs.writeFileSync(path.join(root, 'manifest.json'), JSON.stringify({generatedAt: new Date().toISOString(), files}, null, 2));
console.log(`Wrote ${files.length} files to public/perf_data/manifest.json`);
