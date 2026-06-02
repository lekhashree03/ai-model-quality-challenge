import fs from 'node:fs';
import path from 'node:path';

const perfDir = path.resolve('public', 'perf_data');
const manifestPath = path.join(perfDir, 'manifest.json');

function walk(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) return walk(full);
    if (entry.isFile() && entry.name.toLowerCase().endsWith('.xlsx')) {
      const publicPath = '/' + path.relative('public', full).replaceAll(path.sep, '/');
      return [{ path: publicPath }];
    }
    return [];
  });
}

fs.mkdirSync(perfDir, { recursive: true });
const files = walk(perfDir).filter((file) => !file.path.endsWith('/manifest.json'));
fs.writeFileSync(manifestPath, JSON.stringify({ generatedAt: new Date().toISOString(), files }, null, 2));
console.log(`Wrote ${manifestPath} with ${files.length} .xlsx file(s).`);
