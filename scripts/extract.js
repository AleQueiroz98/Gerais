const fs = require('fs');
const src = fs.readFileSync('/root/.claude/uploads/840b6d74-7f0f-5bc9-b4b2-c44bce458b7b/f7e4af3c-dashboard_aceleracao_seminovos_11_1.html','utf8');
const start = src.indexOf('const FRENTES = [');
const end = src.indexOf('\n];', start) + 3;
const code = src.slice(start, end) + '\nmodule.exports = FRENTES;';
fs.writeFileSync('/tmp/_f.js', code);
const F = require('/tmp/_f.js');
fs.writeFileSync('frentes.json', JSON.stringify(F, null, 2));
console.log('frentes:', F.length, F.map(f=>f.id+':'+f.entregaveis.length).join(' '));
