var cp = require('child_process');
var path = require('path');

function usage() {
  console.log('Usage: node run_ai_node.js <cx> <aiFilePath> <jsxFilePath> [progId]');
  process.exit(1);
}

if (process.argv.length < 5) usage();

var cx = process.argv[2];
var aiPath = process.argv[3];
var jsxPath = process.argv[4];
var progId = process.argv[5] || 'Illustrator.Application';

var dataStr = cx + ';' + aiPath;

var vbsPath = path.join(__dirname, 'run_ai.vbs');
var args = ['//nologo', vbsPath, jsxPath, dataStr, progId];

var ret = cp.spawnSync('cscript', args, { stdio: 'inherit' });
process.exit(ret.status || 0);