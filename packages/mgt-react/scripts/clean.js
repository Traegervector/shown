let fs = require('fs-extra');

let temp = `${__dirname}/../temp`;
let dist = `${__dirname}/../dist`;

fs.removeSync(temp);
fs.removeSync(dist);
