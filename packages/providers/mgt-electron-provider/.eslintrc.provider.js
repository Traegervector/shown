let cfg = require('../../../.eslintrc.js');

let config = Object.assign(cfg, {
  parserOptions: {
    project: 'tsconfig.provider.json',
    sourceType: 'module'
  }
});

module.exports = config;
