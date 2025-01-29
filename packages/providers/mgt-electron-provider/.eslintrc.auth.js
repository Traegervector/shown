let cfg = require('../../../.eslintrc.js');

let config = Object.assign(cfg, {
  parserOptions: {
    project: 'tsconfig.authenticator.json',
    sourceType: 'module'
  }
});

module.exports = config;
