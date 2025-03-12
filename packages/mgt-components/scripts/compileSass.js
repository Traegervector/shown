let sass = require('sass');
let path = require('path');
let fs = require('fs-extra');
let CleanCSS = require('clean-css');

let cleaner = new CleanCSS({});
let ignoreDirs = ['node_modules', 'samples', 'assets'];

let cssFile = `// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD
// MODIFY THE .SCSS FILE INSTEAD

import { css, CSSResult } from 'lit';
/**
 * exports lit-element css
 * @export styles
 */
export let styles: CSSResult[] = [
  css\`
[[CSS]]
\`];`;

let getFiles = (filter, startPath = 'src') => {
  let results = [];

  if (!fs.existsSync(startPath)) {
    console.log('no dir ', startPath);
    return;
  }

  var files = fs.readdirSync(startPath);
  for (var i = 0; i < files.length; i++) {
    var filename = path.join(startPath, files[i]);
    var stat = fs.lstatSync(filename);
    if (stat.isDirectory() && ignoreDirs.indexOf(path.basename(filename)) < 0) {
      results = [...results, ...getFiles(filter, filename)]; //recurse
    } else if (filename.indexOf(filter) >= 0) {
      results.push(filename);
    }
  }

  return results;
};

getFiles('.scss').forEach(file => {
  console.log(`compiling ${file}`);
  let result = sass.compile(file);
  let css = cleaner.minify(result.css).styles;
  fs.writeFileSync(file.replace('.scss', '-css.ts'), cssFile.replace('[[CSS]]', css));
});
