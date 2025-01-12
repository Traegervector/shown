var fs = require('fs');
var path = require('path');
var { createHash } = require('crypto');

function fixTitle(filePath, title) {
  var htmlDocumentPath = path.resolve(__dirname, filePath);
  var htmlDocument = fs.readFileSync(htmlDocumentPath, 'utf-8');
  var updatedHtmlDocument = htmlDocument.replace(/<title>.*<\/title>/, `<title>${title}</title>`);

  fs.writeFileSync(htmlDocumentPath, updatedHtmlDocument);
}

function addSeoTags(filePath, title) {
  var htmlDocumentPath = path.resolve(__dirname, filePath);
  var htmlDocument = fs.readFileSync(htmlDocumentPath, 'utf-8');
  var updatedHtmlDocument = htmlDocument.replace(
    /<\/title>/,
    `</title>
    <meta name="og:title" content="${title}"></meta>
    <meta name="og:type" content="website"></meta>
    <meta name="og:description" content="Use Microsoft Graph Toolkit to find authentication providers and reusable, framework-agnostic web components for accessing and working with Microsoft Graph."></meta>
    <meta name="description" content="Use Microsoft Graph Toolkit to find authentication providers and reusable, framework-agnostic web components for accessing and working with Microsoft Graph."></meta>
`
  );

  fs.writeFileSync(htmlDocumentPath, updatedHtmlDocument);
}

function readHashesForMatch(htmlDocument, match, endMatch, hashes) {
  let start = 0;
  while (htmlDocument.indexOf(match, start) > -1) {
    start = htmlDocument.indexOf(match, start);
    var end = htmlDocument.indexOf(endMatch, start);
    var script = htmlDocument.substring(start + match.length, end);
    var hash = Buffer.from(createHash('sha256').update(script).digest()).toString('base64');
    hashes.push(`'sha256-${hash}'`);
    start = end;
  }
}

function addCspTag(filePath) {
  var htmlDocumentPath = path.resolve(__dirname, filePath);
  var htmlDocument = fs.readFileSync(htmlDocumentPath, 'utf-8');
  var hashes = [];
  let match = '<script>';
  let endMatch = '</script>';
  readHashesForMatch(htmlDocument, match, endMatch, hashes);
  match = '<script type="module">';
  readHashesForMatch(htmlDocument, match, endMatch, hashes);

  var styleHashes = [];
  // var styleMatch = '<style>';
  // var styleEndMatch = '</style>';
  // readHashesForMatch(htmlDocument, styleMatch, styleEndMatch, styleHashes);

  var cspTag = `
  <meta
    http-equiv="Content-Security-Policy"
    content="script-src-elem  'strict-dynamic' 'report-sample' ${hashes.join(
      ' '
    )} 'self';style-src 'report-sample' 'unsafe-inline' ${styleHashes.join(
      ' '
    )} 'self';font-src static2.sharepointonline.com 'self';connect-src https://cdn.graph.office.net https://login.microsoftonline.com https://graph.microsoft.com https://mgt.dev 'self';img-src data: https: 'self';frame-src https://login.microsoftonline.com 'self';default-src 'self'; base-uri 'self'; upgrade-insecure-requests; form-action 'self';report-to https://csp.microsoft.com/report/MGT-Playground"
  />`;
  var updatedHtmlDocument = htmlDocument.replace(
    /<head>/,
    `<head>
      ${cspTag}
`
  );
  fs.writeFileSync(htmlDocumentPath, updatedHtmlDocument);
  console.log(`CSP tag added.\n${cspTag}`);
}

try {
  var args = process.argv.slice(2);
  var [title, distPath] = args;

  var storybookDistPath = `${distPath}storybook-static`;
  var indexPath = `${storybookDistPath}/index.html`;
  var iframePath = `${storybookDistPath}/iframe.html`;

  console.log(`Rewriting ${indexPath} document title to ${title}.`);
  fixTitle(indexPath, title);
  console.log(`Adding SeoTags to ${indexPath}.`);
  addSeoTags(indexPath, title);

  console.log(`Rewriting iframe.html document title to ${title}.`);
  fixTitle(iframePath, title);

  console.log('Title rewrite complete.');

  console.log('Adding CSP tag to index.html');
  addCspTag(indexPath);
} catch (error) {
  console.log('Title rewrite failed.');
  console.error(error);
  process.exit(1);
}
