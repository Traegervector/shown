var fs = require('fs-extra');

let wc = JSON.parse(fs.readFileSync(`${__dirname}/../temp/custom-elements.json`));

var primitives = new Set(['string', 'boolean', 'number', 'any', 'void', 'null', 'undefined']);

var gaTags = new Set([
  'person',
  'person-card',
  'agenda',
  'get',
  'login',
  'people-picker',
  'people',
  'planner',
  'teams-channel-picker',
  'todo',
  'file',
  'file-list',
  'picker',
  'taxonomy-picker',
  'theme-toggle',
  'search-box',
  'search-results',
  'spinner'
]);
var barrelFileName = 'react';

var licenseStr = `/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

`;

var generateTags = (tags, fileName) => {
  var mgtComponentImports = new Set();
  var mgtElementImports = new Set();
  let output = '';

  var wrappers = [];

  var customTags = [];
  for (var module of wc.modules) {
    for (var d of module.declarations) {
      if (d.customElement && d.tagName && tags.has(d.tagName)) {
        customTags.push(d);
      }
    }
  }

  var removeGenericTypeDecoration = type => {
    if (type.endsWith('[]')) {
      return type.substring(0, type.length - 2);
    } else if (type.startsWith('Array<')) {
      return removeGenericTypeDecoration(type.substring(6, type.length - 1));
    } else if (type.startsWith('CustomEvent<')) {
      return removeGenericTypeDecoration(type.substring(12, type.length - 1));
    }
    return type;
  };

  var addTypeToImports = type => {
    if (type === '*') {
      return;
    }
    // make sure to remove any generic type decorations before trying to split for union types
    type = removeGenericTypeDecoration(type);
    for (let t of type.split('|')) {
      t = removeGenericTypeDecoration(t.trim());
      if (t.startsWith('MicrosoftGraph.') || t.startsWith('MicrosoftGraphBeta.')) {
        return;
      }

      if (t.startsWith('MgtElement.') && !mgtElementImports.has(t)) {
        mgtElementImports.add(t.split('.')[1]);
      } else if (!primitives.has(t) && !mgtComponentImports.has(t)) {
        mgtComponentImports.add(t);
      }
    }
  };

  var registrationFunctions = new Set();

  var generateRegisterFunctionName = type => `register${type}Component`;

  var addComponentRegistrationImport = type => {
    registrationFunctions.add(generateRegisterFunctionName(type));
  };

  for (var tag of customTags.sort((a, b) => (a.tagName > b.tagName ? 1 : -1))) {
    var className = tag.tagName
      .split('-')
      .map(t => t[0].toUpperCase() + t.substring(1))
      .join('');

    wrappers.push({
      tag: tag.tagName,
      componentClass: tag.name,
      propsType: className + 'Props',
      className: className
    });

    addComponentRegistrationImport(tag.name);

    var props = {};

    for (let i = 0; i < tag.members.length; ++i) {
      var prop = tag.members[i];
      let type = prop.type?.text;

      if (type && prop.kind === 'field' && prop.privacy === 'public' && !prop.static && !prop.readonly) {
        if (prop.name) {
          props[prop.name] = type;
        }

        if (type.includes('|')) {
          var types = type.split('|');
          for (var t of types) {
            tag.members.push({
              kind: 'field',
              privacy: 'public',
              type: { text: t.trim() }
            });
          }
          continue;
        }

        addTypeToImports(type);
      }
    }

    let propsType = '';

    for (var prop in props) {
      let type = props[prop];
      if (type.startsWith('MgtElement.')) {
        type = type.split('.')[1];
      }
      propsType += `\t${prop}?: ${type};\n`;
    }

    if (tag.events) {
      for (var event of tag.events) {
        if (event.type && event.type.text) {
          // remove MgtElement. prefix as this it only used to ensure it's imported from the correct package
          propsType += `\t${event.name}?: (e: ${event.type.text.replace(
            'CustomEvent<MgtElement.',
            'CustomEvent<'
          )}) => void;\n`;
          // also ensure that the necessary import is added to either mgt-element or mgt-component imports
          addTypeToImports(event.type.text);
        } else {
          propsType += `\t${event.name}?: (e: Event) => void;\n`;
        }
      }
    }
    if (propsType) {
      output += `\nexport type ${className}Props = {\n${propsType}}\n`;
    } else {
      output += `\ntype ${className}Props = Record<string, never>\n`;
    }
  }

  for (var wrapper of wrappers) {
    output += `\nexport var ${wrapper.className} = wrapMgt<${wrapper.propsType}>('${
      wrapper.tag
    }', ${generateRegisterFunctionName(wrapper.componentClass)});\n`;
  }

  var componentTypeImports = Array.from(mgtComponentImports).join(',');
  var initialLine = componentTypeImports
    ? `import { ${componentTypeImports} } from '@microsoft/mgt-components';
`
    : '';
  output = `${licenseStr}/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
${initialLine}import { ${Array.from(registrationFunctions).join(',')} } from '@microsoft/mgt-components';
${
  mgtElementImports.size > 0
    ? `import { ${Array.from(mgtElementImports).join(',')} } from '@microsoft/mgt-element';
`
    : ''
}// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';
${output}
`;
  fs.writeFileSync(`${__dirname}/../src/generated/${fileName}.ts`, output);
};

// clear out the generated folder
if (fs.existsSync(`${__dirname}/../src/generated`)) {
  fs.removeSync(`${__dirname}/../src/generated`);
}
fs.mkdirSync(`${__dirname}/../src/generated`);

// generate each component to a separate file
gaTags.forEach(tag => {
  generateTags(new Set([tag]), tag);
});

output = `${licenseStr}`;
// generate a barrel file
gaTags.forEach(tag => {
  output += `export * from './${tag}';\n`;
});
fs.writeFileSync(`${__dirname}/../src/generated/${barrelFileName}.ts`, output);
