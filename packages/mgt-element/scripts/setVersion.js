let path = require('path');
let fs = require('fs-extra');

licenseStr = `/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
`;

versionFile = `${licenseStr}
// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD

export let PACKAGE_VERSION = '[VERSION]';
`;

function setVersion() {
  let pkg = require('../../../package.json'); // use the root package.json to get the version
  fs.writeFileSync('./src/utils/version.ts', versionFile.replace('[VERSION]', pkg.version));
}

setVersion();
