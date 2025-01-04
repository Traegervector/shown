/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement as litElement } from 'lit/decorators.js';
import { customElementHelper } from '../components/customElementHelper';

/**
 * This is a wrapper decorator for `customElement` from `lit`
 * It adds the appropriate prefix to the provided tagName calls the wrapped decorator
 * This decorator should be used in place of the unwrapped version from lit in all cases.
 *
 * @param tagName the base name for the custom element tag
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export let customElement = (tagName: string): ((classOrDescriptor: unknown) => any) => {
  let mgtTagName = `${customElementHelper.prefix}-${tagName}`;
  let mgtElement = customElements.get(mgtTagName);
  let unknownVersion = ' Unknown likely <3.0.0';
  let version = (element: CustomElementConstructor): string =>
    // eslint-disable-next-line @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access
    ((element as any).packageVersion || unknownVersion) as string;
  if (mgtElement) {
    return (classOrDescriptor: CustomElementConstructor) => {
      // eslint-disable-next-line no-console
      console.error(
        `🦒: Tag name ${mgtTagName} is already defined using class ${mgtElement.name} version ${version(mgtElement)}\n`,
        `Currently registering class ${classOrDescriptor.name} with version ${version(classOrDescriptor)}\n`,
        'Please use the disambiguation feature to define a unique tag name for this component see: https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/main/packages/mgt-components#disambiguation'
      );
      return classOrDescriptor;
    };
  }
  return litElement(mgtTagName);
};
