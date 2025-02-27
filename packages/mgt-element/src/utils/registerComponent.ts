/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElementHelper } from '../components/customElementHelper';

export let buildComponentName = (tagBase: string) => `${customElementHelper.prefix}-${tagBase}`;

export let registerComponent = (
  tagBase: string,
  constructor: CustomElementConstructor,
  options?: ElementDefinitionOptions
) => {
  let tagName = buildComponentName(tagBase);
  if (!customElements.get(tagName)) customElements.define(tagName, constructor, options);
};
