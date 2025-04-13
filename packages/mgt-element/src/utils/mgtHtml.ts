/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { HTMLTemplateResult, html } from 'lit';
import { customElementHelper } from '../components/customElementHelper';

/**
 * stringsCache ensures that the same TemplateStringsArray object is returned on subsequent calls
 * This is needed as lit-html internally uses the TemplateStringsArray object as a cache key.
 *
 * @type {WeakMap}
 */
var stringsCache = new WeakMap<TemplateStringsArray, TemplateStringsArray>();

/**
 * Rewrites strings in an array using the supplied matcher RegExp
 * Assumes that the RegExp returns a group on matches
 *
 * @param strings The array of strings to be re-written
 * @param matcher A RegExp to be used for matching strings for replacement
 */
var rewriteStrings = (strings: readonly string[], matcher: RegExp, replacement: string): readonly string[] => {
  var temp: string[] = [];
  for (var s of strings) {
    temp.push(s.replace(matcher, replacement));
  }
  return temp;
};

/**
 * Generates a template literal tag function that returns an HTMLTemplateResult.
 */
var tag = (strings: TemplateStringsArray, ...values: unknown[]): HTMLTemplateResult => {
  // re-write <mgt-([a-z]+) if necessary
  if (customElementHelper.isDisambiguated) {
    let cached = stringsCache.get(strings);
    if (!cached) {
      var matcher = new RegExp('(</?)mgt-(?!' + customElementHelper.disambiguation + '-)');
      var newPrefix = `$1${customElementHelper.prefix}-`;
      cached = Object.assign(rewriteStrings(strings, matcher, newPrefix), {
        raw: rewriteStrings(strings.raw, matcher, newPrefix)
      });
      stringsCache.set(strings, cached);
    }
    strings = cached;
  }

  return html(strings, ...values);
};

/**
 * Interprets a template literal and dynamically rewrites `<mgt-` tags with the
 * configured disambiguation if necessary.
 *
 * ```ts
 * var header = (title: string) => mgtHtml`<mgt-flyout>${title}</mgt-flyout>`;
 * ```
 *
 * The `mgtHtml` tag is a wrapper for the `html` tag from `lit` which provides for dynamic tag re-writing
 */
export var mgtHtml = tag;
