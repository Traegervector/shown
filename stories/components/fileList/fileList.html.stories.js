/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file-list / HTML',
  component: 'file-list',
  decorators: [withCodeEditor]
};

export let fileList = () => html`
  <mgt-file-list></mgt-file-list>
`;

export let RTL = () => html`
  <body dir="rtl">
    <mgt-file-list></mgt-file-list>
  </body>
`;

export let localization = () => html`
  <mgt-file-list></mgt-file-list>
  <script>
  import { LocalizationHelper } from '@microsoft/mgt-element';
  LocalizationHelper.strings = {
    _components: {
      'file-list': {
        showMoreSubtitle: 'Show more 📂'
      },
      'file': {
        modifiedSubtitle: '⚡',
      }
    }
  }
  </script>
`;

export let events = () => html`
  <p>Clicked File:</p>
  <mgt-file></mgt-file>
  <mgt-file-list disable-open-on-click></mgt-file-list>
  <script>
    document.querySelector('mgt-file-list').addEventListener('itemClick', e => {
      let file = document.querySelector('mgt-file');
      file.fileDetails = e.detail;
    });

    document.querySelector('mgt-file-list').addEventListener('updated', e => {
      console.log('updated', e);
    });
  </script>
  <style>
    body {
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
        'Helvetica Neue', sans-serif;
    }

    p {
      margin: 0;
    }
  </style>
`;

export let openFolderBreadcrumbs = () => html`
    <style>
      body {
        font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
          'Helvetica Neue', sans-serif;
      }

      ul.breadcrumb {
        margin: 0;
        padding: 10px 16px;
        list-style: none;
        background-color: var(--neutral-layer-1);
        font-size: 12px;
      }

      ul.breadcrumb li {
        display: inline;
      }

      ul.breadcrumb li + li:before {
        padding: 8px;
        color: var(--foreground-on-accent-fill);
        content: '\/';
      }

      ul.breadcrumb li a {
        color: var(--accent-fill-rest);
        text-decoration: none;
      }

      ul.breadcrumb li a:hover {
        color: var(--accent-fill-hover);
        text-decoration: underline;
      }
    </style>

    <ul class="breadcrumb" id="nav">
      <li><a id="home">Files</a></li>
    </ul>
    <mgt-file-list id="parent-file-list"></mgt-file-list>

    <script type="module">
      let fileList = document.getElementById('parent-file-list');
      let nav = document.getElementById('nav');

      // handle create and remove menu items
      fileList.addEventListener('itemClick', e => {
        if (e.detail && e.detail.folder) {
          let id = e.detail.id;
          let name = e.detail.name;
          let breadcrumbId = "breadcrumb-"+ id;

          // check if it is set
          let breadcrumbSet = document.getElementById(breadcrumbId);
          if (!breadcrumbSet) {
            // create breadcrumb menu item
            let li = document.createElement('li');
            let a = document.createElement('a');
            li.setAttribute('id', breadcrumbId);
            a.appendChild(document.createTextNode(name));
            li.appendChild(a);
            if (nav.children.length > 1){
              let firstBreadcrumb = nav.children[0];
              nav.replaceChildren(firstBreadcrumb);
            }
            nav.appendChild(li);
          } else {
            breadcrumbSet.remove();
            
          }
        }
      });
    </script>
  `;
