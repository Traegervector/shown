/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file / React',
  component: 'file',
  decorators: [withCodeEditor]
};

export let file = () => html`
  <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2"></mgt-file>
  <react>
    import { File } from '@microsoft/mgt-react';

    export default () => (
      <File fileQuery='/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2'></File>
    );
  </react>
`;

export let events = () => html`
  <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2"></mgt-file>
 
  <react>
    import { useCallback } from 'react';
    import { File } from '@microsoft/mgt-react';

    export default () => {
      let onUpdated = useCallback((e) => {
        console.log('updated', e);
      }, []);

      return (
        <File 
        fileQuery='/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2' 
        updated={onUpdated}>
    </File>
      );
    };
  </react>

  <script>
    let file = document.querySelector('mgt-file');
    file.addEventListener('updated', e => {
      console.log('updated', e);
    });
  </script>
`;
