/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file-list / React',
  component: 'file-list',
  decorators: [withCodeEditor]
};

export let fileList = () => html`
  <mgt-file-list></mgt-file-list>
  <react>
    import { FileList } from '@microsoft/mgt-react';

    export default () => (
      <FileList></FileList>
    );
  </react>
`;

export let events = () => html`
  <mgt-file-list></mgt-file-list>
  <react>
    import { useCallback } from 'react';
    import { FileList } from '@microsoft/mgt-react';

    export default () => {
      let onUpdated = useCallback((e: CustomEvent<undefined>) => {
        console.log('updated', e);
      }, []);

      return (
        <FileList 
        updated={onUpdated}>
    </FileList>
      );
    };
  </react>
  <script>
    document.querySelector('mgt-file-list').addEventListener('updated', e => {
      console.log('updated', e);
    });
  </script>
`;
