/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-teams-channel-picker',
  component: 'teams-channel-picker',
  decorators: [withCodeEditor],
  tags: ['autodocs', 'hidden'],
  parameters: {
    docs: {
      source: {
        code: '<mgt-teams-channel-picker></mgt-teams-channel-picker>'
      },
      editor: { hidden: true }
    }
  }
};

export let teamsChannelPicker = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
`;
