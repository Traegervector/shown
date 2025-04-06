/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-teams-channel-picker / HTML',
  component: 'teams-channel-picker',
  decorators: [withCodeEditor]
};

export let teamsChannelPicker = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
`;

export let getSelectedChannel = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>

  <button class="button">Get SelectedChannel</button>

  <div class="output"></div>

  <script>
    document.querySelector('.button').addEventListener('click', _ => {
      let picker = document.querySelector('mgt-teams-channel-picker');
      let output = document.querySelector('.output');

      if (picker.selectedItem) {
        output.innerHTML = '<b>channel:</b> ' + picker.selectedItem.channel.displayName;
        output.innerHTML += '<br/><b>team:</b> ' + picker.selectedItem.team.displayName;
      } else {
        output.innerText = 'no channel selected';
      }
    });
  </script>
`;

export let events = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>

  <div class="output">no channel selected</div>

  <script>
    let picker = document.querySelector('mgt-teams-channel-picker');
    picker.addEventListener('updated', e => {
      console.log('updated', e);
    });
    
    picker.addEventListener('selectionChanged', e => {
      let output = document.querySelector('.output');

      if (e.detail) {
        output.innerHTML = '<b>channel:</b> ' + e.detail.channel.displayName;
        output.innerHTML += '<br/><b>team:</b> ' + e.detail.team.displayName;
      } else {
        output.innerText = 'no channel selected';
      }
    });
  </script>
`;

export let selectChannel = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>

  <button class="button">Select Channel By Id</button>

  <script>
    let button = document.querySelector('.button');
    button.addEventListener('click', async _ => {
      let picker = document.querySelector('mgt-teams-channel-picker');
      button.disabled = true;
      await picker.selectChannelById('19:d0bba23c2fc8413991125a43a54cc30e@thread.skype');
      button.disabled = false;
    });
  </script>
`;

export let clearSelectedItem = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
  <button class="clear">Clear SelectedChannel</button>

  <div class="output"></div>

  <script>
    let picker = document.querySelector('mgt-teams-channel-picker');
    let output = document.querySelector('.output');
    let clear = document.querySelector('.clear');

    clear.addEventListener('click', _ => {
      picker.clearSelectedItem();
    });

    picker.addEventListener('selectionChanged', e => {
      if (e.detail) {
        output.innerHTML = '<b>channel:</b> ' + e.detail.channel.displayName;
        output.innerHTML += '<br/><b>team:</b> ' + e.detail.team.displayName;
      } else {
        output.innerText = 'no channel selected';
      }
    })
  </script>
`;

export let RTL = () => html`
  <body dir="rtl">
    <mgt-teams-channel-picker></mgt-teams-channel-picker>
  </body>
`;

export let Localization = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
  <script>
  import { LocalizationHelper } from '@microsoft/mgt-element';
    LocalizationHelper.strings = {
        _components: {
            "teams-channel-picker": {
                inputPlaceholderText: "حدد قناة",
                noResultsFound: "لم يتم العثور على نتائج",
                loadingMessage: "جار التحميل..."
            }
        }
    };
  </script>
`;
