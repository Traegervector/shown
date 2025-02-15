/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person / HTML',
  component: 'person',
  decorators: [withCodeEditor]
};

export let person = () => html`
  <mgt-person person-query="me"></mgt-person>
  <br>
  <mgt-person person-query="me" view="oneline"></mgt-person>
  <br>
  <mgt-person person-query="me" view="twolines"></mgt-person>
  <br>
  <mgt-person person-query="me" view="threelines"></mgt-person>
  <br>
  <mgt-person person-query="me" view="fourlines"></mgt-person>
`;

export let events = () => html`
  <!-- Check JS tab to see event listeners -->
  <div style="margin-bottom: 10px">Click on each line</div>
  <div class="example">
    <mgt-person person-query="me" view="fourlines">
      <template data-type="line1">
        <div tabindex="0">
          Hello, my name is: {{person.displayName}}
        </div>
      </template>
      <template data-type="line2">
        <div tabindex="0">
          ✨ {{person.jobTitle}}
        </div>
      </template>
      <template data-type="line3">
        <div tabindex="0">
          👀 {{person.department}}
        </div>
      </template>
      <template data-type="line4">
        <div tabindex="0">
          📍{{person.mail}}
        </div>
      </template>
    </mgt-person>
  </div>
  <div class="output">no line clicked</div>

  <script>
    let person = document.querySelector('mgt-person');
    person.addEventListener('updated', e => {
        console.log('updated', e);
    });

    person.addEventListener('line1clicked', e => {
        let output = document.querySelector('.output');

        if (e && e.detail && e.detail.displayName) {
            output.innerHTML = '<b>line1clicked:</b> ' + e.detail.displayName;
        }
    });
    person.addEventListener('line2clicked', e => {
        let output = document.querySelector('.output');

        if (e && e.detail && e.detail.jobTitle) {
            output.innerHTML = '<b>line2clicked:</b> ' + e.detail.jobTitle;
        }
    });
    person.addEventListener('line3clicked', e => {
        let output = document.querySelector('.output');

        if (e && e.detail && e.detail.department) {
            output.innerHTML = '<b>line3clicked:</b> ' + e.detail.department;
        }
    });
    person.addEventListener('line4clicked', e => {
        let output = document.querySelector('.output');

        if (e && e.detail && e.detail.mail) {
            output.innerHTML = '<b>line4clicked:</b> ' + e.detail.mail;
        }
    });
  </script>

  <style>
    .example {
      margin-bottom: 20px;
    }
  </style>
`;

export let RTL = () => html`
  <body dir="rtl">
    <div class="example">
      <mgt-person person-query="me" view="oneline"></mgt-person>
    </div>
    <div class="example">
      <mgt-person person-query="me" view="twolines"></mgt-person>
    </div>
    <div class="example">
      <mgt-person person-query="me" view="threelines"></mgt-person>
    </div>
    <div class="example">
      <mgt-person person-query="me" view="fourlines"></mgt-person>
    </div>

    <!-- RTL with vertical layout -->
    <div class="row">
      <mgt-person person-query="me" class="example" vertical-layout id="online" view="oneline"></mgt-person>
    </div>
    <div class="row">
      <mgt-person person-query="me" class="example" vertical-layout id="online2" view="twolines"></mgt-person>
    </div>
    <div class="row">
      <mgt-person person-query="me" class="example" vertical-layout id="online3" view="threelines"></mgt-person>
    </div>
    <div class="row">
      <mgt-person person-query="me" class="example" vertical-layout id="online4" view="fourlines"></mgt-person>
    </div>
  </body>
  <style>
  .example {
      margin-bottom: 20px;
    }
    </style>
`;

export let personVertical = () => html`


<mgt-person person-query="me" class="example" vertical-layout view="oneline" person-card="hover"></mgt-person>
<mgt-person person-query="me" class="example" vertical-layout view="twolines" person-card="hover"></mgt-person>
<mgt-person person-query="me" class="example" vertical-layout view="threelines"></mgt-person>
<mgt-person person-query="me" class="example" vertical-layout view="fourlines"></mgt-person>

<!-- With Presence; Check JS tab -->

<mgt-person person-query="me" class="example" vertical-layout id="online" show-presence view="oneline" person-card="hover"></mgt-person>
<mgt-person person-query="me" class="example" vertical-layout id="online2" show-presence view="twolines" person-card="hover"></mgt-person>
<mgt-person person-query="me" class="example" vertical-layout id="online3" show-presence view="threelines"></mgt-person>
<mgt-person person-query="me" class="example" vertical-layout id="online4" show-presence view="fourlines"></mgt-person>

<!-- Person unauthenticated vertical layout-->

<mgt-person person-query="mbowen" vertical-layout view="twolines"
fallback-details='{"mail":"MeganB@M365x214355.onmicrosoft.com"}'>
</mgt-person>

<script>
            let online = {
          activity: 'Available',
          availability: 'Available',
          id: null
      };
      let onlinePerson = document.getElementById('online');
      let onlinePerson2 = document.getElementById('online2');
      let onlinePerson3 = document.getElementById('online3');
      let onlinePerson4 = document.getElementById('online4');

      onlinePerson.personPresence = online;
      onlinePerson2.personPresence = online;
      onlinePerson3.personPresence = online;
      onlinePerson4.personPresence = online;
    </script>
<style>
  .example {
      margin-bottom: 20px;
    }
    </style>
`;
