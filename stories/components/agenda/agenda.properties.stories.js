/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-agenda / Properties',
  component: 'agenda',
  decorators: [withCodeEditor]
};

export let getByEventQuery = () => html`
  <mgt-agenda event-query="/me/calendarview?$orderby=start/dateTime&startdatetime=2023-07-12T00:00:00.000Z&enddatetime=2023-07-18T00:00:00.000Z"></mgt-agenda>
`;

export let groupByDay = () => html`
  <mgt-agenda group-by-day></mgt-agenda>
`;

export let showMax = () => html`
  <mgt-agenda show-max="4"></mgt-agenda>
`;

export let getByDate = () => html`
  <mgt-agenda group-by-day date="May 7, 2019" days="3"></mgt-agenda>
`;

export let getEventsForNextWeek = () => html`
  <mgt-agenda group-by-day days="7"></mgt-agenda>
`;

export let preferredTimezone = () => html`
  <mgt-agenda preferred-timezone="Europe/Paris"></mgt-agenda>
`;
