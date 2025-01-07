/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-search-results / Properties',
  component: 'search-results',
  decorators: [withCodeEditor]
};

export let setSearchResultsQueryString = () => html`
  <mgt-search-results query-string="contoso">
  </mgt-search-results>
`;

export let setSearchResultsQueryTemplate = () => html`
  <mgt-search-results version="beta" query-string="contoso" query-template="({searchTerms}) Title:Northwind">
  </mgt-search-results>
`;

export let setSearchResultsEntityTypes = () => html`
  <style>
    .example {
       margin-bottom: 20px;
     }
  </style>

  <div class="example">
	  <h3>Sites</h3>
    <mgt-search-results query-string="contoso" entity-types="site">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Drives</h3>
    <mgt-search-results query-string="contoso" entity-types="drive">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Drive Items</h3>
    <mgt-search-results query-string="contoso" entity-types="driveItem">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Lists</h3>
    <mgt-search-results query-string="contoso" entity-types="list">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>List Items</h3>
    <mgt-search-results query-string="contoso" entity-types="listItem">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Messages</h3>
    <mgt-search-results query-string="marketing" entity-types="message">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Events</h3>
    <mgt-search-results query-string="marketing" entity-types="event">
    </mgt-search-results>
  </div>


  <div class="example">
	  <h3>Chat Messages</h3>
    <mgt-search-results query-string="marketing" entity-types="chatMessage">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Persons</h3>
    <mgt-search-results query-string="bowen" version="beta" entity-types="person">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>External Items</h3>
    <mgt-search-results query-string="contoso" content-sources="contosoproducts" version="beta" entity-types="externalItem">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Q&A</h3>
    <mgt-search-results query-string="contoso" version="beta" entity-types="qna">
    </mgt-search-results>
  </div>


  <div class="example">
	  <h3>Bookmarks</h3>
    <mgt-search-results query-string="contoso" version="beta" entity-types="bookmark">
    </mgt-search-results>
  </div>


  <div class="example">
	  <h3>Acronyms</h3>
    <mgt-search-results query-string="contoso" version="beta" entity-types="acronym">
    </mgt-search-results>
  </div>
`;

export let setSearchResultsEntityTypesCombined = () => html`
  <mgt-search-results query-string="contoso" entity-types="driveItem,listItem">
  </mgt-search-results>
`;

export let setSearchResultsScopes = () => html`
  <mgt-search-results query-string="contoso" scopes="User.Read.All,Files.Read.All">
  </mgt-search-results>
`;

export let setSearchResultsContentSources = () => html`
  <mgt-search-results query-string="contoso" entity-types="externalItem" content-sources="/external/connections/contosoProducts" scopes="ExternalItem.Read.All">
  </mgt-search-results>
`;

export let setSearchResultsVersion = () => html`
  <mgt-search-results query-string="contoso" entity-types="bookmark" version="beta" scopes="Bookmark.Read.All">
  </mgt-search-results>
`;

export let setSearchResultsSize = () => html`
  <mgt-search-results query-string="contoso" size="20">
  </mgt-search-results>
`;

export let setSearchResultsPagingMax = () => html`
  <mgt-search-results query-string="contoso" paging-max="10">
  </mgt-search-results>
`;

export let setSearchResultsFetchThumbnail = () => html`
  <mgt-search-results query-string="contoso" fetch-thumbnail>
  </mgt-search-results>
`;

export let setSearchResultsFields = () => html`
  <mgt-search-results query-string="contoso" version="beta" entity-types="driveItem" fields="Title,ID,ContentTypeId">
  </mgt-search-results>
`;

export let setSearchResultsEnableTopResults = () => html`
  <mgt-search-results query-string="marketing" entity-types="message" enable-top-results scopes="Mail.Read">
  </mgt-search-results>
`;

export let setSearchResultsCacheEnabled = () => html`
  <mgt-search-results query-string="contoso" cache-enabled>
  </mgt-search-results>
`;

export let setSearchResultsCacheInvalidationPeriod = () => html`
  <mgt-search-results query-string="contoso" cache-enabled cache-invalidation-period="30000">
  </mgt-search-results>
`;
