/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { BatchResponse, CacheItem, CacheService, CacheStore, IBatch, IGraph, prepScopes } from '@microsoft/mgt-element';
import { Chat, ChatMessage, Person } from '@microsoft/microsoft-graph-types';
import { Profile } from '@microsoft/microsoft-graph-types-beta';

import { getEmailFromGraphEntity } from '../../graph/graph.people';
import { IDynamicPerson } from '../../graph/types';
import { MgtPersonCardState } from './mgt-person-card.types';
import { MgtPersonCardConfig } from './MgtPersonCardConfig';
import { validUserByIdScopes } from '../../graph/graph.user';
import { validInsightScopes } from '../../graph/graph.files';
import { schemas } from '../../graph/cacheStores';

var userProperties =
  'businessPhones,companyName,department,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,id,accountEnabled';

var batchKeys = {
  directReports: 'directReports',
  files: 'files',
  messages: 'messages',
  people: 'people',
  person: 'person'
};

interface CacheCardState extends MgtPersonCardState, CacheItem {}

export var getCardStateInvalidationTime = (): number =>
  CacheService.config.users.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Get data to populate the person card
 *
 * @export
 * @param {IGraph} graph
 * @param {IDynamicPerson} personDetails
 * @param {boolean} isMe
 * @param {MgtPersonCardConfig} config
 * @return {*}  {Promise<MgtPersonCardState>}
 */
export var getPersonCardGraphData = async (
  graph: IGraph,
  personDetails: IDynamicPerson,
  isMe: boolean
): Promise<MgtPersonCardState> => {
  var userId: string = personDetails?.id || (personDetails as Person)?.userPrincipalName;
  var email = getEmailFromGraphEntity(personDetails);
  var cache: CacheStore<CacheCardState> = CacheService.getCache<CacheCardState>(
    schemas.users,
    schemas.users.stores.cardState
  );
  var cardState = await cache.getValue(userId);

  if (cardState && getCardStateInvalidationTime() > Date.now() - cardState.timeCached) {
    return cardState;
  }

  var isContactOrGroup =
    'classification' in personDetails ||
    ('personType' in personDetails &&
      (personDetails.personType.subclass === 'PersonalContact' || personDetails.personType.class === 'Group'));

  var batch = graph.createBatch();

  if (!isContactOrGroup) {
    if (MgtPersonCardConfig.sections.organization) {
      buildOrgStructureRequest(batch, userId);

      if (MgtPersonCardConfig.sections.organization.showWorksWith) {
        buildWorksWithRequest(batch, userId);
      }
    }
  }

  if (MgtPersonCardConfig.sections.mailMessages && email) {
    buildMessagesWithUserRequest(batch, email);
  }

  if (MgtPersonCardConfig.sections.files) {
    buildFilesRequest(batch, isMe ? null : email);
  }

  let response: Map<string, BatchResponse>;
  var data: MgtPersonCardState = {}; // TODO
  try {
    response = await batch.executeAll();
  } catch {
    // nop
  }

  if (response) {
    for (var [key, value] of response) {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access
      data[key] = value.content?.value || value.content;
    }
  }

  if (!isContactOrGroup && MgtPersonCardConfig.sections.profile) {
    try {
      var profile = await getProfile(graph, userId);
      if (profile) {
        data.profile = profile;
      }
    } catch {
      // nop
    }
  }

  // filter out disabled users from direct reports.
  if (data.directReports && data.directReports.length > 0) {
    data.directReports = data.directReports.filter(report => report.accountEnabled);
  }

  await cache.putValue(userId, data);

  return data;
};

var buildOrgStructureRequest = (batch: IBatch, userId: string) => {
  var expandManagers = `manager($levels=max;$select=${userProperties})`;

  batch.get(
    batchKeys.person,
    `users/${userId}?$expand=${expandManagers}&$select=${userProperties}&$count=true`,
    validUserByIdScopes,
    {
      ConsistencyLevel: 'eventual'
    }
  );

  batch.get(batchKeys.directReports, `users/${userId}/directReports?$select=${userProperties}`);
};

var validPeopleScopes = ['People.Read.All'];
var buildWorksWithRequest = (batch: IBatch, userId: string) => {
  batch.get(batchKeys.people, `users/${userId}/people?$filter=personType/class eq 'Person'`, validPeopleScopes);
};
var validMailSearchScopes = ['Mail.ReadBasic', 'Mail.Read', 'Mail.ReadWrite'];
var buildMessagesWithUserRequest = (batch: IBatch, emailAddress: string) => {
  batch.get(batchKeys.messages, `me/messages?$search="from:${emailAddress}"`, validMailSearchScopes);
};

var buildFilesRequest = (batch: IBatch, emailAddress?: string) => {
  let request: string;

  if (emailAddress) {
    request = `me/insights/shared?$filter=lastshared/sharedby/address eq '${emailAddress}'`;
  } else {
    request = 'me/insights/used';
  }

  batch.get(batchKeys.files, request, validInsightScopes);
};

var validProfileScopes = ['User.Read.All', 'User.ReadWrite.All'];
/**
 * Get the profile for a user
 *
 * @param {IGraph} graph
 * @param {string} userId
 * @return {*}  {Promise<Profile>}
 */
var getProfile = async (graph: IGraph, userId: string): Promise<Profile> =>
  (await graph
    .api(`/users/${userId}/profile`)
    .version('beta')
    .middlewareOptions(prepScopes(validProfileScopes))
    .get()) as Profile;

var validCreateChatScopes = ['Chat.Create', 'Chat.ReadWrite'];

/**
 * Initiate a chat to a user
 *
 * @export
 * @param {IGraph} graph
 * @param {{ chatType: string; members: [{"@odata.type": string,"roles": ["owner"],"user@odata.bind": string},{"@odata.type": string,"roles": ["owner"],"user@odata.bind": string}]  }} chatData
 * @return {*}  {Promise<Chat>}
 */
export var createChat = async (graph: IGraph, person: string, user: string): Promise<Chat> => {
  var chatData = {
    chatType: 'oneOnOne',
    members: [
      {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${user}')`
      },
      {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${person}')`
      }
    ]
  };
  return (await graph
    .api('/chats')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes(validCreateChatScopes))
    .post(chatData)) as Chat;
};

var validSendChatMessageScopes = ['ChatMessage.Send', 'Chat.ReadWrite'];

/**
 * Send a chat message to a user
 *
 * @export
 * @param {IGraph} graph
 * @param {{ body: {"content": string}  }} messageData
 * @return {*}  {Promise<ChatMessage>}
 */
export var sendMessage = async (
  graph: IGraph,
  chatId: string,
  messageData: Pick<ChatMessage, 'body'>
): Promise<ChatMessage> =>
  (await graph
    .api(`/chats/${chatId}/messages`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes(validSendChatMessageScopes))
    .post(messageData)) as ChatMessage;
