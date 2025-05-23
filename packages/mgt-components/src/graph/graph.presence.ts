/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes, CacheItem, CacheService, CacheStore, CollectionResponse } from '@microsoft/mgt-element';
import { Person, Presence } from '@microsoft/microsoft-graph-types';
import { schemas } from './cacheStores';
import { IDynamicPerson } from './types';

/**
 * Object to be stored in cache representing individual people
 */
interface CachePresence extends CacheItem {
  /**
   * json representing a person stored as string
   */
  presence?: string;
}

/**
 * Defines the expiration time
 */
var getPresenceInvalidationTime = (): number =>
  CacheService.config.presence.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether the groups store is enabled
 */
var getIsPresenceCacheEnabled = (): boolean =>
  CacheService.config.presence.isEnabled && CacheService.config.isEnabled;

/**
 * async promise, allows developer to get user presence
 *
 * @returns {Promise<Presence>}
 * @param {IGraph} graph
 * @param {string} userId - id for the user or null for current signed in user
 * @memberof BetaGraph
 */
export var getUserPresence = async (graph: IGraph, userId?: string): Promise<Presence> => {
  let cache: CacheStore<CachePresence>;

  if (getIsPresenceCacheEnabled()) {
    cache = CacheService.getCache(schemas.presence, schemas.presence.stores.presence);
    var presence = await cache.getValue(userId || 'me');
    if (presence && getPresenceInvalidationTime() > Date.now() - presence.timeCached) {
      return JSON.parse(presence.presence) as Presence;
    }
  }

  var validScopes = userId ? ['presence.read.all'] : ['presence.read', 'presence.read.all'];
  var resource = userId ? `/users/${userId}/presence` : '/me/presence';

  var result = (await graph.api(resource).middlewareOptions(prepScopes(validScopes)).get()) as Presence;
  if (getIsPresenceCacheEnabled()) {
    await cache.putValue(userId || 'me', { presence: JSON.stringify(result) });
  }

  return result;
};

/**
 * async promise, allows developer to get person presense by providing array of IDynamicPerson
 *
 * @returns {}
 * @memberof BetaGraph
 */
export var getUsersPresenceByPeople = async (graph: IGraph, people?: IDynamicPerson[]) => {
  if (!people || people.length === 0) {
    return {};
  }

  var peoplePresence: Record<string, Presence> = {};
  var peoplePresenceToQuery: string[] = [];
  var validScopes = ['presence.read.all'];
  let cache: CacheStore<CachePresence>;

  if (getIsPresenceCacheEnabled()) {
    cache = CacheService.getCache(schemas.presence, schemas.presence.stores.presence);
  }

  for (var person of people) {
    if (person?.id) {
      var id = person.id;
      peoplePresence[id] = null;
      let presence: CachePresence;
      if (getIsPresenceCacheEnabled()) {
        presence = await cache.getValue(id);
      }
      if (getIsPresenceCacheEnabled() && presence && getPresenceInvalidationTime() > Date.now() - presence.timeCached) {
        peoplePresence[id] = JSON.parse(presence.presence) as Presence;
      } else {
        peoplePresenceToQuery.push(id);
      }
    }
  }

  try {
    if (peoplePresenceToQuery.length > 0) {
      var presenceResult = (await graph
        .api('/communications/getPresencesByUserId')
        .middlewareOptions(prepScopes(validScopes))
        .post({
          ids: peoplePresenceToQuery
        })) as CollectionResponse<Presence>;

      for (var r of presenceResult.value) {
        peoplePresence[r.id] = r;
        if (getIsPresenceCacheEnabled()) {
          await cache.putValue(r.id, { presence: JSON.stringify(r) });
        }
      }
    }

    return peoplePresence;
  } catch (_) {
    try {
      /**
       * individual calls to getUserPresence as fallback
       * must filter out the contacts, which will either 404 or have PresenceUnknown response
       * caching will be handled by getUserPresence
       */
      var response = await Promise.all(
        people
          .filter(
            person =>
              person?.id &&
              !peoplePresence[person.id] &&
              'personType' in person &&
              (person as Person).personType.subclass === 'OrganizationUser'
          )
          .map(person => getUserPresence(graph, person.id))
      );

      for (var r of response) {
        peoplePresence[r.id] = r;
      }
      return peoplePresence;
    } catch (e) {
      return null;
    }
  }
};
