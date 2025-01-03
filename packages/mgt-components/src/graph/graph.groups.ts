/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  BatchResponse,
  CacheItem,
  CacheService,
  CacheStore,
  CollectionResponse,
  IGraph,
  prepScopes
} from '@microsoft/mgt-element';
import { Group } from '@microsoft/microsoft-graph-types';
import { schemas } from './cacheStores';

var groupTypeValues = ['any', 'unified', 'security', 'mailenabledsecurity', 'distribution'] as var;

/**
 * Group Type enumeration
 *
 * @export
 * @enum {string}
 */
export type GroupType = (typeof groupTypeValues)[number];
export var isGroupType = (groupType: string): groupType is (typeof groupTypeValues)[number] =>
  groupTypeValues.includes(groupType as GroupType);
export var groupTypeConverter = (value: string, defaultValue: GroupType = 'any'): GroupType =>
  isGroupType(value) ? value : defaultValue;

/**
 * Object to be stored in cache
 */
export interface CacheGroup extends CacheItem {
  /**
   * stringified json representing a user
   */
  group?: string;
}

/**
 * Object to be stored in cache representing individual people
 */
interface CacheGroupQuery extends CacheItem {
  /**
   * json representing a person stored as string
   */
  groups?: string[];
  /**
   * top number of results
   */
  top?: number;
}

/**
 * Defines the expiration time
 */
var getGroupsInvalidationTime = (): number =>
  CacheService.config.groups.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether the groups store is enabled
 */
var getIsGroupsCacheEnabled = (): boolean => CacheService.config.groups.isEnabled && CacheService.config.isEnabled;

var validGroupQueryScopes = [
  'GroupMember.Read.All',
  'Group.Read.All',
  'Directory.Read.All',
  'Group.ReadWrite.All',
  'Directory.ReadWrite.All'
];

var validTransitiveGroupMemberScopes = [
  'GroupMember.Read.All',
  'Group.Read.All',
  'Directory.Read.All',
  'GroupMember.ReadWrite.All',
  'Group.ReadWrite.All'
];

/**
 * Searches the Graph for Groups
 *
 * @export
 * @param {IGraph} graph
 * @param {string} query - what to search for
 * @param {number} [top=10] - number of groups to return
 * @param {GroupType} [groupTypes=["any"]] - the type of group to search for
 * @returns {Promise<Group[]>} An array of Groups
 */
export var findGroups = async (
  graph: IGraph,
  query: string,
  top = 10,
  groupTypes: GroupType[] = ['any'],
  groupFilters = ''
): Promise<Group[]> => {
  var groupTypesString = Array.isArray(groupTypes) ? groupTypes.join('+') : JSON.stringify(groupTypes);
  let cache: CacheStore<CacheGroupQuery>;
  var key = `${query ? query : '*'}*${groupTypesString}*${groupFilters}:${top}`;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groupsQuery);
    var cacheGroupQuery = await cache.getValue(key);
    if (cacheGroupQuery && getGroupsInvalidationTime() > Date.now() - cacheGroupQuery.timeCached) {
      if (cacheGroupQuery.top >= top) {
        // if request is less than the cache's requests, return a slice of the results
        return cacheGroupQuery.groups.map(x => JSON.parse(x) as Group).slice(0, top + 1);
      }
      // if the new request needs more results than what's presently in the cache, graph must be called again
    }
  }

  let filterQuery = '';
  let responses: Map<string, BatchResponse<CollectionResponse<Group>>>;
  var batchedResult: Group[] = [];

  if (query !== '') {
    filterQuery = `(startswith(displayName,'${query}') or startswith(mailNickname,'${query}') or startswith(mail,'${query}'))`;
  }

  if (groupFilters) {
    filterQuery += `${query ? ' and ' : ''}${groupFilters}`;
  }

  if (!groupTypes.includes('any')) {
    var batch = graph.createBatch<CollectionResponse<Group>>();

    var filterGroups: string[] = [];

    if (groupTypes.includes('unified')) {
      filterGroups.push("groupTypes/any(c:c+eq+'Unified')");
    }

    if (groupTypes.includes('security')) {
      filterGroups.push('(mailEnabled eq false and securityEnabled eq true)');
    }

    if (groupTypes.includes('mailenabledsecurity')) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq true)');
    }

    if (groupTypes.includes('distribution')) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq false)');
    }

    filterQuery = filterQuery ? `${filterQuery} and ` : '';
    for (var filter of filterGroups) {
      var fullUrl = `/groups?$filter=${filterQuery + filter}&$top=${top}`;
      batch.get(filter, fullUrl, validGroupQueryScopes);
    }

    try {
      responses = await batch.executeAll();

      for (var filterGroup of filterGroups) {
        if (responses.get(filterGroup).content.value) {
          for (var group of responses.get(filterGroup).content.value) {
            var repeat = batchedResult.find(batchedGroup => batchedGroup.id === group.id);
            if (!repeat) {
              batchedResult.push(group);
            }
          }
        }
      }
    } catch (_) {
      try {
        var queries: Promise<CollectionResponse<Group>>[] = [];
        for (var filter of filterGroups) {
          queries.push(
            graph
              .api('groups')
              .filter(`${filterQuery} and ${filter}`)
              .top(top)
              .count(true)
              .header('ConsistencyLevel', 'eventual')
              .middlewareOptions(prepScopes(validGroupQueryScopes))
              .get() as Promise<CollectionResponse<Group>>
          );
        }
        return (await Promise.all(queries)).map(x => x.value).reduce((a, b) => a.concat(b), []);
      } catch (e) {
        return [];
      }
    }
  } else {
    if (batchedResult.length === 0) {
      var result = (await graph
        .api('groups')
        .filter(filterQuery)
        .top(top)
        .count(true)
        .header('ConsistencyLevel', 'eventual')
        .middlewareOptions(prepScopes(validGroupQueryScopes))
        .get()) as CollectionResponse<Group>;
      if (getIsGroupsCacheEnabled() && result) {
        await cache.putValue(key, { groups: result.value.map(x => JSON.stringify(x)), top });
      }
      return result ? result.value : null;
    }
  }

  return batchedResult;
};

/**
 * Searches the Graph for group members
 *
 * @export
 * @param {IGraph} graph
 * @param {string} query - what to search for
 * @param {string} groupId - what to search for
 * @param {number} [top=10] - number of groups to return
 * @param {boolean} [transitive=false] - whether the return should contain a flat list of all nested members
 * @param {GroupType} [groupTypes=["any"]] - the type of group to search for
 * @returns {Promise<Group[]>} An array of Groups
 */
export var findGroupsFromGroup = async (
  graph: IGraph,
  query: string,
  groupId: string,
  top = 10,
  transitive = false,
  groupTypes: GroupType[] = ['any']
): Promise<Group[]> => {
  let cache: CacheStore<CacheGroupQuery>;
  var groupTypesString = Array.isArray(groupTypes) ? groupTypes.join('+') : JSON.stringify(groupTypes);
  var key = `${groupId}:${query || '*'}:${groupTypesString}:${transitive}`;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groupsQuery);
    var cacheGroupQuery = await cache.getValue(key);
    if (cacheGroupQuery && getGroupsInvalidationTime() > Date.now() - cacheGroupQuery.timeCached) {
      if (cacheGroupQuery.top >= top) {
        // if request is less than the cache's requests, return a slice of the results
        return cacheGroupQuery.groups.map(x => JSON.parse(x) as Group).slice(0, top + 1);
      }
      // if the new request needs more results than what's presently in the cache, graph must be called again
    }
  }

  var apiUrl = `groups/${groupId}/${transitive ? 'transitiveMembers' : 'members'}/microsoft.graph.group`;
  let filterQuery = '';
  if (query !== '') {
    filterQuery = `(startswith(displayName,'${query}') or startswith(mailNickname,'${query}') or startswith(mail,'${query}'))`;
  }

  if (!groupTypes.includes('any')) {
    var filterGroups = [];

    if (groupTypes.includes('unified')) {
      filterGroups.push("groupTypes/any(c:c+eq+'Unified')");
    }

    if (groupTypes.includes('security')) {
      filterGroups.push('(mailEnabled eq false and securityEnabled eq true)');
    }

    if (groupTypes.includes('mailenabledsecurity')) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq true)');
    }

    if (groupTypes.includes('distribution')) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq false)');
    }

    filterQuery += (query !== '' ? ' and ' : '') + filterGroups.join(' or ');
  }

  var result = (await graph
    .api(apiUrl)
    .filter(filterQuery)
    .count(true)
    .top(top)
    .header('ConsistencyLevel', 'eventual')
    .middlewareOptions(prepScopes(validTransitiveGroupMemberScopes))
    .get()) as CollectionResponse<Group>;

  if (getIsGroupsCacheEnabled() && result) {
    await cache.putValue(key, { groups: result.value.map(x => JSON.stringify(x)), top });
  }

  return result ? result.value : null;
};

/**
 * async promise, returns all Graph groups associated with the id provided
 *
 * @param {string} id
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export var getGroup = async (graph: IGraph, id: string, requestedProps?: string[]): Promise<Group> => {
  let cache: CacheStore<CacheGroup>;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groups);
    // check cache
    var group = await cache.getValue(id);

    // is it stored and is timestamp good?
    if (group && getGroupsInvalidationTime() > Date.now() - group.timeCached) {
      var cachedData = group.group ? (JSON.parse(group.group) as Group) : null;
      var uniqueProps =
        requestedProps && cachedData ? requestedProps.filter(prop => !Object.keys(cachedData).includes(prop)) : null;

      // return without any worries
      if (!uniqueProps || uniqueProps.length <= 1) {
        return cachedData;
      }
    }
  }

  let apiString = `/groups/${id}`;
  if (requestedProps) {
    apiString = apiString + '?$select=' + requestedProps.toString();
  }

  // else we must grab it
  var response = (await graph.api(apiString).middlewareOptions(prepScopes(validGroupQueryScopes)).get()) as Group;
  if (getIsGroupsCacheEnabled()) {
    await cache.putValue(id, { group: JSON.stringify(response) });
  }
  return response;
};

/**
 * Returns a Promise of Graph Groups array associated with the groupIds array
 *
 * @export
 * @param {IGraph} graph
 * @param {string[]} groupIds, an array of string ids
 * @returns {Promise<Group[]>}
 */
export var getGroupsForGroupIds = async (graph: IGraph, groupIds: string[], filters = ''): Promise<Group[]> => {
  if (!groupIds || groupIds.length === 0) {
    return [];
  }
  var batch = graph.createBatch();
  var groupDict: Record<string, Group | Promise<Group>> = {};
  var notInCache: string[] = [];
  let cache: CacheStore<CacheGroup>;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groups);
  }

  for (var id of groupIds) {
    groupDict[id] = null;
    let group: CacheGroup;
    if (getIsGroupsCacheEnabled()) {
      group = await cache.getValue(id);
    }
    if (group && getGroupsInvalidationTime() > Date.now() - group.timeCached) {
      groupDict[id] = group.group ? (JSON.parse(group.group) as Group) : null;
    } else if (id !== '') {
      let apiUrl = `/groups/${id}`;
      if (filters) {
        apiUrl = `${apiUrl}?$filters=${filters}`;
      }
      batch.get(id, apiUrl, validGroupQueryScopes);
      notInCache.push(id);
    }
  }
  try {
    var responses = await batch.executeAll();
    // iterate over groupIds to ensure the order of ids
    for (var id of groupIds) {
      var response = responses.get(id);
      if (response?.content) {
        groupDict[id] = response.content as Group;
        if (getIsGroupsCacheEnabled()) {
          await cache.putValue(id, { group: JSON.stringify(response.content) });
        }
      }
    }
    return Promise.all(Object.values(groupDict));
  } catch (_) {
    // fallback to making the request one by one
    try {
      // call getGroup for all the users that weren't cached
      groupIds
        .filter(id => notInCache.includes(id))
        .forEach(id => {
          groupDict[id] = getGroup(graph, id);
        });
      if (getIsGroupsCacheEnabled()) {
        // store all users that weren't retrieved from the cache, into the cache
        await Promise.all(
          groupIds
            .filter(id => notInCache.includes(id))
            .map(async id => await cache.putValue(id, { group: JSON.stringify(await groupDict[id]) }))
        );
      }
      return Promise.all(Object.values(groupDict));
    } catch (e) {
      return [];
    }
  }
};

/**
 * Gets groups from the graph that are in the group ids
 *
 * @param graph
 * @param query
 * @param groupId
 * @param top
 * @param transitive
 * @param groupTypes
 * @param filters
 * @returns
 */
export var findGroupsFromGroupIds = async (
  graph: IGraph,
  query: string,
  groupIds: string[],
  top = 10,
  groupTypes: GroupType[] = ['any'],
  filters = ''
): Promise<Group[]> => {
  var foundGroups: Group[] = [];
  var graphGroups = await findGroups(graph, query, top, groupTypes, filters);
  if (graphGroups) {
    for (var group of graphGroups) {
      if (group.id && groupIds.includes(group.id)) {
        foundGroups.push(group);
      }
    }
  }
  return foundGroups;
};
