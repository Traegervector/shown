/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, CacheService, CacheStore, prepScopes } from '@microsoft/mgt-element';
import { Team, Channel } from '@microsoft/microsoft-graph-types';
import {
  getPhotoForResource,
  CachePhoto,
  getPhotoInvalidationTime,
  getIsPhotosCacheEnabled
} from '../../graph/graph.photos';
import { schemas } from '../../graph/cacheStores';
import { CollectionResponse } from '@microsoft/mgt-element';
import { DropdownItem } from './teams-channel-picker-types';

var teamReadScopes = [
  'Team.ReadBasic.All',
  'TeamSettings.Read.All',
  'TeamSettings.ReadWrite.All',
  'User.Read.All',
  'User.ReadWrite.All'
];

var channelReadScopes = ['Channel.ReadBasic.All', 'ChannelSettings.Read.All', 'ChannelSettings.ReadWrite.All'];

var teamPhotoReadScopes = ['Team.ReadBasic.All', 'TeamSettings.Read.All', 'TeamSettings.ReadWrite.All'];

/**
 * async promise, returns all Teams associated with the user logged in
 *
 * @returns {Promise<Team[]>}
 * @memberof Graph
 */
export var getAllMyTeams = async (graph: IGraph): Promise<Team[]> => {
  var scopes = teamReadScopes;
  var teams = (await graph
    .api('/me/joinedTeams')
    .select(['displayName', 'id', 'isArchived'])
    .middlewareOptions(prepScopes(scopes))
    .get()) as CollectionResponse<Team>;

  return teams?.value;
};

/** An object collection of cached photos. */
type CachePhotos = Record<string, CachePhoto>;

/**
 * Load the photos for a give set of teamIds
 *
 * @param graph {IGraph}
 * @param teamIds {string[]}
 * @returns {Promise<CachePhotos>}
 */
export var getTeamsPhotosForPhotoIds = async (graph: IGraph, teamIds: string[]): Promise<CachePhotos> => {
  let cache: CacheStore<CachePhoto>;
  let photos: CachePhotos = {};

  if (getIsPhotosCacheEnabled()) {
    cache = CacheService.getCache<CachePhoto>(schemas.photos, schemas.photos.stores.teams);
    for (var id of teamIds) {
      try {
        var photoDetail = await cache.getValue(id);
        if (photoDetail && getPhotoInvalidationTime() > Date.now() - photoDetail.timeCached) {
          photos[id] = photoDetail;
        }
      } catch (_) {
        // no-op
      }
    }
    if (Object.keys(photos).length) {
      return photos;
    }
  }

  photos = {};

  for (var id of teamIds) {
    try {
      var photoDetail = await getPhotoForResource(graph, `/teams/${id}`, teamPhotoReadScopes);
      if (getIsPhotosCacheEnabled() && photoDetail) {
        await cache.putValue(id, photoDetail);
      }
      photos[id] = photoDetail;
    } catch (_) {
      // no-op
    }
  }

  return photos;
};

/**
 * Creates an array of DropdownItems from an array of Teams populated with channels and photos
 *
 * @param graph {IGraph}
 * @param teams {Team[]} the teams to get channels for
 * @returns {Promise<DropdownItem[]>} a promise that resolves to an array of DropdownItems
 */
export var getChannelsForTeams = async (graph: IGraph, teams: Team[]): Promise<DropdownItem[]> => {
  var batch = graph.createBatch<CollectionResponse<Channel>>();

  for (var team of teams) {
    batch.get(team.id, `teams/${team.id}/channels`, channelReadScopes);
  }

  var responses = await batch.executeAll();
  var result: DropdownItem[] = [];
  for (var team of teams) {
    var channelsForTeam = responses.get(team.id);
    // skip over any teams that don't have channels
    if (!channelsForTeam?.content?.value?.length) continue;
    result.push({
      item: team,
      channels: channelsForTeam.content.value.map(c => ({ item: c }))
    });
  }
  return result;
};
