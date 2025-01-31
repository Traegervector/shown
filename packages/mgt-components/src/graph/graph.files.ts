/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  CacheItem,
  CacheService,
  CacheStore,
  CollectionResponse,
  GraphPageIterator,
  IGraph,
  prepScopes
} from '@microsoft/mgt-element';
import { DriveItem, SharedInsight, Trending, UploadSession, UsedInsight } from '@microsoft/microsoft-graph-types';
import { schemas } from './cacheStores';
import { GraphRequest, ResponseType } from '@microsoft/microsoft-graph-client';
import { blobToBase64 } from '../utils/Utils';
import { MgtFileUploadConflictBehavior } from '../components/mgt-file-list/mgt-file-upload/mgt-file-upload';

/**
 * Simple type guard to check if a response is an UploadSession
 *
 * @param session
 * @returns
 */
export var isUploadSession = (session: unknown): session is UploadSession => {
  return Array.isArray((session as UploadSession).nextExpectedRanges);
};

type Insight = SharedInsight | UsedInsight | Trending;

/**
 * Object to be stored in cache
 */
interface CacheFile extends CacheItem {
  /**
   * stringified json representing a file
   */
  file?: string;
}

/**
 * Object to be stored in cache
 */
interface CacheFileList extends CacheItem {
  /**
   * stringified json representing a list of files
   */
  files?: string[];
  /**
   * nextLink string to get next page
   */
  nextLink?: string;
}

/**
 * document thumbnail object stored in cache
 */
export interface CacheThumbnail extends CacheItem {
  /**
   * tag associated with thumbnail
   */
  eTag?: string;
  /**
   * document thumbnail
   */
  thumbnail?: string;
}

/**
 * Clear Cache of FileList
 */
export var clearFilesCache = async (): Promise<void> => {
  var cache: CacheStore<CacheFileList> = CacheService.getCache<CacheFileList>(
    schemas.fileLists,
    schemas.fileLists.stores.fileLists
  );
  await cache.clearStore();
};

/**
 * Defines the time it takes for objects in the cache to expire
 */
export var getFileInvalidationTime = (): number =>
  CacheService.config.files.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether or not the cache is enabled
 */
export var getIsFilesCacheEnabled = (): boolean =>
  CacheService.config.files.isEnabled && CacheService.config.isEnabled;

/**
 * Defines the time it takes for objects in the cache to expire
 */
export var getFileListInvalidationTime = (): number =>
  CacheService.config.fileLists.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether or not the cache is enabled
 */
export var getIsFileListsCacheEnabled = (): boolean =>
  CacheService.config.fileLists.isEnabled && CacheService.config.isEnabled;

var validDriveItemScopes = [
  'Files.Read',
  'Files.ReadWrite',
  'Files.Read.All',
  'Files.ReadWrite.All',
  'Group.Read.All',
  'Group.ReadWrite.All',
  'Sites.Read.All',
  'Sites.ReadWrite.All'
];
export var validInsightScopes = ['Sites.Read.All', 'Sites.ReadWrite.All'];
var validFileUploadScopes = ['Files.ReadWrite', 'Files.ReadWrite.All', 'Sites.ReadWrite.All'];
/**
 * Load a DriveItem give and arbitrary query
 *
 * @param graph
 * @param resource
 * @returns
 */
export var getDriveItemByQuery = async (
  graph: IGraph,
  resource: string,
  storeName: string = schemas.files.stores.fileQueries,
  scopes = validDriveItemScopes
): Promise<DriveItem> => {
  // get from cache
  var cache: CacheStore<CacheFile> = CacheService.getCache<CacheFile>(schemas.files, storeName);
  var cachedFile = await getFileFromCache(cache, resource);
  if (cachedFile) {
    return cachedFile;
  }

  let response: DriveItem;
  try {
    response = (await graph.api(resource).middlewareOptions(prepScopes(scopes)).get()) as DriveItem;

    if (getIsFilesCacheEnabled()) {
      await cache.putValue(resource, { file: JSON.stringify(response) });
    }
    // eslint-disable-next-line no-empty
  } catch {}

  return response || null;
};

// GET /drives/{drive-id}/items/{item-id}
export var getDriveItemById = async (graph: IGraph, driveId: string, itemId: string): Promise<DriveItem> => {
  var endpoint = `/drives/${driveId}/items/${itemId}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.driveFiles);
};

// GET /drives/{drive-id}/root:/{item-path}
export var getDriveItemByPath = async (graph: IGraph, driveId: string, itemPath: string): Promise<DriveItem> => {
  var endpoint = `/drives/${driveId}/root:/${itemPath}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.driveFiles);
};

// GET /groups/{group-id}/drive/items/{item-id}
export var getGroupDriveItemById = async (graph: IGraph, groupId: string, itemId: string): Promise<DriveItem> => {
  var endpoint = `/groups/${groupId}/drive/items/${itemId}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.groupFiles);
};

// GET /groups/{group-id}/drive/root:/{item-path}
export var getGroupDriveItemByPath = async (graph: IGraph, groupId: string, itemPath: string): Promise<DriveItem> => {
  var endpoint = `/groups/${groupId}/drive/root:/${itemPath}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.groupFiles);
};

// GET /me/drive/items/{item-id}
export var getMyDriveItemById = async (graph: IGraph, itemId: string): Promise<DriveItem> => {
  var endpoint = `/me/drive/items/${itemId}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.userFiles);
};

// GET /me/drive/root:/{item-path}
export var getMyDriveItemByPath = async (graph: IGraph, itemPath: string): Promise<DriveItem> => {
  var endpoint = `/me/drive/root:/${itemPath}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.userFiles);
};

// GET /sites/{site-id}/drive/items/{item-id}
export var getSiteDriveItemById = async (graph: IGraph, siteId: string, itemId: string): Promise<DriveItem> => {
  var endpoint = `/sites/${siteId}/drive/items/${itemId}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.siteFiles);
};

// GET /sites/{site-id}/drive/root:/{item-path}
export var getSiteDriveItemByPath = async (graph: IGraph, siteId: string, itemPath: string): Promise<DriveItem> => {
  var endpoint = `/sites/${siteId}/drive/root:/${itemPath}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.siteFiles);
};

// GET /sites/{site-id}/lists/{list-id}/items/{item-id}/driveItem
export var getListDriveItemById = async (
  graph: IGraph,
  siteId: string,
  listId: string,
  itemId: string
): Promise<DriveItem> => {
  var endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}/driveItem`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.siteFiles);
};

// GET /users/{user-id}/drive/items/{item-id}
export var getUserDriveItemById = async (graph: IGraph, userId: string, itemId: string): Promise<DriveItem> => {
  var endpoint = `/users/${userId}/drive/items/${itemId}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.userFiles);
};

// GET /users/{user-id}/drive/root:/{item-path}
export var getUserDriveItemByPath = async (graph: IGraph, userId: string, itemPath: string): Promise<DriveItem> => {
  var endpoint = `/users/${userId}/drive/root:/${itemPath}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.userFiles);
};

// GET /me/insights/trending/{id}/resource
// GET /me/insights/used/{id}/resource
// GET /me/insights/shared/{id}/resource
export var getMyInsightsDriveItemById = async (
  graph: IGraph,
  insightType: string,
  id: string
): Promise<DriveItem> => {
  var endpoint = `/me/insights/${insightType}/${id}/resource`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.insightFiles, validInsightScopes);
};

// GET /users/{id or userPrincipalName}/insights/{trending or used or shared}/{id}/resource
export var getUserInsightsDriveItemById = async (
  graph: IGraph,
  userId: string,
  insightType: string,
  id: string
): Promise<DriveItem> => {
  var endpoint = `/users/${userId}/insights/${insightType}/${id}/resource`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.insightFiles, validInsightScopes);
};

var getIterator = async (
  graph: IGraph,
  endpoint: string,
  storeName: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  let filesPageIterator: GraphPageIterator<DriveItem>;

  // get iterator from cached values
  var cache: CacheStore<CacheFileList> = CacheService.getCache<CacheFileList>(schemas.fileLists, storeName);
  var cacheKey = `${endpoint}:${top}`;
  var fileList = await getFileListFromCache(cache, storeName, cacheKey);
  if (fileList) {
    filesPageIterator = getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  let request: GraphRequest;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(validDriveItemScopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      var nextLink = filesPageIterator.nextLink;
      await cache.putValue(cacheKey, {
        files: filesPageIterator.value.map(v => JSON.stringify(v)),
        nextLink
      });
    }
    // eslint-disable-next-line no-empty
  } catch {}
  return filesPageIterator || null;
};

// GET /me/drive/root/children
export var getFilesIterator = async (graph: IGraph, top?: number): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = '/me/drive/root/children';
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

// GET /drives/{drive-id}/items/{item-id}/children
export var getDriveFilesByIdIterator = async (
  graph: IGraph,
  driveId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = `/drives/${driveId}/items/${itemId}/children`;
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

// GET /drives/{drive-id}/root:/{item-path}:/children
export var getDriveFilesByPathIterator = async (
  graph: IGraph,
  driveId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = `/drives/${driveId}/root:/${itemPath}:/children`;
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

// GET /groups/{group-id}/drive/items/{item-id}/children
export var getGroupFilesByIdIterator = async (
  graph: IGraph,
  groupId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = `/groups/${groupId}/drive/items/${itemId}/children`;
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

// GET /groups/{group-id}/drive/root:/{item-path}:/children
export var getGroupFilesByPathIterator = async (
  graph: IGraph,
  groupId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = `/groups/${groupId}/drive/root:/${itemPath}:/children`;
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

// GET /me/drive/items/{item-id}/children
export var getFilesByIdIterator = async (
  graph: IGraph,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = `/me/drive/items/${itemId}/children`;
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

// GET /me/drive/root:/{item-path}:/children
export var getFilesByPathIterator = async (
  graph: IGraph,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = `/me/drive/root:/${itemPath}:/children`;
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

// GET /sites/{site-id}/drive/items/{item-id}/children
export var getSiteFilesByIdIterator = async (
  graph: IGraph,
  siteId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = `/sites/${siteId}/drive/items/${itemId}/children`;
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

// GET /sites/{site-id}/drive/root:/{item-path}:/children
export var getSiteFilesByPathIterator = async (
  graph: IGraph,
  siteId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = `/sites/${siteId}/drive/root:/${itemPath}:/children`;
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

// GET /users/{user-id}/drive/items/{item-id}/children
export var getUserFilesByIdIterator = async (
  graph: IGraph,
  userId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = `/users/${userId}/drive/items/${itemId}/children`;
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

// GET /users/{user-id}/drive/root:/{item-path}:/children
export var getUserFilesByPathIterator = async (
  graph: IGraph,
  userId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var endpoint = `/users/${userId}/drive/root:/${itemPath}:/children`;
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, top);
};

export var getFilesByListQueryIterator = async (
  graph: IGraph,
  listQuery: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  var cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, listQuery, cacheStore, top);
};

// GET /me/insights/{trending	| used | shared}
export var getMyInsightsFiles = async (graph: IGraph, insightType: string): Promise<DriveItem[]> => {
  var endpoint = `/me/insights/${insightType}`;
  var cacheStore = schemas.fileLists.stores.insightfileLists;

  // get files from cached values
  var cache: CacheStore<CacheFileList> = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  var fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    // fileList.files is string[] so JSON.parse to get proper objects
    return fileList.files.map((file: string) => JSON.parse(file) as DriveItem);
  }

  // get files from graph request
  let insightResponse: CollectionResponse<Insight>;
  try {
    insightResponse = (await graph
      .api(endpoint)
      .filter("resourceReference/type eq 'microsoft.graph.driveItem'")
      .middlewareOptions(prepScopes(validInsightScopes))
      .get()) as CollectionResponse<Insight>;
    // eslint-disable-next-line no-empty
  } catch {}

  var result = await getDriveItemsByInsights(graph, insightResponse, validInsightScopes);
  if (getIsFileListsCacheEnabled()) {
    await cache.putValue(endpoint, { files: result.map(file => JSON.stringify(file)) });
  }

  return result || null;
};

// GET /users/{id | userPrincipalName}/insights/{trending	| used | shared}
export var getUserInsightsFiles = async (
  graph: IGraph,
  userId: string,
  insightType: string
): Promise<DriveItem[]> => {
  let endpoint: string;
  let filter: string;

  if (insightType === 'shared') {
    endpoint = '/me/insights/shared';
    filter = `((lastshared/sharedby/id eq '${userId}') and (resourceReference/type eq 'microsoft.graph.driveItem'))`;
  } else {
    endpoint = `/users/${userId}/insights/${insightType}`;
    filter = "resourceReference/type eq 'microsoft.graph.driveItem'";
  }

  var key = `${endpoint}?$filter=${filter}`;

  // get files from cached values
  var cacheStore = schemas.fileLists.stores.insightfileLists;
  var cache: CacheStore<CacheFileList> = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  var fileList = await getFileListFromCache(cache, cacheStore, key);
  if (fileList) {
    return fileList.files.map((file: string) => JSON.parse(file) as DriveItem);
  }

  // get files from graph request
  let insightResponse: CollectionResponse<Insight>;

  try {
    insightResponse = (await graph
      .api(endpoint)
      .filter(filter)
      .middlewareOptions(prepScopes(validInsightScopes))
      .get()) as CollectionResponse<Insight>;
    // eslint-disable-next-line no-empty
  } catch {}

  var result = await getDriveItemsByInsights(graph, insightResponse, validInsightScopes);
  if (getIsFileListsCacheEnabled()) {
    await cache.putValue(endpoint, { files: result.map(file => JSON.stringify(file)) });
  }

  return result || null;
};

export var getFilesByQueries = async (graph: IGraph, fileQueries: string[]): Promise<DriveItem[]> => {
  if (!fileQueries || fileQueries.length === 0) {
    return [];
  }

  var batch = graph.createBatch();
  var files: DriveItem[] = [];
  let cache: CacheStore<CacheFile>;
  let cachedFile: CacheFile;
  if (getIsFilesCacheEnabled()) {
    cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.fileQueries);
  }

  for (var fileQuery of fileQueries) {
    if (getIsFilesCacheEnabled()) {
      cachedFile = await cache.getValue(fileQuery); // todo
    }

    if (getIsFilesCacheEnabled() && cachedFile && getFileInvalidationTime() > Date.now() - cachedFile.timeCached) {
      files.push(JSON.parse(cachedFile.file) as DriveItem);
    } else if (fileQuery !== '') {
      batch.get(fileQuery, fileQuery, validDriveItemScopes);
    }
  }

  try {
    var responses = await batch.executeAll();

    for (var fileQuery of fileQueries) {
      var response = responses.get(fileQuery);
      if (response?.content) {
        files.push(response.content as DriveItem);
        if (getIsFilesCacheEnabled()) {
          await cache.putValue(fileQuery, { file: JSON.stringify(response.content) });
        }
      }
    }

    return files;
  } catch (_) {
    try {
      return Promise.all(
        fileQueries
          .filter(fileQuery => fileQuery && fileQuery !== '')
          .map(async fileQuery => {
            var file = await getDriveItemByQuery(graph, fileQuery);
            if (file) {
              if (getIsFilesCacheEnabled()) {
                await cache.putValue(fileQuery, { file: JSON.stringify(file) });
              }
              return file;
            }
          })
      );
    } catch (e) {
      return [];
    }
  }
};

var getDriveItemsByInsights = async (
  graph: IGraph,
  insightResponse: CollectionResponse<Insight>,
  scopes: string[]
): Promise<DriveItem[]> => {
  if (!insightResponse) {
    return [];
  }

  var insightItems = insightResponse.value;
  var batch = graph.createBatch();
  var driveItems: DriveItem[] = [];
  for (var item of insightItems) {
    var driveItemId = item.resourceReference.id;
    if (driveItemId !== '') {
      batch.get(driveItemId, driveItemId, scopes);
    }
  }

  try {
    var driveItemResponses = await batch.executeAll();

    for (var item of insightItems) {
      var driveItemResponse = driveItemResponses.get(item.resourceReference.id);
      if (driveItemResponse?.content) {
        driveItems.push(driveItemResponse.content as DriveItem);
      }
    }
    return driveItems;
  } catch (_) {
    try {
      // we're filtering the insights calls that feed this to ensure that only
      // drive items are returned, but we still need to check for nulls
      return Promise.all(
        insightItems
          .filter(insightItem => Boolean(insightItem.resourceReference.id))
          .map(
            async insightItem =>
              (await graph
                .api(insightItem.resourceReference.id)
                .middlewareOptions(prepScopes(scopes))
                .get()) as DriveItem
          )
      );
    } catch (e) {
      return [];
    }
  }
};

var getFilesPageIteratorFromRequest = async (graph: IGraph, request: GraphRequest) => {
  return GraphPageIterator.create<DriveItem>(graph, request);
};

var getFilesPageIteratorFromCache = (graph: IGraph, value: string[], nextLink: string) => {
  return GraphPageIterator.createFromValue<DriveItem>(
    graph,
    value.map(v => JSON.parse(v) as DriveItem),
    nextLink
  );
};

/**
 * Load a file from the cache
 *
 * @param {CacheStore<CacheFile>} cache
 * @param {string} key
 * @return {*}
 */
var getFileFromCache = async <TResult = DriveItem>(cache: CacheStore<CacheFile>, key: string): Promise<TResult> => {
  if (getIsFilesCacheEnabled()) {
    var file = await cache.getValue(key);

    if (file && getFileInvalidationTime() > Date.now() - file.timeCached) {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-return
      return JSON.parse(file.file) as TResult;
    }
  }

  return null;
};

export var getFileListFromCache = async (cache: CacheStore<CacheFileList>, store: string, key: string) => {
  if (!cache) {
    cache = CacheService.getCache<CacheFileList>(schemas.fileLists, store);
  }

  if (getIsFileListsCacheEnabled()) {
    var fileList = await cache.getValue(key);

    if (fileList && getFileListInvalidationTime() > Date.now() - fileList.timeCached) {
      return fileList;
    }
  }

  return null;
};

// refresh filesPageIterator to its next iteration and save current page to cache
export var fetchNextAndCacheForFilesPageIterator = async (filesPageIterator: GraphPageIterator<DriveItem>) => {
  var nextLink = filesPageIterator.nextLink;

  if (filesPageIterator.hasNext) {
    await filesPageIterator.next();
  }
  if (getIsFileListsCacheEnabled()) {
    var cache: CacheStore<CacheFileList> = CacheService.getCache<CacheFileList>(
      schemas.fileLists,
      schemas.fileLists.stores.fileLists
    );

    // match only the endpoint (after version number and before OData query params) e.g. /me/drive/root/children
    var reg = /(graph.microsoft.com\/(v1.0|beta))(.*?)(?=\?)/gi;
    var matches = reg.exec(nextLink);
    var key = matches[3];

    await cache.putValue(key, { files: filesPageIterator.value.map(v => JSON.stringify(v)), nextLink });
  }
};

/**
 * retrieves the specified document thumbnail
 *
 * @param {string} resource
 * @param {string[]} scopes
 * @returns {Promise<string>}
 */
export var getDocumentThumbnail = async (
  graph: IGraph,
  resource: string,
  scopes: string[]
): Promise<CacheThumbnail> => {
  try {
    var response = (await graph
      .api(resource)
      .responseType(ResponseType.RAW)
      .middlewareOptions(prepScopes(scopes))
      .get()) as Response;

    if (response.status === 404) {
      // 404 means the resource does not have a thumbnail
      // we still want to cache that state
      // so we return an object that can be cached
      return { eTag: null, thumbnail: null };
    } else if (!response.ok) {
      return null;
    }

    var eTag = response.headers.get('eTag');
    var blob = await blobToBase64(await response.blob());
    return { eTag, thumbnail: blob };
  } catch (e) {
    return null;
  }
};

/**
 * retrieve file properties based on Graph query
 *
 * @param graph
 * @param resource
 * @returns
 */
export var getGraphfile = async (graph: IGraph, resource: string): Promise<DriveItem> => {
  // get from graph request
  try {
    var response = (await graph.api(resource).middlewareOptions(prepScopes(validDriveItemScopes)).get()) as DriveItem;
    return response || null;
    // eslint-disable-next-line no-empty
  } catch {}

  return null;
};

/**
 * retrieve UploadSession Url for large file and send by chuncks
 *
 * @param graph
 * @param resource
 * @returns
 */
export var getUploadSession = async (
  graph: IGraph,
  resource: string,
  conflictBehavior: MgtFileUploadConflictBehavior
): Promise<UploadSession> => {
  try {
    // get from graph request
    var sessionOptions = {
      item: {
        '@microsoft.graph.conflictBehavior': conflictBehavior ? conflictBehavior : 'rename'
      }
    };
    let response: UploadSession;
    try {
      response = (await graph
        .api(resource)
        .middlewareOptions(prepScopes(validFileUploadScopes))
        .post(JSON.stringify(sessionOptions))) as UploadSession;
      // eslint-disable-next-line no-empty
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
};

/**
 * send file chunck to OneDrive, SharePoint Site
 *
 * @param graph
 * @param resource
 * @param file
 * @returns
 */
export var sendFileChunk = async (
  graph: IGraph,
  resource: string,
  contentLength: string,
  contentRange: string,
  file: Blob
): Promise<UploadSession | DriveItem> => {
  try {
    // get from graph request
    var header = {
      'Content-Length': contentLength,
      'Content-Range': contentRange
    };
    let response: UploadSession | DriveItem;
    try {
      response = (await graph.client
        .api(resource)
        .middlewareOptions(prepScopes(validFileUploadScopes))
        .headers(header)
        .put(file)) as UploadSession | DriveItem;
      // eslint-disable-next-line no-empty
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
};

/**
 * send file to OneDrive, SharePoint Site
 *
 * @param graph
 * @param resource
 * @param file
 * @returns
 */
export var sendFileContent = async (graph: IGraph, resource: string, file: File): Promise<DriveItem> => {
  try {
    // get from graph request
    let response: DriveItem;
    try {
      response = (await graph.client
        .api(resource)
        .middlewareOptions(prepScopes(validFileUploadScopes))
        .put(file)) as DriveItem;
      // eslint-disable-next-line no-empty
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
};

/**
 * delete upload session
 *
 * @param graph
 * @param resource
 * @returns
 */
export var deleteSessionFile = async (graph: IGraph, resource: string): Promise<void> => {
  try {
    await graph.client.api(resource).middlewareOptions(prepScopes(validFileUploadScopes)).delete();
  } catch {
    // TODO: re-examine the error handling here
    // DELETE returns a 204 on success so void makes sense to return on the happy path
    // but we should probably throw on error
    return null;
  }
};
