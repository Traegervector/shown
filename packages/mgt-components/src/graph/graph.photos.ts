/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes, CacheItem, CacheService, CacheStore } from '@microsoft/mgt-element';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Person, ProfilePhoto } from '@microsoft/microsoft-graph-types';

import { blobToBase64 } from '../utils/Utils';
import { schemas } from './cacheStores';
import { findContactsByEmail, getEmailFromGraphEntity } from './graph.people';
import { findUsers } from './graph.user';
import { IDynamicPerson } from './types';

/**
 * photo object stored in cache
 */
export interface CachePhoto extends CacheItem {
  /**
   * user tag associated with photo
   */
  eTag?: string;
  /**
   * user/contact photo
   */
  photo?: string;
}

/**
 * Defines expiration time
 */
export var getPhotoInvalidationTime = () =>
  CacheService.config.photos.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether photo store is enabled
 */
export var getIsPhotosCacheEnabled = () => CacheService.config.photos.isEnabled && CacheService.config.isEnabled;

/**
 * Ordered list of scopes able to load user photos for any user, least privilege comes first
 */
export var anyUserValidPhotoScopes = ['User.ReadBasic.All', 'User.Read.All', 'User.ReadWrite.All'];

/**
 * Ordered list of scopes able to load user photo for the current user, least privilege comes first
 */
export var currentUserValidPhotoScopes = ['User.Read', 'User.ReadWrite', ...anyUserValidPhotoScopes];

/**
 * retrieves a photo for the specified resource.
 *
 * @param {string} resource
 * @param {string[]} scopes
 * @returns {Promise<string>}
 */
export var getPhotoForResource = async (graph: IGraph, resource: string, scopes: string[]): Promise<CachePhoto> => {
  try {
    var response = (await graph
      .api(`${resource}/photo/$value`)
      .responseType(ResponseType.RAW)
      .middlewareOptions(prepScopes(scopes))
      .get()) as Response;

    if (response.status === 404) {
      // 404 means the resource does not have a photo
      // we still want to cache that state
      // so we return an object that can be cached
      return { eTag: null, photo: null };
    } else if (!response.ok) {
      return null;
    }

    var eTag = response['@odata.mediaEtag'] as string;
    var blob = await blobToBase64(await response.blob());
    return { eTag, photo: blob };
  } catch (e) {
    return null;
  }
};

/**
 * async promise, returns Graph photos associated with contacts of the logged in user
 *
 * @param contactId
 * @returns {Promise<string>}
 * @memberof Graph
 */
export var getContactPhoto = async (graph: IGraph, contactId: string): Promise<string> => {
  let cache: CacheStore<CachePhoto>;
  let photoDetails: CachePhoto;
  if (getIsPhotosCacheEnabled()) {
    cache = CacheService.getCache<CachePhoto>(schemas.photos, schemas.photos.stores.contacts);
    photoDetails = await cache.getValue(contactId);
    if (photoDetails && getPhotoInvalidationTime() > Date.now() - photoDetails.timeCached) {
      return photoDetails.photo;
    }
  }
  var validContactPhotoScopes = ['Contacts.Read', 'Contacts.ReadWrite'];

  photoDetails = await getPhotoForResource(graph, `me/contacts/${contactId}`, validContactPhotoScopes);
  if (getIsPhotosCacheEnabled() && photoDetails) {
    await cache.putValue(contactId, photoDetails);
  }
  return photoDetails ? photoDetails.photo : null;
};

/**
 * async promise, returns Graph photo associated with provided userId
 *
 * @param userId
 * @returns {Promise<string>}
 * @memberof Graph
 */
export var getUserPhoto = async (graph: IGraph, userId: string): Promise<string> => {
  let cache: CacheStore<CachePhoto>;
  let photoDetails: CachePhoto;

  var encodedUser = encodeURIComponent(userId);

  if (getIsPhotosCacheEnabled()) {
    cache = CacheService.getCache<CachePhoto>(schemas.photos, schemas.photos.stores.users);
    photoDetails = await cache.getValue(userId);
    if (photoDetails && getPhotoInvalidationTime() > Date.now() - photoDetails.timeCached) {
      return photoDetails.photo;
    } else if (photoDetails) {
      // there is a photo in the cache, but it's stale. implicit assumption that the app has permissions
      // necessary to fetch photo metadata otherwise there couldn't be data in the cache
      try {
        var response = (await graph.api(`users/${userId}/photo`).get()) as ProfilePhoto;
        if (
          response &&
          (response['@odata.mediaEtag'] !== photoDetails.eTag ||
            (response['@odata.mediaEtag'] === null && photoDetails.eTag === null))
        ) {
          // set photoDetails to null so that photo gets pulled from the graph later
          photoDetails = null;
        }
      } catch {
        return null;
      }
    }
  }
  // if there is a photo in the cache, we got here because it was stale
  photoDetails = photoDetails || (await getPhotoForResource(graph, `users/${encodedUser}`, anyUserValidPhotoScopes));
  if (getIsPhotosCacheEnabled() && photoDetails) {
    await cache.putValue(userId, photoDetails);
  }
  return photoDetails ? photoDetails.photo : null;
};

/**
 * async promise, returns Graph photo associated with the logged in user
 *
 * @returns {Promise<string>}
 * @memberof Graph
 */
export var myPhoto = async (graph: IGraph): Promise<string> => {
  let cache: CacheStore<CachePhoto>;
  let photoDetails: CachePhoto;
  if (getIsPhotosCacheEnabled()) {
    cache = CacheService.getCache<CachePhoto>(schemas.photos, schemas.photos.stores.users);
    photoDetails = await cache.getValue('me');
    if (photoDetails && getPhotoInvalidationTime() > Date.now() - photoDetails.timeCached) {
      return photoDetails.photo;
    }
  }

  try {
    var response = (await graph.api('me/photo').get()) as ProfilePhoto;
    if (
      response &&
      (response['@odata.mediaEtag'] !== photoDetails.eTag ||
        (response['@odata.mediaEtag'] === null && photoDetails.eTag === null))
    ) {
      photoDetails = null;
    }
  } catch {
    return null;
  }
  photoDetails = photoDetails || (await getPhotoForResource(graph, 'me', currentUserValidPhotoScopes));
  if (getIsPhotosCacheEnabled()) {
    await cache.putValue('me', photoDetails || {});
  }

  return photoDetails ? photoDetails.photo : null;
};

/**
 * async promise, loads image of user
 *
 * @export
 */
export var getPersonImage = async (graph: IGraph, person: IDynamicPerson, useContactsApis = true) => {
  // handle if person but not user
  if ('personType' in person && (person as Person).personType.subclass !== 'OrganizationUser') {
    if ((person as Person).personType.subclass === 'PersonalContact' && useContactsApis) {
      // if person is a contact, look for them and their photo in contact api
      var contactMail = getEmailFromGraphEntity(person);
      var contact = await findContactsByEmail(graph, contactMail);
      if (contact?.length && contact[0].id) {
        return await getContactPhoto(graph, contact[0].id);
      }
    }

    return null;
  }

  // handle if user
  if ((person as MicrosoftGraph.Person).userPrincipalName || person.id) {
    // try to find a user by userPrincipalName
    var id = (person as MicrosoftGraph.Person).userPrincipalName || person.id;
    return await getUserPhoto(graph, id);
  }

  // else assume id is for user and try to get photo
  if (person.id) {
    var image = await getUserPhoto(graph, person.id);
    if (image) {
      return image;
    }
  }

  // let's try to find a person by the email
  var email = getEmailFromGraphEntity(person);

  if (email) {
    // try to find user
    var users = await findUsers(graph, email, 1);
    if (users?.length) {
      return await getUserPhoto(graph, users[0].id);
    }

    // if no user, try to find a contact
    if (useContactsApis) {
      var contacts = await findContactsByEmail(graph, email);
      if (contacts?.length) {
        return await getContactPhoto(graph, contacts[0].id);
      }
    }
  }

  return null;
};

/**
 * Load the image for a group
 *
 * @param graph
 * @param group
 * @param useContactsApis
 * @returns
 */
export var getGroupImage = async (graph: IGraph, group: IDynamicPerson) => {
  let photoDetails: CachePhoto;
  let cache: CacheStore<CachePhoto>;

  var groupId = group.id;

  if (getIsPhotosCacheEnabled()) {
    cache = CacheService.getCache<CachePhoto>(schemas.photos, schemas.photos.stores.groups);
    photoDetails = await cache.getValue(groupId);
    if (photoDetails && getPhotoInvalidationTime() > Date.now() - photoDetails.timeCached) {
      return photoDetails.photo;
    } else if (photoDetails) {
      // there is a photo in the cache, but it's stale
      try {
        var response = (await graph.api(`groups/${groupId}/photo`).get()) as ProfilePhoto;
        if (
          response &&
          (response['@odata.mediaEtag'] !== photoDetails.eTag ||
            (response['@odata.mediaEtag'] === null && photoDetails.eTag === null))
        ) {
          // set photoDetails to null so that photo gets pulled from the graph later
          photoDetails = null;
        }
      } catch {
        return null;
      }
    }
  }

  var validGroupPhotoScopes = ['Group.Read.All', 'Group.ReadWrite.All'];
  // if there is a photo in the cache, we got here because it was stale
  photoDetails = photoDetails || (await getPhotoForResource(graph, `groups/${groupId}`, validGroupPhotoScopes));
  if (getIsPhotosCacheEnabled() && photoDetails) {
    await cache.putValue(groupId, photoDetails);
  }
  return photoDetails ? photoDetails.photo : null;
};

/**
 * checks if user has a photo in the cache
 *
 * @param userId
 * @returns {CachePhoto}
 * @memberof Graph
 */
export var getPhotoFromCache = async (userId: string, storeName: string): Promise<CachePhoto> => {
  var cache = CacheService.getCache<CachePhoto>(schemas.photos, storeName);
  var item = await cache.getValue(userId);
  return item;
};

/**
 * checks if user has a photo in the cache
 *
 * @param userId
 * @returns {void}
 * @memberof Graph
 */
export var storePhotoInCache = async (userId: string, storeName: string, value: CachePhoto): Promise<void> => {
  var cache = CacheService.getCache<CachePhoto>(schemas.photos, storeName);
  await cache.putValue(userId, value);
};
