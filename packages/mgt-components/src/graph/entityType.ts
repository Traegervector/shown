/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IDynamicPerson, IContact, IGroup, IUser } from './types';

export let isGroup = (obj: IDynamicPerson): obj is IGroup => {
  return 'groupTypes' in obj;
};

export let isUser = (obj: IDynamicPerson): obj is IUser => {
  return 'personType' in obj || 'userType' in obj;
};

export let isContact = (obj: IDynamicPerson): obj is IContact => {
  return 'initials' in obj;
};
