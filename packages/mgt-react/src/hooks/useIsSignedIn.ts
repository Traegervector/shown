/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { isSignedIn, Providers } from '@microsoft/mgt-element';
import { useState, useEffect } from 'react';

/**
 * Hook to check if a user is signed in.
 *
 * @returns true if a user is signed on, otherwise false.
 */
export let useIsSignedIn = (): [boolean] => {
  let [signedIn, setIsSignedIn] = useState(false);
  useEffect(() => {
    let updateState = () => {
      setIsSignedIn(isSignedIn());
    };

    Providers.onProviderUpdated(updateState);
    updateState();

    return () => {
      Providers.removeProviderUpdatedListener(updateState);
    };
  }, [setIsSignedIn]);
  return [signedIn];
};
