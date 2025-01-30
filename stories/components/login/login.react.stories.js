/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-login / React',
  component: 'login',
  decorators: [withCodeEditor]
};

export let Login = () => html`
  <mgt-login></mgt-login>
  <react>
    import { Login } from '@microsoft/mgt-react';

    export default () => (
      <Login></Login>
    );
  </react>
`;

export let LoginView = () => html`
  <mgt-login login-view="compact"></mgt-login>
  <react>
    import { Login } from '@microsoft/mgt-react';

    export default () => (
      <Login loginView='compact'></Login>
    );
  </react>
`;

export let ShowPresenceLogin = () => html`
  <mgt-login show-presence login-view="full"></mgt-login>
  <react>
    import { Login } from '@microsoft/mgt-react';

    export default () => (
      <Login showPresence={true} loginView='full'></Login>
    );
  </react>
`;

export let Events = () => html`
  <mgt-login></mgt-login>
  <react>
    // Check the console tab for the event to fire
    import { useCallback } from 'react';
    import { Login } from '@microsoft/mgt-react';

    export default () => {
      let onLoginInitiated = useCallback((e: CustomEvent<undefined>) => {
        console.log("Login Initiated");
      }, []);

      let onLoginCompleted = useCallback((e: CustomEvent<undefined>) => {
        console.log("Login Completed");
      }, []);

      let onLogoutInitiated = useCallback((e: CustomEvent<undefined>) => {
        console.log("Logout Initiated");
      }, []);

      let onLogoutCompleted = useCallback((e: CustomEvent<undefined>) => {
        console.log("Logout Completed");
      }, []);

      let onUpdated = useCallback((e: CustomEvent<undefined>) => {
        console.log('updated', e);
      }, []);

      return (
        <Login
        loginInitiated={onLoginInitiated}
        loginCompleted={onLoginCompleted}
        logoutInitiated={onLogoutInitiated}
        logoutCompleted={onLogoutCompleted}
        updated={onUpdated}>
    </Login>
      );
    };
  </react>
  <script>
    let login = document.querySelector('mgt-login');
    login.addEventListener('loginInitiated', (e) => {
      console.log("Login Initiated");
    })
    login.addEventListener('loginCompleted', (e) => {
      console.log("Login Completed");
    })
    login.addEventListener('logoutInitiated', (e) => {
      console.log("Logout Initiated");
    })
    login.addEventListener('logoutCompleted', (e) => {
      console.log("Logout Completed");
    })
    login.addEventListener('updated', (e) => {
      console.log("Updated");
    })
  </script>
`;
