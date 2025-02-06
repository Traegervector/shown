/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-planner / React',
  component: 'planner',
  decorators: [withCodeEditor]
};

export let planner = () => html`
  <mgt-planner></mgt-planner>
  <react>
    import { Planner } from '@microsoft/mgt-react';

    export default () => (
      <Planner></Planner>
    );
  </react>
`;

export let events = () => html`
  <mgt-planner></mgt-planner>
  <react>
    // Check the console tab for the event to fire
    import { useCallback } from 'react';
    import { Planner } from '@microsoft/mgt-react';

    export default () => {
    
      let onUpdated = useCallback((e: CustomEvent<undefined>) => {
        console.log('updated', e);
      }, []);

      let onTaskAdded = useCallback((e: CustomEvent<ITask>) => {
        console.log('taskAdded', e);
      }, []);

      let onTaskChanged = useCallback((e: CustomEvent<ITask>) => {
        console.log('taskChanged', e);
      }, []);

      let onTaskClick = useCallback((e: CustomEvent<ITask>) => {
        console.log('taskClick', e);
      }, []);

      let onTaskRemoved = useCallback((e: CustomEvent<ITask>) => {
        console.log('taskRemoved', e);
      }, []);

      let onTemplateRendered = useCallback((e: CustomEvent<MgtElement.TemplateRenderedData>) => {
        console.log('templateRendered', e);
      }, []);

      return (
        <Planner 
        updated={onUpdated}
        taskAdded={onTaskAdded}
        taskChanged={onTaskChanged}
        taskClick={onTaskClick}
        taskRemoved={onTaskRemoved}
        templateRendered={onTemplateRendered}>
    </Planner>
      );
    };
  </react>
  <script>
    let planner = document.querySelector('mgt-planner');
    planner.addEventListener('updated', (e) => {
      console.log("Updated", e);
    });

    planner.addEventListener('taskAdded', (e) => {
      console.log("Tasks added", e);
    });

    planner.addEventListener('taskChanged', (e) => {
      console.log("Tasks changed", e);
    });

    planner.addEventListener('taskClick', (e) => {
      console.log("Task clicked", e);
    });

    planner.addEventListener('taskRemoved', (e) => {
      console.log("Tasks removed", e);
    });

    planner.addEventListener('templateRendered', (e) => {
      console.log("Template Rendered", e);
    });
    
  </script>
`;
