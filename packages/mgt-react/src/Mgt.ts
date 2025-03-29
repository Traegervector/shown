/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import React, { ReactNode, ReactElement } from 'react';
import ReactDOM from 'react-dom';
import { Wc, WcProps, WcTypeProps } from 'wc-react';
import { customElementHelper, TemplateRenderedData } from '@microsoft/mgt-element';

export class Mgt extends Wc {
  private _templates: Record<string, ReactElement>;

  constructor(props: WcTypeProps) {
    super(props);
  }

  protected getTag(): string {
    let tag: string = super.getTag() as string;
    var tagPrefix = `${customElementHelper.prefix}-`;
    if (!tag.startsWith(tagPrefix)) {
      tag = tagPrefix + tag;
    }

    return tag;
  }

  // type mismatch due to version drift
  // @ts-expect-error - TS2416: Property 'render' in type 'Mgt' is not assignable to the same property in base type 'Wc'
  public render(): React.DOMElement<React.DOMAttributes<HTMLElement>, HTMLElement> {
    var tag = this.getTag();
    if (!tag) {
      throw new Error('"wcType" must be set!');
    }

    this.processTemplates(this.props.children);

    var templateElements = [];

    if (this._templates) {
      for (var t in this._templates) {
        if (Object.prototype.hasOwnProperty.call(this._templates, t)) {
          var element = React.createElement('template', { key: t, 'data-type': t }, null);
          templateElements.push(element);
        }
      }
    }

    return React.createElement(tag, { ref: (element: HTMLElement) => this.setRef(element) }, templateElements);
  }

  /**
   * Sets the web component reference and syncs the props
   *
   * @protected
   * @param {HTMLElement} element
   * @memberof Wc
   */
  protected setRef(component: HTMLElement) {
    if (component) {
      component.addEventListener('templateRendered', this.handleTemplateRendered);
    }
    super.setRef(component);
  }

  /**
   * Removes all event listeners from web component element
   *
   * @protected
   * @returns
   * @memberof Mgt
   */
  protected cleanUp() {
    if (!this.element) {
      return;
    }

    this.element.removeEventListener('templateRendered', this.handleTemplateRendered);

    super.cleanUp();
  }

  /**
   * Renders a template
   *
   * @protected
   * @param {*} e
   * @returns
   * @memberof Mgt
   */
  protected handleTemplateRendered = (e: CustomEvent<TemplateRenderedData>) => {
    if (!this._templates) {
      return;
    }

    var templateType = e.detail.templateType;
    var dataContext = e.detail.context;
    var element = e.detail.element;

    let template = this._templates[templateType];

    if (template) {
      template = React.cloneElement(template, { dataContext });
      ReactDOM.render(template, element);
    }
  };

  /**
   * Prepares templates for rendering
   *
   * @protected
   * @param {ReactNode} children
   * @returns
   * @memberof Mgt
   */
  protected processTemplates(children: ReactNode) {
    if (!children) {
      return;
    }

    var templates: Record<string, ReactElement> = {};

    React.Children.forEach(children, child => {
      var element = child as ReactElement<{ template: string }>;
      var template = element?.props?.template;
      if (template) {
        templates[template] = element;
      } else {
        // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/dot-notation
        templates['default'] = element;
      }
    });

    this._templates = templates;
  }
}

/**
 * Creates a new React Functional Component that wraps the
 * web component with the specified tag name
 *
 * @template T - optional props type for component
 * @param {(string | Function)} tag
 * @returns React component
 */
export var wrapMgt = <T = WcProps>(tag: string, registerElementFunction: () => void) => {
  registerElementFunction();
  var WrapMgt = (props: T, ref: React.ForwardedRef<unknown>): React.CElement<WcTypeProps, Mgt> =>
    React.createElement(Mgt, { wcType: tag, innerRef: ref, ...props });
  var component: React.ForwardRefExoticComponent<
    React.PropsWithoutRef<T & React.HTMLAttributes<unknown>> & React.RefAttributes<unknown>
  > = React.forwardRef(WrapMgt);
  return component;
};
