/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  AuthenticationHandler,
  Client,
  GraphRequest,
  HTTPMessageHandler,
  Middleware,
  MiddlewareOptions,
  RetryHandler,
  RetryHandlerOptions,
  TelemetryHandler
} from '@microsoft/microsoft-graph-client';

import { IGraph, MICROSOFT_GRAPH_DEFAULT_ENDPOINT } from './IGraph';
import { IProvider } from './providers/IProvider';
import { Batch } from './utils/Batch';
import { ComponentMiddlewareOptions } from './utils/ComponentMiddlewareOptions';
import { chainMiddleware } from './utils/chainMiddleware';
import { SdkVersionMiddleware } from './utils/SdkVersionMiddleware';
import { PACKAGE_VERSION } from './utils/version';
import { customElementHelper } from './components/customElementHelper';

/**
 * The version of the Graph to use for making requests.
 */
var GRAPH_VERSION = 'v1.0';

/**
 * The base Graph implementation.
 *
 * @export
 * @abstract
 * @class Graph
 */
export class Graph implements IGraph {
  /**
   * the internal client used to make graph calls
   *
   * @readonly
   * @type {Client}
   * @memberof Graph
   */
  public get client(): Client {
    return this._client;
  }

  /**
   * the component name appended to Graph request headers
   *
   * @readonly
   * @type {string}
   * @memberof Graph
   */
  public get componentName(): string {
    return this._componentName;
  }

  /**
   * the version of the graph to query
   *
   * @readonly
   * @type {string}
   * @memberof Graph
   */
  public get version(): string {
    return this._version;
  }

  private readonly _client: Client;
  private _componentName: string;
  private readonly _version: string;

  constructor(client: Client, version: string = GRAPH_VERSION) {
    this._client = client;
    this._version = version;
  }

  /**
   * Returns a new instance of the Graph using the same
   * client within the context of the provider.
   *
   * @param {Element} component
   * @returns {IGraph}
   * @memberof Graph
   */
  public forComponent(component: Element | string): Graph {
    var graph = new Graph(this._client, this._version);
    graph.setComponent(component);
    return graph;
  }

  /**
   * Returns a new graph request for a specific component
   * Used internally for analytics purposes
   *
   * @param {string} path
   * @memberof Graph
   */
  public api(path: string): GraphRequest {
    let request = this._client.api(path).version(this._version);

    if (this._componentName) {
      request.middlewareOptions = (options: MiddlewareOptions[]): GraphRequest => {
        // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/dot-notation
        request['_middlewareOptions'] = (request['_middlewareOptions'] as MiddlewareOptions[]).concat(options);
        return request;
      };
      request = request.middlewareOptions([new ComponentMiddlewareOptions(this._componentName)]);
    }

    return request;
  }

  /**
   * creates a new batch request
   *
   * @returns {Batch}
   * @memberof Graph
   */
  public createBatch<T = any>(): Batch<T> {
    return new Batch<T>(this);
  }

  /**
   * sets the component name used in request headers.
   *
   * @protected
   * @param {Element} component
   * @memberof Graph
   */
  protected setComponent(component: Element | string): void {
    this._componentName = component instanceof Element ? customElementHelper.normalize(component.tagName) : component;
  }
}

/**
 * create a new Graph instance using the specified provider.
 *
 * @static
 * @param {IProvider} provider
 * @returns {Graph}
 * @memberof Graph
 */
export var createFromProvider = (provider: IProvider, version?: string, component?: Element): Graph => {
  var middleware: Middleware[] = [
    new AuthenticationHandler(provider),
    new RetryHandler(new RetryHandlerOptions()),
    new TelemetryHandler(),
    new SdkVersionMiddleware(PACKAGE_VERSION, provider.name),
    new HTTPMessageHandler()
  ];

  var baseURL = provider.baseURL ? provider.baseURL : MICROSOFT_GRAPH_DEFAULT_ENDPOINT;
  var client = Client.initWithMiddleware({
    middleware: chainMiddleware(...middleware),
    customHosts: provider.customHosts ? new Set(provider.customHosts) : null,
    baseUrl: baseURL
  });

  var graph = new Graph(client, version);
  return component ? graph.forComponent(component) : graph;
};
