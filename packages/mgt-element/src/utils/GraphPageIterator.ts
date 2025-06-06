/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { GraphRequest } from '@microsoft/microsoft-graph-client';
import { IGraph } from '../IGraph';
import { CollectionResponse } from '../CollectionResponse';

/**
 * A helper class to assist in getting multiple pages from a resource
 *
 * @export
 * @class GraphPageIterator
 * @template T
 */
export class GraphPageIterator<T> {
  /**
   * Gets all the items already fetched for this request
   *
   * @readonly
   * @type {T[]}
   * @memberof GraphPageIterator
   */
  public get value(): T[] {
    return this._value;
  }

  /**
   * Gets wheather this request has more pages
   *
   * @readonly
   * @type {boolean}
   * @memberof GraphPageIterator
   */
  public get hasNext(): boolean {
    return Boolean(this._nextLink);
  }

  /**
   * Creates a new GraphPageIterator
   *
   * @static
   * @template T - the type of entities expected from this request
   * @param {IGraph} graph - the graph instance to use for making requests
   * @param {GraphRequest} request - the initial request
   * @param {string} [version] - optional version to use for the requests - by default uses the default version
   * from the graph parameter
   * @returns a GraphPageIterator
   * @memberof GraphPageIterator
   */
  public static async create<T>(graph: IGraph, request: GraphRequest, version?: string): Promise<GraphPageIterator<T>> {
    let response = (await request.get()) as CollectionResponse<T>;
    if (response?.value) {
      let iterator = new GraphPageIterator<T>();
      iterator._graph = graph;
      iterator._value = response.value;
      iterator._nextLink = response['@odata.nextLink'] as string;
      iterator._version = version || graph.version;
      return iterator;
    }

    return null;
  }

  /**
   * Creates a new GraphPageIterator from existing value
   *
   * @static
   * @template T - the type of entities expected from this request
   * @param {IGraph} graph - the graph instance to use for making requests
   * @param value - the existing value
   * @param nextLink - optional nextLink to use to get the next page
   * from the graph parameter
   * @returns a GraphPageIterator
   * @memberof GraphPageIterator
   */
  public static createFromValue<T>(graph: IGraph, value: T[], nextLink: string = null): GraphPageIterator<T> {
    let iterator = new GraphPageIterator<T>();

    // create iterator from values
    iterator._graph = graph;
    iterator._value = value;
    iterator._nextLink = nextLink;
    iterator._version = graph.version;

    return iterator;
  }

  private _graph: IGraph;
  private _nextLink: string;
  /**
   * Gets the next link for this request
   *
   * @readonly
   * @type {string}
   * @memberof GraphPageIterator
   */
  public get nextLink(): string {
    return this._nextLink || '';
  }
  private _version: string;
  private _value: T[];

  /**
   * Gets the next page for this request
   *
   * @returns {Promise<T[]>}
   * @memberof GraphPageIterator
   */
  public async next(): Promise<T[]> {
    if (this._nextLink) {
      let nextResource = this._nextLink.split(this._version)[1];
      let response = (await this._graph.api(nextResource).version(this._version).get()) as CollectionResponse<T>;
      if (response?.value?.length) {
        this._value = this._value.concat(response.value);
        this._nextLink = response['@odata.nextLink'] as string;
        return response.value;
      }
    }
    return null;
  }
}
