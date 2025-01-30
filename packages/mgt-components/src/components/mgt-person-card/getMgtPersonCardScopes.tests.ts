/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MgtPersonCardConfig } from './MgtPersonCardConfig';
import { expect } from '@open-wc/testing';
import { getMgtPersonCardScopes } from './getMgtPersonCardScopes';

describe('getMgtPersonCardScopes() tests', () => {
  let originalConfigMessaging = MgtPersonCardConfig.isSendMessageVisible;
  let originalConfigContactApis = MgtPersonCardConfig.useContactApis;
  let originalConfigOrgSection = { ...MgtPersonCardConfig.sections.organization };
  let originalConfigSections = { ...MgtPersonCardConfig.sections };
  beforeEach(() => {
    MgtPersonCardConfig.sections = { ...originalConfigSections };
    MgtPersonCardConfig.sections.organization = { ...originalConfigOrgSection };
    MgtPersonCardConfig.useContactApis = originalConfigContactApis;
    MgtPersonCardConfig.isSendMessageVisible = originalConfigMessaging;
  });
  it('should have a minimal permission set', () => {
    let expectedScopes = [
      'User.Read.All',
      'People.Read.All',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read',
      'Chat.ReadWrite'
    ];
    expect(getMgtPersonCardScopes()).to.have.members(expectedScopes);
  });

  it('should have not have Sites.Read.All if files is configured off', () => {
    MgtPersonCardConfig.sections.files = false;

    let expectedScopes = [
      'User.Read.All',
      'People.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read',
      'Chat.ReadWrite'
    ];
    expect(getMgtPersonCardScopes()).to.have.members(expectedScopes);
  });

  it('should have not have Mail scopes if mail is configured off', () => {
    MgtPersonCardConfig.sections.mailMessages = false;

    let expectedScopes = ['User.Read.All', 'People.Read.All', 'Sites.Read.All', 'Contacts.Read', 'Chat.ReadWrite'];
    expect(getMgtPersonCardScopes()).to.have.members(expectedScopes);
  });

  it('should have People.Read but not People.Read.All if showWorksWith is false', () => {
    MgtPersonCardConfig.sections.organization.showWorksWith = false;
    let expectedScopes = [
      'User.Read.All',
      'People.Read',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read',
      'Chat.ReadWrite'
    ];
    expect(getMgtPersonCardScopes()).to.have.members(expectedScopes);
  });

  it('should have not have User.Read.All if profile and organization are false', () => {
    MgtPersonCardConfig.sections.organization = undefined;
    MgtPersonCardConfig.sections.profile = false;

    let expectedScopes = [
      'User.Read',
      'User.ReadBasic.All',
      'People.Read',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read',
      'Chat.ReadWrite'
    ];
    let actualScopes = getMgtPersonCardScopes();
    expect(actualScopes).to.have.members(expectedScopes);

    expect(actualScopes).to.not.include('User.Read.All');
  });

  it('should have not have Chat.ReadWrite if isSendMessageVisible is false', () => {
    MgtPersonCardConfig.isSendMessageVisible = false;

    let expectedScopes = [
      'User.Read.All',
      'People.Read.All',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read'
    ];
    let actualScopes = getMgtPersonCardScopes();
    expect(actualScopes).to.have.members(expectedScopes);

    expect(actualScopes).to.not.include('Chat.ReadWrite');
  });

  it('should have not have Chat.ReadWrite if useContactApis is false', () => {
    MgtPersonCardConfig.useContactApis = false;

    let expectedScopes = [
      'User.Read.All',
      'People.Read.All',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Chat.ReadWrite'
    ];
    let actualScopes = getMgtPersonCardScopes();
    expect(actualScopes).to.have.members(expectedScopes);

    expect(actualScopes).to.not.include('Contacts.Read');
  });
});
