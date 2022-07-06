// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

const settings: AppSettings = {
    'clientId': 'xx',
    'clientSecret': 'xx',
    'tenantId': 'xx',
    'authTenant': 'xx',
    'graphUserScopes': [
      'Calendars.Read',
      'Calendars.ReadWrite',
      'User.Read.All'
    ]
  };
  
  export interface AppSettings {
    clientId: string;
    clientSecret: string;
    tenantId: string;
    authTenant: string;
    graphUserScopes: string[];
  }
  
  export default settings;