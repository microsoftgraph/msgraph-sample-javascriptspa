// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <graphInit>
// Create an options object with the same scopes from the login
const options =
  new MicrosoftGraph.MSALAuthenticationProviderOptions([
    'user.read',
    'calendars.read'
  ]);
// Create an authentication provider for the implicit flow
const authProvider =
  new MicrosoftGraph.ImplicitMSALAuthenticationProvider(msalClient, options);
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({authProvider});
// </graphInit>

// <getEvents>
async function getEvents() {
  try {
    let events = await graphClient
        .api('/me/events')
        .select('subject,organizer,start,end')
        .orderby('createdDateTime DESC')
        .get();

    updatePage(msalClient.getAccount(), Views.calendar, events);
  } catch (error) {
    updatePage(msalClient.getAccount(), Views.error, {
      message: 'Error getting events',
      debug: error
    });
  }
}
// </getEvents>
