// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <graphInitSnippet>
let graphClient = undefined;

function initializeGraphClient(msalClient, account, scopes)
{
  // Create an authentication provider
  const authProvider = new MSGraphAuthCodeMSALBrowserAuthProvider
  .AuthCodeMSALBrowserAuthenticationProvider(msalClient, {
    account: account,
    scopes: scopes,
    interactionType: msal.InteractionType.PopUp
  });

  // Initialize the Graph client
  graphClient = MicrosoftGraph.Client.initWithMiddleware({authProvider});
}
// </graphInitSnippet>

// <getUserSnippet>
async function getUser() {
    return graphClient
      .api('/me')
      // Only get the fields used by the app
      .select('id,displayName,mail,userPrincipalName,mailboxSettings')
      .get();
  }
  // </getUserSnippet>

  // <getEventsSnippet>
async function getEvents() {
  const user = JSON.parse(sessionStorage.getItem('graphUser'));

  // Convert user's Windows time zone ("Pacific Standard Time")
  // to IANA format ("America/Los_Angeles")
  // Moment needs IANA format
  let ianaTimeZone = getIanaFromWindows(user.mailboxSettings.timeZone);
  console.log(`Converted: ${ianaTimeZone}`);

  // Configure a calendar view for the current week
  // Get midnight on the start of the current week in the user's timezone,
  // but in UTC. For example, for Pacific Standard Time, the time value would be
  // 07:00:00Z
  let startOfWeek = moment.tz(ianaTimeZone).startOf('week').utc();
  // Set end of the view to 7 days after start of week
  let endOfWeek = moment(startOfWeek).add(7, 'day');

  try {
    // GET /me/calendarview?startDateTime=''&endDateTime=''
    // &$select=subject,organizer,start,end
    // &$orderby=start/dateTime
    // &$top=50
    let response = await graphClient
      .api('/me/calendarview')
      // Set the Prefer=outlook.timezone header so date/times are in
      // user's preferred time zone
      .header("Prefer", `outlook.timezone="${user.mailboxSettings.timeZone}"`)
      // Add the startDateTime and endDateTime query parameters
      .query({ startDateTime: startOfWeek.format(), endDateTime: endOfWeek.format() })
      // Select just the fields we are interested in
      .select('subject,organizer,start,end')
      // Sort the results by start, earliest first
      .orderby('start/dateTime')
      // Maximum 50 events in response
      .top(50)
      .get();

    updatePage(Views.calendar, response.value);
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting events',
      debug: error
    });
  }
}
// </getEventsSnippet>

// <showCalendarSnippet>
function showCalendar(events) {
  let div = document.createElement('div');

  div.appendChild(createElement('h1', 'mb-3', 'Calendar'));

  let newEventButton = createElement('button', 'btn btn-light btn-sm mb-3', 'New event');
  newEventButton.setAttribute('onclick', 'showNewEventForm();');
  div.appendChild(newEventButton);

  let table = createElement('table', 'table');
  div.appendChild(table);

  let thead = document.createElement('thead');
  table.appendChild(thead);

  let headerrow = document.createElement('tr');
  thead.appendChild(headerrow);

  let organizer = createElement('th', null, 'Organizer');
  organizer.setAttribute('scope', 'col');
  headerrow.appendChild(organizer);

  let subject = createElement('th', null, 'Subject');
  subject.setAttribute('scope', 'col');
  headerrow.appendChild(subject);

  let start = createElement('th', null, 'Start');
  start.setAttribute('scope', 'col');
  headerrow.appendChild(start);

  let end = createElement('th', null, 'End');
  end.setAttribute('scope', 'col');
  headerrow.appendChild(end);

  let tbody = document.createElement('tbody');
  table.appendChild(tbody);

  for (const event of events) {
    let eventrow = document.createElement('tr');
    eventrow.setAttribute('key', event.id);
    tbody.appendChild(eventrow);

    let organizercell = createElement('td', null, event.organizer.emailAddress.name);
    eventrow.appendChild(organizercell);

    let subjectcell = createElement('td', null, event.subject);
    eventrow.appendChild(subjectcell);

    // Use moment.utc() here because times are already in the user's
    // preferred timezone, and we don't want moment to try to change them to the
    // browser's timezone
    let startcell = createElement('td', null,
      moment.utc(event.start.dateTime).format('M/D/YY h:mm A'));
    eventrow.appendChild(startcell);

    let endcell = createElement('td', null,
      moment.utc(event.end.dateTime).format('M/D/YY h:mm A'));
    eventrow.appendChild(endcell);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(div);
}
// </showCalendarSnippet>

// <createEventSnippet>
async function createNewEvent() {
  const user = JSON.parse(sessionStorage.getItem('graphUser'));

  // Get the user's input
  const subject = document.getElementById('ev-subject').value;
  const attendees = document.getElementById('ev-attendees').value;
  const start = document.getElementById('ev-start').value;
  const end = document.getElementById('ev-end').value;
  const body = document.getElementById('ev-body').value;

  // Require at least subject, start, and end
  if (!subject || !start || !end) {
    updatePage(Views.error, {
      message: 'Please provide a subject, start, and end.'
    });
    return;
  }

  // Build the JSON payload of the event
  let newEvent = {
    subject: subject,
    start: {
      dateTime: start,
      timeZone: user.mailboxSettings.timeZone
    },
    end: {
      dateTime: end,
      timeZone: user.mailboxSettings.timeZone
    }
  };

  if (attendees)
  {
    const attendeeArray = attendees.split(';');
    newEvent.attendees = [];

    for (const attendee of attendeeArray) {
      if (attendee.length > 0) {
        newEvent.attendees.push({
          type: 'required',
          emailAddress: {
            address: attendee
          }
        });
      }
    }
  }

  if (body)
  {
    newEvent.body = {
      contentType: 'text',
      content: body
    };
  }

  try {
    // POST the JSON to the /me/events endpoint
    await graphClient
      .api('/me/events')
      .post(newEvent);

    // Return to the calendar view
    getEvents();
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error creating event',
      debug: error
    });
  }
}
// </createEventSnippet>