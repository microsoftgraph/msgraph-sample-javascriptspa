// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Select DOM elements to work with
const authenticatedNav = document.getElementById('authenticated-nav');
const accountNav = document.getElementById('account-nav');
const mainContainer = document.getElementById('main-container');

const Views = { error: 1, home: 2, calendar: 3 };

function createElement(type, className, text) {
  var element = document.createElement(type);
  element.className = className;

  if (text) {
    var textNode = document.createTextNode(text);
    element.appendChild(textNode);
  }

  return element;
}

function showAuthenticatedNav(account, view) {
  authenticatedNav.innerHTML = '';

  if (account) {
    // Add Calendar link
    var calendarNav = createElement('li', 'nav-item');

    var calendarLink = createElement('button',
      `btn btn-link nav-link${view === Views.calendar ? ' active' : '' }`,
      'Calendar');
    calendarLink.setAttribute('onclick', 'getEvents();');
    calendarNav.appendChild(calendarLink);

    authenticatedNav.appendChild(calendarNav);
  }
}

function showAccountNav(account, view) {

  accountNav.innerHTML = '';

  if (account) {
    // Show the "signed-in" nav
    accountNav.className = 'nav-item dropdown';

    var dropdown = createElement('a', 'nav-link dropdown-toggle');
    dropdown.setAttribute('data-toggle', 'dropdown');
    dropdown.setAttribute('role', 'button');
    accountNav.appendChild(dropdown);

    var userIcon = createElement('i',
      'far fa-user-circle fa-lg rounded-circle align-self-center');
    userIcon.style.width = '32px';
    dropdown.appendChild(userIcon);

    var menu = createElement('div', 'dropdown-menu dropdown-menu-right');
    dropdown.appendChild(menu);

    var userName = createElement('h5', 'dropdown-item-text mb-0', account.name);
    menu.appendChild(userName);

    var userEmail = createElement('p', 'dropdown-item-text text-muted mb-0', account.userName);
    menu.appendChild(userEmail);

    var divider = createElement('div', 'dropdown-divider');
    menu.appendChild(divider);

    var signOutButton = createElement('button', 'dropdown-item', 'Sign out');
    signOutButton.setAttribute('onclick', 'signOut();');
    menu.appendChild(signOutButton);
  } else {
    // Show a "sign in" button
    accountNav.className = 'nav-item';

    var signInButton = createElement('button', 'btn btn-link nav-link', 'Sign in');
    signInButton.setAttribute('onclick', 'signIn();');
    accountNav.appendChild(signInButton);
  }
}

function showWelcomeMessage(account) {
  // Create jumbotron
  var jumbotron = createElement('div', 'jumbotron');

  var heading = createElement('h1', null, 'JavaScript SPA Graph Tutorial');
  jumbotron.appendChild(heading);

  var lead = createElement('p', 'lead',
    'This sample app shows how to use the Microsoft Graph API to access' +
    ' a user\'s data from JavaScript.');
  jumbotron.appendChild(lead);

  if (account) {
    // Welcome the user by name
    var welcomeMessage = createElement('h4', null, `Welcome ${account.name}!`);
    jumbotron.appendChild(welcomeMessage);

    var callToAction = createElement('p', null,
      'Use the navigation bar at the top of the page to get started.');
    jumbotron.appendChild(callToAction);
  } else {
    // Show a sign in button in the jumbotron
    var signInButton = createElement('button', 'btn btn-primary btn-large',
      'Click here to sign in');
    signInButton.setAttribute('onclick', 'signIn();')
    jumbotron.appendChild(signInButton);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(jumbotron);
}

function showError(error) {
  var alert = createElement('div', 'alert alert-danger');

  var message = createElement('p', 'mb-3', error.message);
  alert.appendChild(message);

  if (error.debug)
  {
    var pre = createElement('pre', 'alert-pre border bg-light p-2');
    alert.appendChild(pre);

    var code = createElement('code', 'text-break text-wrap',
      JSON.stringify(error.debug, null, 2));
    pre.appendChild(code);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(alert);
}

// <showCalendar>
function showCalendar(events) {
  var div = document.createElement('div');

  div.appendChild(createElement('h1', null, 'Calendar'));

  var table = createElement('table', 'table');
  div.appendChild(table);

  var thead = document.createElement('thead');
  table.appendChild(thead);

  var headerrow = document.createElement('tr');
  thead.appendChild(headerrow);

  var organizer = createElement('th', null, 'Organizer');
  organizer.setAttribute('scope', 'col');
  headerrow.appendChild(organizer);

  var subject = createElement('th', null, 'Subject');
  subject.setAttribute('scope', 'col');
  headerrow.appendChild(subject);

  var start = createElement('th', null, 'Start');
  start.setAttribute('scope', 'col');
  headerrow.appendChild(start);

  var end = createElement('th', null, 'End');
  end.setAttribute('scope', 'col');
  headerrow.appendChild(end);

  var tbody = document.createElement('tbody');
  table.appendChild(tbody);

  for (const event of events.value) {
    var eventrow = document.createElement('tr');
    eventrow.setAttribute('key', event.id);
    tbody.appendChild(eventrow);

    var organizercell = createElement('td', null, event.organizer.emailAddress.name);
    eventrow.appendChild(organizercell);

    var subjectcell = createElement('td', null, event.subject);
    eventrow.appendChild(subjectcell);

    var startcell = createElement('td', null,
      moment.utc(event.start.dateTime).local().format('M/D/YY h:mm A'));
    eventrow.appendChild(startcell);

    var endcell = createElement('td', null,
      moment.utc(event.end.dateTime).local().format('M/D/YY h:mm A'));
    eventrow.appendChild(endcell);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(div);
}
// </showCalendar>

// <updatePage>
function updatePage(account, view, data) {
  if (!view || !account) {
    view = Views.home;
  }

  showAccountNav(account);
  showAuthenticatedNav(account, view);

  switch (view) {
    case Views.error:
      showError(data);
      break;
    case Views.home:
      showWelcomeMessage(account);
      break;
    case Views.calendar:
      showCalendar(data);
      break;
  }
}
// </updatePage>

updatePage(null, Views.home);
