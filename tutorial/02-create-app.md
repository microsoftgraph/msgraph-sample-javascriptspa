<!-- markdownlint-disable MD002 MD041 -->

Start by creating an empty directory for the project. This can be on an HTTP server, or a directory on your development machine. If it is on your development machine, you'll need to copy it to a server for testing, or run an HTTP server on your development machine. If you don't have either of those, the next section provides instructions.

## Start a local web server (optional)

> [!NOTE]
> The steps in this section require [Node.js](https://nodejs.org).

In this section you will use [http-server](https://www.npmjs.com/package/http-server) to run a simple HTTP server from the command line.

1. Open your command-line interface (CLI) in the directory you created for the project.
1. Run the following command to start a web server in that directory.

    ```Shell
    npx http-server -c-1
    ```

1. Open your browser and browse to `http://localhost:8080`.

You should see an **Index of /** page. This confirms that the HTTP server is running.

![A screenshot of the index page served by http-server.](images/run-web-server.png)

## Design the app

In this section you'll create the basic UI layout for the application.

1. Create a new file in the root of the project named `index.html` and add the following code.

    ```html
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0, shrink-to-fit=no">
      <title>JavaScript SPA Graph Tutorial</title>

      <link rel="shortcut icon" href="g-raph.png">
      <link rel="stylesheet"
            href="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css"
            integrity="sha384-Vkoo8x4CGsO3+Hhxv8T/Q5PaXtkKtu6ug5TOeNV6gBiFeWPGFN9MuhOf23Q9Ifjh"
            crossorigin="anonymous">
      <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.12.1/css/all.css"
            crossorigin="anonymous">
      <link href="style.css" rel="stylesheet" type="text/css" />
    </head>

    <body>
      <nav class="navbar navbar-expand-md navbar-dark fixed-top bg-dark">
        <div class="container">
          <a href="/" class="navbar-brand">Javascript SPA Graph Tutorial</a>
          <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarCollapse"
            aria-controls="navbarCollapse" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
          </button>
          <div class="collapse navbar-collapse" id="navbarCollapse">
            <ul id="authenticated-nav" class="navbar-nav mr-auto"></ul>
            <ul class="navbar-nav justify-content-end">
              <li class="nav-item">
                <a class="nav-link" href="https://developer.microsoft.com/graph/docs/concepts/overview" target="_blank">
                  <i class="fas fa-external-link-alt mr-1"></i>Docs
                </a>
              </li>
              <li id="account-nav" class="nav-item"></li>
            </ul>
          </div>
        </div>
      </nav>

      <main id="main-container" role="main" class="container">

      </main>

      <!-- Bootstrap/jQuery -->
      <script src="https://code.jquery.com/jquery-3.4.1.slim.min.js"
              integrity="sha384-J6qa4849blE2+poT4WnyKhv5vZF5SrPo0iEjwBvKU7imGFAV0wwj1yYfoRSJoZ+n"
              crossorigin="anonymous"></script>
      <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js"
              integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo"
              crossorigin="anonymous"></script>
      <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/js/bootstrap.min.js"
              integrity="sha384-wfSDF2E50Y2D1uUdj0O3uMBJnjuUD4Ih7YwaYd1iqfktj0Uod8GCExl3Og8ifwB6"
              crossorigin="anonymous"></script>

      <!-- MSAL -->
      <script src="//cdn.jsdelivr.net/npm/bluebird@3.7.2/js/browser/bluebird.min.js"></script>
      <script src="https://alcdn.msftauth.net/lib/1.2.1/js/msal.js"
              integrity="sha384-9TV1245fz+BaI+VvCjMYL0YDMElLBwNS84v3mY57pXNOt6xcUYch2QLImaTahcOP"
              crossorigin="anonymous"></script>

      <!-- Graph SDK -->
      <script src="https://cdn.jsdelivr.net/npm/@microsoft/microsoft-graph-client/lib/graph-js-sdk.js"></script>

      <script src="config.js"></script>
      <script src="ui.js"></script>
      <script src="auth.js"></script>
      <script src="graph.js"></script>
    </body>
    </html>
    ```

    This defines the basic layout of the app, including a navigation bar. It also adds the following:

    - [Bootstrap](https://getbootstrap.com/) and its supporting JavaScript
    - [FontAwesome](https://fontawesome.com/)
    - [Microsoft Authentication Library for JavaScript (MSAL.js)](https://github.com/AzureAD/microsoft-authentication-library-for-js)
    - [Microsoft Graph JavaScript Client Library](https://github.com/microsoftgraph/msgraph-sdk-javascript)

    > [!TIP]
    > The page includes a favicon, (`<link rel="shortcut icon" href="g-raph.png">`). You can remove this line, or you can download the `g-raph.png` file from [GitHub](https://github.com/microsoftgraph/g-raph).

1. Create a new file named `style.css` and add the following code.

    ```css
    body {
      padding-top: 70px;
    }
    ```

1. Create a new file named `auth.js` and add the following code.

    ```javascript
    function signIn() {
      // TEMPORARY
      updatePage({name: 'Megan Bowen', userName: 'meganb@contoso.com'});
    }

    function signOut() {
      // TEMPORARY
      updatePage();
    }
    ```

1. Create a new file named `ui.js` and add the following code.

    ```javascript
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

    function showAccountNav(account) {
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

    function updatePage(account, view, error) {
      showAccountNav(account);

      if (!view || !account) {
        view = Views.home;
      }

      switch (view) {
        case Views.error:
          showError(error);
          break;
        case Views.home:
          showWelcomeMessage(account);
          break;
        case Views.calendar:
          break;
      }
    }

    updatePage(null, Views.home);
    ```

Save all of your changes and refresh the page. Now, the app should look very different.

![A screenshot of the redesigned home page](images/app-layout.png)
