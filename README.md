---
page_type: sample
description: This sample demonstrates how to use the Microsoft Graph JavaScript SDK to access data in Office 365 from JavaScript browser apps.
products:
- ms-graph
- microsoft-graph-calendar-api
- office-exchange-online
languages:
- javascript
---

# Microsoft Graph sample JavaScript single-page app

This sample demonstrates how to use the Microsoft Graph JavaScript SDK to access data in Office 365 from JavaScript browser apps.

## Prerequisites

Before you start this tutorial, you will need access to an HTTP server to host the sample files. This could be a test server on your development machine, or a remote server. The tutorial includes instructions to use a Node.js package to run a simple test server on your development machine. If you plan to use this option, you should have [Node.js](https://nodejs.org) installed on your development machine. If you do not have Node.js, visit the previous link for download options.

You should also have either a personal Microsoft account with a mailbox on Outlook.com, or a Microsoft work or school account. If you don't have a Microsoft account, there are a couple of options to get a free account:

- You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).
- You can [sign up for the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) to get a free Microsoft 365 subscription.

## Register the app

Create a new Azure AD web application registration using the Azure Active Directory admin center.

1. Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.

1. Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.

    > [!NOTE]
    > Azure AD B2C users may only see **App registrations (legacy)**. In this case, please go directly to https://aka.ms/appregistrations.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `JavaScript Graph Tutorial`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `http://localhost:8080`.

1. Choose **Register**. On the **JavaScript Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.

## Configure the sample

1. Rename the `./graph-tutorial/src/Config.example.ts` file to `./graph-tutorial/src/Config.ts`.
1. Edit the `./graph-tutorial/src/Config.ts` file and make the following changes.
1. Replace `YOUR_APP_ID_HERE` with the **Application Id** you got from the App Registration Portal.

> [!NOTE]
> The steps in this section require [Node.js](https://nodejs.org).

## Running the sample
In this section you will use [http-server](https://www.npmjs.com/package/http-server) to run a simple HTTP server from the command line.

1. Open your command-line interface (CLI) in the directory you created for the project.
1. Run the following command to start a web server in that directory.

    ```Shell
    npx http-server -c-1
    ```
1. Open your browser and browse to `http://localhost:8080`.

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
