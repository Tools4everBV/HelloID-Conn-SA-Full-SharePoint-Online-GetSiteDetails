<!-- Description -->
## Description
This HelloID Service Automation Delegated Form provides SharePoint Online functionality. The following options are available:
 1. Search site based on sitename
 2. Select a site from search results
 3. View the details 

## Versioning
| Version | Description | Date |
| - | - | - |
| 1.1.0   | Updated to Graph API with SA- & audit-logging | 2022/10/17  |
| 1.0.1   | Added version number and updated all-in-one script | 2021/12/13  |
| 1.0.0   | Initial release | 2020/12/05  |

<!-- Requirements -->
## Requirements
This script uses the Microsoft Graph API and requires an App Registration with App permissions:
*	Read all groups in an organization’s directory by using <b><i>Group.Read.All</i></b>
*	Allow the application to read documents and list items in all site collections on your behalf by using <b><i>Sites.Read.All</i></b>

<!-- TABLE OF CONTENTS -->
## Table of Contents
- [Description](#description)
- [Versioning](#versioning)
- [Requirements](#requirements)
- [Table of Contents](#table-of-contents)
- [Introduction](#introduction)
- [Getting the Azure AD graph API access](#getting-the-azure-ad-graph-api-access)
  - [Application Registration](#application-registration)
  - [Configuring App Permissions](#configuring-app-permissions)
  - [Authentication and Authorization](#authentication-and-authorization)
- [All-in-one PowerShell setup script](#all-in-one-powershell-setup-script)
  - [Getting started](#getting-started)
- [Post-setup configuration](#post-setup-configuration)
- [Manual resources](#manual-resources)
- [Getting help](#getting-help)
- [HelloID Docs](#helloid-docs)

## Introduction
The interface to communicate with Microsoft Azure AD is through the Microsoft Graph API.

<!-- GETTING STARTED -->
## Getting the Azure AD graph API access

By using this connector you will have the ability to enable or disable an Azure AD User.

### Application Registration
The first step to connect to Graph API and make requests, is to register a new <b>Azure Active Directory Application</b>. The application is used to connect to the API and to manage permissions.

* Navigate to <b>App Registrations</b> in Azure, and select “New Registration” (<b>Azure Portal > Azure Active Directory > App Registration > New Application Registration</b>).
* Next, give the application a name. In this example we are using “<b>HelloID PowerShell</b>” as application name.
* Specify who can use this application (<b>Accounts in this organizational directory only</b>).
* Specify the Redirect URI. You can enter any url as a redirect URI value. In this example we used http://localhost because it doesn't have to resolve.
* Click the “<b>Register</b>” button to finally create your new application.

Some key items regarding the application are the Application ID (which is the Client ID), the Directory ID (which is the Tenant ID) and Client Secret.

### Configuring App Permissions
The [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph) provides details on which permission are required for each permission type.

To assign your application the right permissions, navigate to <b>Azure Portal > Azure Active Directory >App Registrations</b>.
Select the application we created before, and select “<b>API Permissions</b>” or “<b>View API Permissions</b>”.
To assign a new permission to your application, click the “<b>Add a permission</b>” button.
From the “<b>Request API Permissions</b>” screen click “<b>Microsoft Graph</b>”.
For this connector the following permissions are used as <b>Application permissions</b>:
*	Read all groups in an organization’s directory by using <b><i>Group.Read.All</i></b>
*	Allow the application to read documents and list items in all site collections on your behalf by using <b><i>Sites.Read.All</i></b>

Some high-privilege permissions can be set to admin-restricted and require an administrators consent to be granted.

To grant admin consent to our application press the “<b>Grant admin consent for TENANT</b>” button.

### Authentication and Authorization
There are multiple ways to authenticate to the Graph API with each has its own pros and cons, in this example we are using the Authorization Code grant type.

*	First we need to get the <b>Client ID</b>, go to the <b>Azure Portal > Azure Active Directory > App Registrations</b>.
*	Select your application and copy the Application (client) ID value.
*	After we have the Client ID we also have to create a <b>Client Secret</b>.
*	From the Azure Portal, go to <b>Azure Active Directory > App Registrations</b>.
*	Select the application we have created before, and select "<b>Certificates and Secrets</b>". 
*	Under “Client Secrets” click on the “<b>New Client Secret</b>” button to create a new secret.
*	Provide a logical name for your secret in the Description field, and select the expiration date for your secret.
*	It's IMPORTANT to copy the newly generated client secret, because you cannot see the value anymore after you close the page.
*	At least we need to get is the <b>Tenant ID</b>. This can be found in the Azure Portal by going to <b>Azure Active Directory > Custom Domain Names</b>, and then finding the .onmicrosoft.com domain.

## All-in-one PowerShell setup script
The PowerShell script "createform.ps1" contains a complete PowerShell script using the HelloID API to create the complete Form including user defined variables, tasks and data sources.

 _Please note that this script asumes none of the required resources do exists within HelloID. The script does not contain versioning or source control_

 
## Post-setup configuration
After the all-in-one PowerShell script has run and created all the required resources. The following items need to be configured according to your own environment
 1. Update the following [user defined variables](https://docs.helloid.com/hc/en-us/articles/360014169933-How-to-Create-and-Manage-User-Defined-Variables)
<table>
  <tr><td><strong>Variable name</strong></td><td><strong>Example value</strong></td><td><strong>Description</strong></td></tr>
  <tr><td>AADtenantID</td><td>Azure AD Tenant Id</td><td>Id of the Azure tenant</td></tr>
  <tr><td>AADAppId</td><td>Azure AD App Id</td><td>Id of the Azure app</td></tr>
  <tr><td>AADAppSecret</td><td>Azure AD App Secret</td><td>Secreat of the Azure app</td></tr>
</table>

## Manual resources
This Delegated Form uses the following resources in order to run

### Powershell data source '[powershell-datasource]_Sharepoint-generate-table-sites-wildcard'
This Powershell data source performs a search on available SharePoint sites.

### Powershell data source '[powershell-datasource]_Sharepoint-get-site-details'
This Powershell data source gets the details of the selected SharePoint site.

## Getting help
_If you need help, feel free to ask questions on our [forum](https://forum.helloid.com/forum/helloid-connectors/service-automation/179-helloid-sa-sharepoint-online-get-site-details)_

## HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/