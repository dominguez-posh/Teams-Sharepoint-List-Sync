# Teams-Sharepoint-List-Sync
A small sync tool for syncing a Sharepointlist with Private Adressbooks of Company Users using Powershell

## What Problem will be resolved ?
In actual release of Microsoft Teams, you wont find any built in soulution of using a global managed adressbook.
Every user can use his own Outlook adressbook but you are not able to use a shared adressbook within Teams, but only your private.

So the script solves that problem. You create a Sharepoint table in Sharepoint online in any location you actually want to.
This table will be the central Adressbook for your tenant.

The Script supports following actions in the Sharepoint list:
 - Adding brandnew entries to the lis
 - Editing entries. (The Script will only Update edited entrys to the adressbooks of the users)
 - Deleting entries from the Sharepoint list.

After deployment, the user will be still able to use his own Adress Entrys.
The synced contacts will get a flag in the "notes", to identify a entry as a synced entry.
So no current saved contact will be touched for the enduser.

Also you have a detailed output, what changes are made within a run.

Also the Script ist completely ready and tested in a Azure function. (Setup-Guid down below)

## How to Setup

There are two ways to Deploy the Script.
1. On-Prem (Manual Run or Scheduled in a Scheduled Windows Task)
2. On Azure with Azure funtions (Manual with URL-Query or automated with a time trigger)

### Setup Sharepoint List
First of all you need to setup a Sharepoint-Site.
Use the Template .xlsx File to setup the Sharepoint Site.
![image](https://github.com/dominguez-posh/Teams-Sharepoint-List-Sync/assets/9081611/49fa5a5d-6077-43b3-b351-05cf2f305709)

Select Excel and Upload the Template.

![image](https://github.com/dominguez-posh/Teams-Sharepoint-List-Sync/assets/9081611/76652a06-4e2f-465c-9953-f96727eaae93)

You can add some data to the Excel template if you want to Upload a adressbook initially. :-)

Change the data-type of the phonenumbers to "Text"

Now the List is ready to go !

You need Flowoing things to edit in the Script:

The Site-URL: This is the URL you see in the browser
The Site-ID: You can get it by adding  "_api/site/id" to the URL. Then you can see the Site-ID in your browser and you can put it into the script.


### Setting Up Graph API Application and secret

https://entra.microsoft.com/#view/Microsoft_AAD_IAM/StartboardApplicationsMenuBlade/~/AppAppsPreview
You need to create a new Application in Azure AD for the API-Token.

Create a New a client application to access a web API:
https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-configure-app-access-web-apis

Following rights are needed: 

![image](https://github.com/dominguez-posh/Teams-Sharepoint-List-Sync/assets/9081611/578d759d-d299-440f-bbb6-88662637f18c)

Create a new Client Secret, and add Tenant-ID, Application-ID and Client-Secret to the Script.

Now the preperation is done and you can decide how to run the script.

Personally I prefer running the script in a Azure function.
It is clientless, you do not need to think about it again (exept to renew the Clientsecret) and you can run it scheduled, but also with a URL-Call.
Also you have no cost at all. But of cause you can run it On-Prem as well

### Setup On-Prem
For setting up On-Prem you actually only need to install the following Modules:

    Install-Module Microsoft.Graph.Authentication
    Install-Module Microsoft.Graph.Users
    Install-Module Microsoft.Graph.PersonalContacts

When you added the Sharepoint-List and Graph API Informations, you are ready to Go :-)
First, try to add a new contact to the Sharepoint list and see the magic after running.

You can schedule the Script with a Scheduled Windows Task as well.

### Setup with Azure Function

To Run the Script in a Azure Function, you only need to Create a new default Powershell (Clientless) Azure Funktion.

Then you need to edit the profile.ps1 for performance and the requirements.ps1 for adding the needed Modules.

Now you can Add a Function. I Reccomend a URL or Time Call.

Paste the Script and you are ready to Go ! :) 

