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

### Setup On-Prem

    Install-Module Microsoft.Graph.Authentication
    Install-Module Microsoft.Graph.Users
    Install-Module Microsoft.Graph.PersonalContacts

### Setup with Azure Function
