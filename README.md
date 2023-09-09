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



### Setting Up Graph API Application and secret


### Setup On-Prem

    Install-Module Microsoft.Graph.Authentication
    Install-Module Microsoft.Graph.Users
    Install-Module Microsoft.Graph.PersonalContacts

### Setup with Azure Function
