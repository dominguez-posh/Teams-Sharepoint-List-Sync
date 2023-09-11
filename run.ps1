#  A Script for syncing a Sharepoint list with contacts for using it as global Adresslist in Teams for every user
#  Author: dominguez-posh : https://github.com/dominguez-posh/
## Version: 1.0
using namespace System.Net
param($Request, $TriggerMetadata)

#Required API Permissions
#Contacts.ReadWrite
#User.ReadWrite.All

#

$global:SPOListURI = "https://contooso.com/sites/AllCompany.11168440321.gykonuin/"
$global:SPOListName = "Contacts"
$global:SPOSiteID = "64094f6f-12e1-4be5-b350-66ee86b5d2a3"

$SiteID = $SPOListURI + "_api/site/id"

$AppSecret = "X8w8R!BorTqp6nmdG6871jtzcr~OnuzS5VV!Oa59"
$AppId = "2c3aa1a5-bcc2-4c32-8daa-0decc310095b"
$Tenantid = "aa9b233829c82-4010-93b9-967352daee96"


#$ProcessFilter = "*azure admin*" # Set This for Testing :) 
$ProcessFilter = "*"

###################################################################################################

$returnbody = @()
$returnbody += ("---Processing Log---")

$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}

$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

$Connection = Connect-MgGraph -AccessToken (ConvertTo-SecureString $token -AsPlainText -Force)

function Get-SPOListContacts {
Param()
    $Header = @{
        Authorization = $token
    }
    
   $GraphUrl = "https://graph.microsoft.com/v1.0/sites/$SPOSiteID/lists/$SPOListName/items?expand=fields"

   $ListItems = (Invoke-RestMethod -Headers $Header -Uri $GraphUrl -UseBasicParsing -Method "GET").value

    $Contacts = @()

    foreach ($Item in $ListItems){

        $Values = $Item.Fields
        $Contact = [PSCustomObject]@{

            Vorname            = $Values.field_1
            Nachname           = $Values.Title
            Firma              = $Values.field_3
            Email              = $Values.field_2
            Department         = $Values.field_4
            Telefonnummer      = $Values.field_5
            Telefonnummer2     = $Values.field_6
            TelefonnummerMobil = $Values.field_7
            SpoListId          = $Values.ID

        }
        $Contacts += $Contact
  
    }

    return $Contacts

}


$SharePointListContacts = Get-SPOListContacts

$Users = Get-MgUser -All | ? DisplayName -like $ProcessFilter


foreach ($User in $Users){
    
    $Err = $False
    $UserContacts = @()

    try{
       $UserContacts = Get-MgUserContact -UserId $User.Id -erroraction Stop | ? PersonalNotes -like "*SPOLISTID:*" 
    }
    catch{
        $Err = $True
    }

    if($Err){}
    else{

        Write-Host "Processing Contact Updates for User : "  $User.DisplayName
        $returnbody += ("Processing Contact Updates for User : " + $User.DisplayName).tostring()


        $UserContacts = @()
        try{
            $UserContacts = Get-MgUserContact -UserId $User.Id -all -erroraction Stop | ? PersonalNotes -like "*SPOLISTID:*" 
        }
        catch{
            $Err = $True
        }


    

        $ContactsToAdd = @()
        $ContactsToUpdate = @()
        $ContactsToDelete = @()

        #Getting Contacts to Add
        foreach($SharePointListContact in $SharePointListContacts){
        $ToAdd = $True
        foreach($UserContact in $UserContacts){
            if($UserContact.PersonalNotes -like ("SPOLISTID:" + $SharePointListContact.SpoListId)){$ToAdd = $False}
        }
        if($ToAdd){$ContactsToAdd += $SharePointListContact}

    }
        #Getting Contacts to Update
        foreach($SharePointListContact in $SharePointListContacts){
        $ToUpdate = $False
        foreach($UserContact in $UserContacts){
            if($UserContact.PersonalNotes -like ("SPOLISTID:" + $SharePointListContact.SpoListId)){
                if($UserContact.GivenName -notlike $SharePointListContact.Vorname){$ToUpdate = $True }
                if($UserContact.Surname -notlike $SharePointListContact.Nachname){$ToUpdate = $True }
                if($UserContact.CompanyName -notlike $SharePointListContact.Firma){$ToUpdate = $True }
                if($UserContact.EmailAddresses[0].Address -notlike $SharePointListContact.Email){$ToUpdate = $True }
                if($UserContact.Department -notlike $SharePointListContact.Department){$ToUpdate = $True  }
                if($UserContact.BusinessPhones[0] -notlike $SharePointListContact.Telefonnummer){$ToUpdate = $True }
                if($UserContact.BusinessPhones[1] -notlike $SharePointListContact.Telefonnummer2){$ToUpdate = $True }
                if($UserContact.MobilePhone -notlike $SharePointListContact.TelefonnummerMobil){$ToUpdate = $True }
            }
        }
        if($ToUpdate){$ContactsToUpdate += $SharePointListContact}

    }
        #Getting Contacts to Delete
        foreach($UserContact in $UserContacts){
        $Contactdelete = $True
        foreach($SharePointListContact in $SharePointListContacts){
            if($UserContact.PersonalNotes -like ("SPOLISTID:" + $SharePointListContact.SpoListId)){ $Contactdelete = $False }
        }
        if($Contactdelete){$ContactsToDelete += $UserContact}
    
        }



    if(-not ($ContactsToAdd -or $ContactsToUpdate -or $ContactsToDelete)){
            Write-Host  ("--- Nothing to Process.")
            $returnbody += ("--- Nothing to Process.")
        }

    


        ##Adding Contacts#

        foreach($ContactToAdd in $ContactsToAdd){
        Write-Host  ("--- Adding Contact: " + $ContactToAdd.Vorname + " " + $ContactToAdd.Nachname + " | " + $ContactToAdd.Firma)
        $returnbody += ("--- Adding Contact: " + $ContactToAdd.Vorname + " " + $ContactToAdd.Nachname + " | " + $ContactToAdd.Firma)

        $params = @{
	        givenName = $ContactToAdd.Vorname
	        surname = $ContactToAdd.Nachname
            companyname = $ContactToAdd.Firma
            mobilephone = $ContactToAdd.TelefonnummerMobil
            department = $ContactToAdd.Department
            JobTitle = $ContactToAdd.Department
            PersonalNotes = ("SPOLISTID:" + $ContactToAdd.SpoListId)
            displayname = ($ContactToAdd.Vorname + " " + $ContactToAdd.Nachname + " | " + $ContactToAdd.Firma)
            ImAddresses = @(
                $ContactToAdd.Email
            )
	        emailAddresses = @(
		        @{
			        address = $ContactToAdd.Email
			        name = ($ContactToAdd.Vorname + " " + $ContactToAdd.Nachname)
		        }
	        )
	        businessPhones = @(
		        $ContactToAdd.Telefonnummer
                $ContactToAdd.Telefonnummer2
	        )
        }

        $Output = New-MgUserContact -UserId $User.Id -BodyParameter $params
    }

        ##Updating Contacts

        foreach($ContactToUpdate in $ContactsToUpdate) {
        
        Write-Host  ("--- Updating Contact: " + $ContactToUpdate.Vorname + " " + $ContactToUpdate.Nachname + " | " + $ContactToUpdate.Firma)
        $returnbody += ("--- Updating Contact: " + $ContactToUpdate.Vorname + " " + $ContactToUpdate.Nachname + " | " + $ContactToUpdate.Firma)
        $params = @{
	        givenName = $ContactToUpdate.Vorname
	        surname = $ContactToUpdate.Nachname
            companyname = $ContactToUpdate.Firma
            department = $ContactToUpdate.Department
            JobTitle = $ContactToUpdate.Department
            PersonalNotes = ("SPOLISTID:" + $ContactToUpdate.SpoListId)
            displayname = ($ContactToUpdate.Vorname + " " + $ContactToUpdate.Nachname + " | " + $ContactToUpdate.Firma)
            ImAddresses = @(
                $ContactToUpdate.Email
            )
	        emailAddresses = @(
		        @{
			        address = $ContactToUpdate.Email
			        name = ($ContactToUpdate.Vorname + " " + $ContactToUpdate.Nachname)
		        }
	        )
	        businessPhones = @(
		        $ContactToUpdate.Telefonnummer
                $ContactToUpdate.Telefonnummer2
	        )
            mobilephone = $ContactToUpdate.TelefonnummerMobil
        }

        $ContactID = ($UserContacts | ? PersonalNotes  -like ("SPOLISTID:" + $ContactToUpdate.SpoListId)).id

        $UpdateUser = Update-MgUserContact -UserId $User.Id -ContactId $ContactID -BodyParameter $params

    }

        ##Deleting Contacts

        foreach($ContactToDelete in $ContactsToDelete){

        Write-Host  ("--- Deleting Contact: " + $ContactToDelete.DisplayName)
        $returnbody += ("--- Deleting Contact: " + $ContactToDelete.DisplayName)

        #Remove-MgUserContact -UserId $User.Id -ContactId $ContactToDelete.Id

    }

    }

}

        Write-Host  "Processing Done."
        $returnbody += "Processing Done."

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $returnbody
})
