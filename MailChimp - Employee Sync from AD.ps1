#This script syncs AD users to MailChimp.com utilized for marketing/communications dept synchronization of employee email addresses using REST API in Powershell
#Requires -RunAsAdministrator
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$AD_EmployeesList = Get-ADUser -Filter { (Enabled -eq $true) -and (Department -like "*Teaching Faculty") -or (Enabled -eq $true) -and (Department -like "*Administration Staff") } | select GivenName,Surname,UserPrincipalName
#filter out test accounts
$AD_EmployeesList = $AD_EmployeesList | ? {$_.UserPrincipalName -ne "test01@lihouston.org"}
$AD_EmployeesList = $AD_EmployeesList | ? {$_.UserPrincipalName -ne "test2@lihouston.org"}

$AD_BlockedEmployeesList = Get-ADUser -Filter { (Enabled -eq $false) -and (Department -like "*Teaching Faculty") -or (Enabled -eq $false) -and (Department -like "*Administration Staff") } | select GivenName,Surname,UserPrincipalName
$AD_BlockedEmployeesList = $AD_BlockedEmployeesList | ? {$_.UserPrincipalName -ne "test01@lihouston.org"}
$AD_BlockedEmployeesList = $AD_BlockedEmployeesList | ? {$_.UserPrincipalName -ne "test2@lihouston.org"}


$user = "LIHmarketingAdmin"
$apiKey = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
$pair = "${user}:${apiKey}"
$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)
$basicAuthValue = "Basic $base64"
$Headers = @{
    Authorization = $basicAuthValue
}
$baseUri = "https://xxxxx.api.mailchimp.com"

function EmployeeInfo($Employee){
    
    $merge_fields = @{
        FNAME = $Employee.EmployeeFirstName
        LNAME = $Employee.EmployeeLastName
    }
        
    $body = @{
        email_address = $Employee.EmployeeEmail
        merge_fields  = $merge_fields
        status = "subscribed"
    }

    $body = $body | ConvertTo-Json -Depth 10
    return $body
}

###Add Employees to MailChimp All Employees List###
#Get MailChimp Members
$MailChimp_EmployeesList = (Invoke-RestMethod -URI $baseUri/3.0/lists/xxxxxxx/members?"&"count=1000"&"status=subscribed -Method Get -Headers $Headers).members | select email_address

if($MailChimp_EmployeesList -eq $null){ #If List is empty
    foreach($Employee in $AD_EmployeesList){
        $EmployeeData=@{
            EmployeeEmail = $Employee.UserPrincipalName
            EmployeeFirstName = $Employee.GivenName
            EmployeeLastName = $Employee.Surname
        }

        $body = EmployeeInfo($EmployeeData) 
        Invoke-RestMethod -URI $baseUri/3.0/lists/xxxxxxx/members?skip_merge_validation=false"&"merge-fields -Method POST -Headers $Headers -Body $body
    }

} else { #If list contains data
    foreach($Employee in $AD_EmployeesList){
        if($MailChimp_EmployeesList -match $Employee.UserPrincipalName){
        } else {
            $EmployeeData=@{
                EmployeeEmail = $Employee.UserPrincipalName
                EmployeeFirstName = $Employee.GivenName
                EmployeeLastName = $Employee.Surname
            }

            $body = EmployeeInfo($EmployeeData) 
            Invoke-RestMethod -URI $baseUri/3.0/lists/xxxxxxx/members?skip_merge_validation=false"&"merge-fields -Method POST -Headers $Headers -Body $body
        }
    }
}

$MailChimp_EmployeesList = (Invoke-RestMethod -URI $baseUri/3.0/lists/xxxxxxx/members?"&"count=1000"&"status=subscribed -Method Get -Headers $Headers).members 
$AD_EmployeesList = Get-ADUser -Filter { (Enabled -eq $true) -and (Department -like "*Teaching Faculty") -or (Enabled -eq $true) -and (Department -like "*Administration Staff") } | select UserPrincipalName

#Unsubscribe Blocked Employee
if($MailChimp_EmployeesList -ne $null){ 
    foreach($Employee in $MailChimp_EmployeesList){
        if($AD_EmployeesList -match $Employee.email_address){
            #$Employee.UserPrincipalName+" is Subscribed"
        } else {
            #$Employee.UserPrincipalName+" is not Subscribed"
            $subscriber_hash = $Employee.id

            $body = @{
                status = "unsubscribed"
            }
            $body = $body | ConvertTo-Json -Depth 10

            Invoke-RestMethod -URI $baseUri/3.0/lists/xxxxxxx/members/$subscriber_hash/?skip_merge_validation=false -Method PATCH -Headers $Headers -Body $body
           
            $Employee.email_address+" has been unsubscribed"
        }
    }
}