#All Employees Access for moderators
$allemployees = Get-ADUser -Filter { (Department -like "*Teaching Faculty") -or (Department -like "*Administration Staff") }

#Filter out test acccounts 
$allemployees = $allemployees | ? {$_.Name -ne "TEST Account 1"}
$allemployees = $allemployees | ? {$_.Name -ne "TEST Account 2"}
$allemployees = $allemployees | ? {$_.Name -ne "TEST Account 3"}

#Faculty Mailboxes Access of All Employees for Managing Members
foreach($User in $allemployees){
    Add-MailboxPermission -Identity $User.name -User employee1@abc.org -AccessRights FullAccess -InheritanceType All -Confirm:$false
    Add-MailboxPermission -Identity $User.name -User employee2@abc.org -AccessRights FullAccess -InheritanceType All -Confirm:$false
}

#Faculty Mailboxes Email Account for HR
$allFaculty = Get-ADUser -Filter { (Department -like "*Teaching Faculty") }

foreach($User in $allFaculty){
    Add-MailboxPermission -Identity $User.name -User HRMailboxes@abc.org -AccessRights FullAccess -InheritanceType All -Confirm:$false
}