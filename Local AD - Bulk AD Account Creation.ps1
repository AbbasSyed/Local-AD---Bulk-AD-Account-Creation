#This script creates AD user accounts from excel import file. 
#It adds a faculty member and staff member based on keywords from their title.

$EmployeeLogonPW = "XXXXXXXXXX"
$global:RandomPin = $null
$EmployeeCount=0

#The CSV file uses the column information from Office365's bulk import feature. 
#The CSV File can also be directly imported into Office365 without the use of this script.
$EmployeeList = Import-Csv -Path "C:\Users\XXXXX\Documents\Bulk_AD_User_Creation_List.csv"

#The script allows to add multiple rows of Emplpoyee Information to be imported from the CSV file
foreach($Employee in $EmployeeList){
    EmployeeCopierPinNo
    $EmployeeName = $Employee."First Name"+" "+$Employee."Last Name"
    $EmployeeFirstName=$Employee."First Name"
    $EmployeeLastName=$Employee."Last Name"
    $EmployeeJobTitle=$Employee."Job title"
    $EmployeePersonalEmail=$Employee."Alternate email address"
    $EmployeePhone = $Employee."Mobile phone"
    
    $EmployeePinNo="pin "+$global:RandomPin
    
    if($Employee."Job Title".Contains("Teacher")){ 
        $EmployeeDepartment= "Teaching Faculty"
    } else{
        $EmployeeDepartment= "Administration Staff"
    }

    #$EmployeeName
    $EmployeeUPN = $Employee.Username.ToLower()
    $EmployeeUPN = $EmployeeUPN+"@lihouston.org"
   
    #Creates a new AD user account using the imported columns from the CSV file
    New-ADUser -Name $EmployeeName -UserPrincipalName $EmployeeUPN -SamAccountName $Employee.Username.ToLower() -givenName $EmployeeFirstName -Surname $EmployeeLastName -DisplayName $EmployeeName -title $EmployeeJobTitle -Department $EmployeeDepartment -Description $EmployeePinNo -EmailAddress $EmployeeEmail -Enabled $True -ChangePasswordAtLogon $True -AccountPassword (ConvertTo-SecureString $EmployeeLogonPW -AsPlainText -Force) -path "OU=USERS, OU=XXXXXXXXX, DC=XXXXXXXXX, DC=ORG"
    
    #Adds the Personal Email Address of Employee in AD "Notes" attribute under the "Telephones" Tab. 
    #The personal email address is taken from the column "Alternate email address" in the imported CSV file. 
    Set-ADUser $Employee.Username.ToLower() -add @{info=$EmployeePersonalEmail}  

    #Test Code
    #New-ADUser -Name $EmployeeName -path "OU=USERS, OU=XXXXXXX, DC=XXXXXXX, DC=ORG"

    if($Employee."Job Title".Contains("Teacher")){ 
        Add-ADGroupMember -Identity "Teachers" -Members $Employee.Username.ToLower()    
    } else{
        Add-ADGroupMember -Identity "Administration" -Members $Employee.Username.ToLower()  
    }
    
    $Employee."First Name"+" "+$Employee."Last Name"+" "+"("+$EmployeeUPN+")"
    $EmployeeCount++
}
    " "
    "Employees Added: "+$EmployeeCount

function EmployeeCopierPinNo {
    $allemployees = Get-ADUser -Filter * -SearchBase "OU=USERS, OU=XXXXXXXXX, DC=XXXXXXXXX, DC=ORG" -Properties *
    
    $PinList= @()
    
         foreach($Employee in $allemployees){
            if($Employee.Description.Length -gt 3){
                $PinList += $Employee.Description.Substring(4)
            }
         }
    
        do{
            $global:RandomPin = Get-Random -Minimum 1000 -Maximum 10000          
        }while($PinList -eq $global:RandomPin)
}