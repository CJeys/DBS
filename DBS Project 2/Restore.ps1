# Christopher Jeys

Function ActiveDirectory {
    Write-Output -Foreground Blue '[AD]: Beginning Active Directory Tasks'
    $AdRoot = (Get-ADDomain).DistinguishedName
    $DnsRoot = (Get-ADDomain).DnsRoot
    $OUCanonicalName = 'Finance'
    $OUDisplayName = 'Finance'
    $ADPath = 'OU=$($OUCanonicalName),$(AdRoot)'

    if (-Not([ADSI]::Exists('LDAP://$($ADPath)'))) {
        New-ADOrganizationalUnit -Path $AdRoot -Name $OUCanonicalName -DisplayName $OUDisplayName -ProtectedFromAccidentalDeletion $false
        Write-Output -Foreground Blue '[AD]: $($OUCanonicalName) OU Created'
    }
    else {
        Write-Output '$($OUCanonicalName) Already Exists'
    }
    # Convert CSV file into a table
    $NewADUsers = Import-Csv -Path $PSScriptRoot\financePersonnel.csv

    # Values being set up for status
    $numberNewUsers = $NewADUsers.Count
    $count = 1

    # Iterate over each row in the table
    ForEach ($ADUser in $NewADUsers) {
        # Assigning variables to column values
        $First = $ADUser.First_Name
        $Last = $ADUser.Last_Name
        $Name = $First + " " + $Last
        $SamAcct = $ADUser.samAccount
        $UPN = "$($SamAcct)@$($DnsRoot)"
        $Postal = $ADUser.PostalCode
        $Office = $ADUser.OfficePhone
        $Mobile = $ADUser.MobilePhone

        # Create and show status
        $status = "[AD]: Adding AD User: $($Name) ($(Count)) of $($numberNewUsers))"
        Write-Progress  -Activity 'C916 Task 2 - Restore' -Status $status -PercentComplete (($count / $numberNewUsers) * 100)
        # Create Active Directory User with given values
        New-ADUser -GivenName $First `
            -Surname $Last `
            -Name $Name `
            -SamAccountName $SamAcct `
            -UserPrincipalName $UPN `
            -DisplayName $Name `
            -PostalCode $Postal `
            -MobilePhone $Mobile `
            -OfficePhone $Office `
            -Path $ADPath
        
        # Increment counter
        $count++
    }
    Write-Output -ForegroundColor Blue '[AD]: Active Directory Tasks Complete'
}
Function SQLServer {
    # Import SqlServer Module
    if (Get-Module -Name sqlps) { Remove-Module sqlps }
    Import-Module -Name SqlServer
    #$myObjectReference = New-Object -TypeName <Library and Type Name> -ArgumentList <Any arguments needed to create the object>

    # Setting a string variable naming the SQL Instance
    $sqlServerInstanceName = '.\SQLEXPRESS'

    # Creating an object referencing the SQL Server
    $sqlServerObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $sqlServerInstanceName

    # Setting a string variable naming the database
    $databaseName = 'ClientDB'

    # Creating an object referencing the DBS
    $databaseObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $sqlServerObject, $databaseName

    # Calling the create method on the database object for creation
    $databaseObject.Create()

    # Executing SQL code against the database
    Invoke-Sqlcmd -ServerInstance $sqlServerInstanceName -Database $databaseName -InputFile $PSScriptRoot\CreateTable_Client_A_Contacts.sql

    # Adding entries from CSV file
    $Insert = "INSERT INTO [$($tablename)] (first_name, last_name, city, county, zip, officePhone, mobilePhone)"

    # Creating the structure for new records
    $NewClientList = Import-Csv $PSScriptRoot\NewClientData.csv
    ForEach ($NewClient in $NewClientList) {
        $Values = "VALUES ( `
                        '$($NewClient.first_name)', `
                        '$($NewClient.last_name)', `
                        '$($NewClient.city)', `
                        '$($NewClient.county)', `
                        '$($NewClient.zip)', `
                        '$($NewClient.officePhone)', `
                        '$($NewClient.mobilePhone)',)"
        $query = $Insert + $Values
        Invoke-Sqlcmd -Database $databaseName - -ServerInstance $sqlServerInstanceName -Query $query
    }
}

try {
    ActiveDirectory
    SQLServer
}
#Error handling to catch System.OutOfMemoryException errors
Catch [System.OutOfMemoryException] {  
    Write-Output 'An Out of Memory Exception Occurred'
}
Finally {   
    # Closes open resources
}