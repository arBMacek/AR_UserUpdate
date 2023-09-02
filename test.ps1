

$global:ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

. ("$ScriptDirectory\UserUpdate_Lib.ps1")


if (Get-Module -ListAvailable -Name SqlServer) {
    
    $EventMessage = "Required PowerShell module 'SqlServer' was found."
    Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 1001 -EntryType "Information" -Message $EventMessage
} 
else {

    try{
        Install-Module SqlServer -Scope CurrentUser -Force
    }
    catch{
        $EventMessage = "Required PowerShell module 'SqlServer' was found. Script tried to install it but ran into an issue: " + $_.Exception.Message
        Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 9101 -EntryType "Error" -Message $EventMessage
        Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage

        
        Exit 1
    }

}



$SQL = Get-Settings -Type 'SQL'

$sqlPW    = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($SQL.Cred.Password)))

$connString = "Data Source=$($SQL.Server);" `
            + "Database=$($SQL.Database);" `
            + "User ID=$($SQL.Cred.UserName);" `
            + "Password=$sqlPW" `

#Create a SQL connection object
$conn = New-Object System.Data.SqlClient.SqlConnection 

$conn.ConnectionString = $connString

#Attempt to open the connection


    $conn.Open()

    $sqlCommand = New-Object System.Data.SQLClient.SQLCommand
    $sqlCommand.Connection = $conn

    $sqlQuery_addADUsers="
    INSERT INTO $($SQL.ADUsers)
        ([EmployeeID],
         [FirstName],
         [LastName],
         [Email],
         [JobTitle],
         [Department],
         [Company],
         [City],
         [State],
         [ManagerEmail],
         [EmployeeType],
         [AccountEnabled])
        VALUES
        ('12345',
         'Brian',
         'Macek',
         'bmacek@alliedreliability.com',
         'IT boss',
         'IT',
         'Allied',
         'Houston',
         'TX',
         'Erick Flores',
         'Permanent',
         'Enabled')"



    $sqlCommand.CommandText = $sqlQuery_addADUsers
    $sqlCommand.ExecuteNonQuery()

<#
$Employee_Email = "tst.03@alliedreliability.com"
$Job_Title = "Shoe"

Get-ADUser -Filter "EmailAddress -eq '$($Employee_Email)'" | 
Set-ADUser -Title $Job_Title -Description $Job_Title
#>