
############################################################################################################ 
############################################################################################################
############################################################################################################


# AR USER UPDATE v1.0
############################################################################################################ 
############################################################################################################
############################################################################################################


$EventMessage = "UserUpdate_Main.ps1 was launched. Initilizing user attribute sync."
Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 0001 -EntryType "Information" -Message $EventMessage


############################################################################################################
# Setup metric variables
############################################################################################################

# initializes variable for tracking total number of updates that were made
$global:Total_Updates = 0

# starts a stopwatch for measurning how long it took the script to run
$Script_Timer = [Diagnostics.Stopwatch]::StartNew()

$html_allUpdateRecords = ""
$NameChange = $false


############################################################################################################
# Get script includes
############################################################################################################


$global:ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
try {
    . ("$ScriptDirectory\UserUpdate_Lib.ps1")

}
catch {
    $EventMessage = "A required script file 'UserUpdate_Lib.ps1' could not be found: " + $_.Exception.Message
    Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 9100 -EntryType "Error" -Message $EventMessage
    Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage

    Exit 1
}



############################################################################################################
# Install required script modules
############################################################################################################


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


if (Get-Module -ListAvailable -Name ActiveDirectory) {
    
    $EventMessage = "Required PowerShell module 'ActiveDirectory' was found."
    Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 1002 -EntryType "Information" -Message $EventMessage
} 
else {

    try{
        Install-Module ActiveDirectory -Scope CurrentUser -Force
    }
    catch{
        $EventMessage = "Required PowerShell module 'ActiveDirectory' was found. Script tried to install it but ran into an issue: " + $_.Exception.Message
        Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 9102 -EntryType "Error" -Message $EventMessage
        Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage

        Exit 1
    }

}


############################################################################################################
# Import Paycom data
############################################################################################################

$PaycomFile = Import-PaycomSFTP-File


# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# Add these variables to the email

Write-Host $PaycomFile.Name
Write-Host $PaycomFile.Date

# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


############################################################################################################
# Run job "Allied - Daily Paycom Import" on the SQL server. Imports data from Paycom file into SQL
############################################################################################################


try
{
    $EventMessage = "Starting SQL job '$($SQL.ImportJobName)'."
    Write-EventLog -LogName "AR User Update" -Source "SQL Job" -EventId 1004 -EntryType "Information" -Message $EventMessage
    
    $RunQuery = "EXEC [msdb].[dbo].[sp_start_job] @job_name = N'$($SQL.ImportJobName)';"

    $sqlJob = Start-Job {Invoke-Sqlcmd -ServerInstance $SQL.Server -Username ($using:SQL.Cred.UserName) -Password $sqlPW -Database $SQL.Database -Query $RunQuery}
    $sqlJob | Wait-Job

    $EventMessage = "SQL job '$($SQL.ImportJobName)' has finished running"
    Write-EventLog -LogName "AR User Update" -Source "SQL Job" -EventId 1004 -EntryType "Information" -Message $EventMessage
}
catch 
{
    $EventMessage = "There was an issue starting SQL job '$($SQL.ImportJobName)' : " + $_.Exception.Message
    Write-EventLog -LogName "AR User Update" -Source "SQL Job" -EventId 9104 -EntryType "Error" -Message $EventMessage
    Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
}



############################################################################################################
# Get all AD users
############################################################################################################

$AD = Get-Settings -Type 'AD'

try {

    $Users = Get-ADUser -Filter * -Properties * -SearchBase $($AD.SearchBase) | 
    Where-Object { ($_.DistinguishedName -notlike '*OU=Service Accounts*') } | 
    Select -Property GivenName,
                     SurName,
                     mail,
                     UserPrincipalName,
                     Title,
                     Department,
                     Manager,
                     Company,
                     City,
                     State,
                     Country,
                     Enabled,
                     employeeID,
                     employeeType


    $global:Total_ADUsers = $Users.Count

    # GET MANAGER EMAIL FROM DN
    foreach ($User in $Users)
    {
        if($User.Manager)
        {
            $ManagerDN = $User.Manager
            $User.Manager = Get-AdUser -Filter {DistinguishedName -eq $ManagerDN} -Properties * | Select -ExpandProperty mail
        }
        else
        {
            # IF MANAGER NOT SET, VERIFY USER IS CEO
            # IF NOT CEO, TRY TO FIND MANAGER EMAIL IN PAYCOM AND UPDATE USER
            # IF CAN'T FIND IT OR UPDATE USER, THROW ERROR    

        }
   }

   $EventMessage = "Successfully retrieved all users from Active Directory"
   Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 1004 -EntryType "Information" -Message $EventMessage

}
catch {
    $EventMessage = "There was an issue retrieving the users from Active Directory: " + $_.Exception.Message
    Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 9104 -EntryType "Error" -Message $EventMessage
    Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage

    Exit 1
}



############################################################################################################
# Add AD users to sql table
############################################################################################################


$SQL = Get-Settings -Type 'SQL'

$sqlPW    = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($SQL.Cred.Password)))

$connString = "Data Source=$($SQL.Server);" `
            + "Database=$($SQL.Database);" `
            + "TrustServerCertificate=True;" `
            + "trusted_connection=true;" `
            + "Integrated Security=True;" `
            + "encrypt=false;" `
            + "User ID=$($SQL.Cred.UserName);" `
            + "Password=$sqlPW" `

            #

#Create a SQL connection object
$conn = New-Object System.Data.SqlClient.SqlConnection 

$conn.ConnectionString = $connString

#Attempt to open the connection

try{
    $conn.Open()

    $EventMessage = "Connected to SQL server $sqlServer"
    Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 1005 -EntryType "Information" -Message $EventMessage
}
catch{
    $EventMessage = "There was an issue connecting to the SQL server $sqlServer : " + $_.Exception.Message
    Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 9105 -EntryType "Error" -Message $EventMessage
    Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
}




if($conn.State -eq "Open")
{
    $sqlCommand = New-Object System.Data.SQLClient.SQLCommand
    $sqlCommand.Connection = $conn

    #=======================================================================================================
    # See if table all_ad_users already exists
    #=======================================================================================================


    $sqlQuery_TableCheck = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '$($SQL.ADUsers)'"
    $sqlCommand.CommandText = $sqlQuery_TableCheck
    $tableExists = $sqlCommand.ExecuteScalar()

    if($tableExists)
    {
        #---------------------------------------------------------------------------------------------------
        # Delete table all_ad_users
        #---------------------------------------------------------------------------------------------------


        $sqlQuery_TableDelete = "DROP TABLE $($SQL.ADUsers)"
        $sqlCommand.CommandText = $sqlQuery_TableDelete
        $sqlCommand.ExecuteNonQuery()

     


        #---------------------------------------------------------------------------------------------------
        # Create new instance of the table all_ad_users if it doesn't exist
        #---------------------------------------------------------------------------------------------------


        $sqlQuery_TableCreate = "CREATE TABLE $($SQL.ADUsers) (
                                 EmployeeID varchar(50),
                                 FirstName varchar(50),
                                 LastName varchar(50),
                                 Email varchar(50),
                                 JobTitle varchar(50),
                                 Department varchar(50),
                                 Company varchar(50),
                                 City varchar(50),
                                 State varchar(50),
                                 ManagerEmail varchar(50),
                                 EmployeeType varchar(50),
                                 AccountEnabled varchar(50)
                                 )"
        
        $sqlCommand.CommandText = $sqlQuery_TableCreate
        $sqlCommand.ExecuteNonQuery()
    }
    else
    {
        $sqlQuery_TableCreate = "CREATE TABLE $($SQL.ADUsers) (
                                 EmployeeID varchar(50),
                                 FirstName varchar(50),
                                 LastName varchar(50),
                                 Email varchar(50),
                                 JobTitle varchar(50),
                                 Department varchar(50),
                                 Company varchar(50),
                                 City varchar(50),
                                 State varchar(50),
                                 ManagerEmail varchar(50),
                                 EmployeeType varchar(50),
                                 AccountEnabled varchar(50)
                                 )"

        $sqlCommand.CommandText = $sqlQuery_TableCreate
        $sqlCommand.ExecuteNonQuery()
    }



    #=======================================================================================================
    # Add AD users to table all_ad_users
    #=======================================================================================================


    foreach($User in $Users)
    {
        
        #if($User.employeeType -ne 'Service' -and $User.employeeType -ne 'Test')

        if($User.employeeType -ne 'Service')
        {
            
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
                ('$($User.employeeID)',
                 '$($User.GivenName)',
                 '$($User.SurName)',
                 '$($User.mail)',
                 '$($User.Title)',
                 '$($User.Department)',
                 '$($User.Company)',
                 '$($User.City)',
                 '$($User.State)',
                 '$($User.Manager)',
                 '$($User.employeeType)',
                 '$($User.Enabled)')"

            $sqlCommand.CommandText = $sqlQuery_addADUsers

            try{
                $sqlCommand.ExecuteNonQuery()
            }
            catch{
                Write-Host $sqlQuery_addADUsers
                Write-Host $User
                Write-Host $_.Exception.Message
            }

        }
        
    }




############################################################################################################
# Record how many totoal records are being compared
############################################################################################################


$sqlQuery_countPaycomData = "SELECT COUNT('AR_Employee_ID')
                        FROM $($SQL.PaycomData)
                        WHERE $($SQL.PaycomData).Employee_Status <> 'Retired'
                        AND $($SQL.PaycomData).Employee_Status <> 'Deceased'
                        AND $($SQL.PaycomData).Employee_Status <> 'Terminated'"

$sqlCommand.CommandText = $sqlQuery_countPaycomData

$global:Total_PaycomData = [Int32]$sqlCommand.ExecuteScalar()

$EventMessage = "Comparing $Total_PaycomData Paycom records against $Total_ADUsers Active Directory users."
Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 2000 -EntryType "Information" -Message $EventMessage
            




    


############################################################################################################
# Compare all_ad_users with paycom_data tables
############################################################################################################
    

    #=======================================================================================================
    # Check EmployeeID
    #=======================================================================================================
    
    # quesry gets a list of ad users who don't have their EmployeeID atrribute set.
    # returns email address and employee iD from employee table
    
    # build query

    $sqlQuery_compare_EmployeID = "SELECT $($SQL.PaycomData).Employee_Email,
                                          $($SQL.ADUsers).FirstName,
                                          $($SQL.ADUsers).LastName,
                                          $($SQL.PaycomData).AR_Employee_ID
                                   FROM $($SQL.PaycomData), $($SQL.ADUsers)
                                   WHERE $($SQL.PaycomData).Employee_Status <> 'Retired'
                                   AND $($SQL.PaycomData).Employee_Status <> 'Deceased'
                                   AND $($SQL.PaycomData).Employee_Status <> 'Terminated'
                                   AND $($SQL.PaycomData).Employee_Email = $($SQL.ADUsers).Email
                                   AND $($SQL.PaycomData).AR_Employee_ID <> $($SQL.ADUsers).EmployeeID"

    $sqlCommand.CommandText = $sqlQuery_compare_EmployeID
    $sqlResult = $sqlCommand.ExecuteReader()

    $sqlResult_update_EmployeeID = New-Object System.Data.DataTable
    $sqlResult_update_EmployeeID.Load($sqlResult)

   
    foreach($Employee in $sqlResult_update_EmployeeID)
    {
        try
        {
            Get-ADUser -Filter "EmailAddress -eq '$($Employee.Employee_Email)'" | 
            Set-ADUser -EmployeeID $Employee.AR_Employee_ID

            $Total_Updates += 1

            $html_allUpdateRecords += BuildReport -FirstName $Employee.FirstName `
                                                  -LastName $Employee.LastName `
                                                  -Email $Employee.Employee_Email `
                                                  -Attribute "EmployeeID" `
                                                  -Value $Employee.AR_Employee_ID

            $EventMessage = "The EmployeeID for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)) was updated to $($Employee.AR_Employee_ID)"
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 2001 -EntryType "Information" -Message $EventMessage
            
        }
        catch
        {
            $EventMessage = "There was an issue updating the EmployeeID for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)): " + $_.Exception.Message
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 9201 -EntryType "Error" -Message $EventMessage
            Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
        }
        
    }



    #=======================================================================================================
    # Check for name updates; notify itsupport if detected
    #=======================================================================================================

    $sqlQuery_compare_FirstName = "SELECT $($SQL.PaycomData).Employee_Email,
                                          $($SQL.PaycomData).Legal_First_Name, 
                                          $($SQL.PaycomData).Preferred_Firstname,
                                          $($SQL.PaycomData).Legal_Lastname,
                                          $($SQL.ADUsers).FirstName,
                                          $($SQL.ADUsers).LastName
                                   FROM $($SQL.PaycomData), $($SQL.ADUsers)
                                   WHERE $($SQL.PaycomData).Employee_Status <> 'Retired'
                                   AND $($SQL.PaycomData).Employee_Status <> 'Deceased'
                                   AND $($SQL.PaycomData).Employee_Status <> 'Terminated'
                                   AND $($SQL.PaycomData).Employee_Email = $($SQL.ADUsers).Email"
    
    $sqlCommand.CommandText = $sqlQuery_compare_FirstName
    $sqlResult = $sqlCommand.ExecuteReader()

    $sqlResult_update_FirstName = New-Object System.Data.DataTable
    $sqlResult_update_FirstName.Load($sqlResult)


    foreach($Employee in $sqlResult_update_FirstName)
    {
        $Employee_NameChanged = $false
        $Employee_Firstname = ""
        $Employee_Lastname = ""

    
        #---------------------------------------------------------------------------------------------------
        # Check First Name
        #---------------------------------------------------------------------------------------------------
    
        if (($Employee.Preferred_Firstname -ne [System.DBNull]::Value) -AND ($Employee.Preferred_Firstname -ne $Employee.FirstName))
        {
            #Write-Host "Setting Preferred FirstName: $($Employee.Preferred_Firstname)"

            $Employee_NameChanged = $true
            $Employee_Firstname = $Employee.Preferred_Firstname

        }
        elseif ($Employee.Legal_First_Name -ne $Employee.FirstName)
        {
            #Write-Host "Setting FirstName"
            
            $Employee_NameChanged = $true
            $Employee_Firstname = $Employee.Legal_First_Name
        }
        else
        {
            $Employee_Firstname = $Employee.FirstName
            
        }


        #---------------------------------------------------------------------------------------------------
        # Check Last Name
        #---------------------------------------------------------------------------------------------------

        if ($Employee.Legal_Lastname -ne $Employee.LastName)
        {
            $Employee_NameChanged = $true
            $Employee_Lastname = $Employee.Legal_Lastname

        }
        else {
            $Employee_Lastname = $Employee.LastName
        }


        #---------------------------------------------------------------------------------------------------
        # Email ITSupport about name change
        #---------------------------------------------------------------------------------------------------

        if ($Employee_NameChanged -eq $true)
        {
            #Send email to helpdesk

            $EventMessage = "Employee name change detected for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)). 
                             `nName changed to $Employee_Firstname $Employee_Lastname.
                             `nEmailing itsupport@alliedreliability.com to manually update Active Directory"
            
                             Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 2002 -EntryType "Information" -Message $EventMessage
            
            Notify_NameChange -old_FirstName $Employee.FirstName `
                              -old_LastName $Employee.LastName `
                              -old_Email $Employee.Employee_Email `
                              -new_FirstName $Employee_Firstname `
                              -new_LastName $Employee_Lastname

            $NameChange = $true

        }

    }



    #===========================================================================================================
    # Check JobTitle
    #===========================================================================================================
    
    #update job title and description

    $sqlQuery_compare_JobTitle = "SELECT $($SQL.PaycomData).Employee_Email,
                                         $($SQL.ADUsers).FirstName,
                                         $($SQL.ADUsers).LastName,
                                         $($SQL.PaycomData).Job_Title
                                FROM $($SQL.PaycomData), $($SQL.ADUsers)
                                WHERE $($SQL.PaycomData).Employee_Status <> 'Retired'
                                AND $($SQL.PaycomData).Employee_Status <> 'Deceased'
                                AND $($SQL.PaycomData).Employee_Status <> 'Terminated'
                                AND $($SQL.PaycomData).Employee_Email = $($SQL.ADUsers).Email
                                AND $($SQL.PaycomData).Job_Title <> $($SQL.ADUsers).JobTitle"

    $sqlCommand.CommandText = $sqlQuery_compare_JobTitle
    $sqlResult = $sqlCommand.ExecuteReader()

    $sqlResult_update_JobTitle = New-Object System.Data.DataTable
    $sqlResult_update_JobTitle.Load($sqlResult)
    
    foreach($Employee in $sqlResult_update_JobTitle)
    {

        try
        {
            Get-ADUser -Filter "EmailAddress -eq '$($Employee.Employee_Email)'" | 
            Set-ADUser -Title $Employee.Job_Title -Description $Employee.Job_Title
            
            $Total_Updates += 1

            $html_allUpdateRecords += BuildReport -FirstName $Employee.FirstName `
                                                         -LastName $Employee.LastName `
                                                         -Email $Employee.Employee_Email `
                                                         -Attribute "Job Title" `
                                                         -Value $Employee.Job_Title
            
            $EventMessage = "The Job Title for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)) was updated to $($Employee.Job_Title)"
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 2003 -EntryType "Information" -Message $EventMessage
        }
        catch
        {
            $EventMessage = "There was an issue updating the Job Title for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)): " + $_.Exception.Message
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 9203 -EntryType "Error" -Message $EventMessage
            Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
        }

    }


    
    #===========================================================================================================
    # Check Department
    #===========================================================================================================
    

    $sqlQuery_compare_Department = "SELECT $($SQL.PaycomData).Employee_Email, 
                                           $($SQL.PaycomData).Department,
                                           $($SQL.ADUsers).FirstName,
                                           $($SQL.ADUsers).LastName                                      
                                FROM $($SQL.PaycomData), $($SQL.ADUsers)
                                WHERE $($SQL.PaycomData).Employee_Status <> 'Retired'
                                AND $($SQL.PaycomData).Employee_Status <> 'Deceased'
                                AND $($SQL.PaycomData).Employee_Status <> 'Terminated'
                                AND $($SQL.PaycomData).Employee_Email = $($SQL.ADUsers).Email
                                AND $($SQL.PaycomData).Department <> $($SQL.ADUsers).Department"


    $sqlCommand.CommandText = $sqlQuery_compare_Department
    $sqlResult = $sqlCommand.ExecuteReader()

    $sqlResult_update_Department = New-Object System.Data.DataTable
    $sqlResult_update_Department.Load($sqlResult)
    
    foreach($Employee in $sqlResult_update_Department)
    {

        try
        {
            Get-ADUser -Filter "EmailAddress -eq '$($Employee.Employee_Email)'" | 
            Set-ADUser -Department $Employee.Department

            $Total_Updates += 1

            $html_allUpdateRecords += BuildReport -FirstName $Employee.FirstName `
                                                         -LastName $Employee.LastName `
                                                         -Email $Employee.Employee_Email `
                                                         -Attribute "Department" `
                                                         -Value $Employee.Department

            $EventMessage = "The Department for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)) was updated to $($Employee.Department)"
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 2004 -EntryType "Information" -Message $EventMessage
        }
        catch
        {
            $EventMessage = "There was an issue updating the Department for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)): " + $_.Exception.Message
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 9204 -EntryType "Error" -Message $EventMessage
            Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
        }

        
    }


    
    #===========================================================================================================
    # Check Company
    #===========================================================================================================


    $sqlQuery_compare_Company = "SELECT $($SQL.PaycomData).Employee_Email, 
                                           $($SQL.PaycomData).Company,
                                           $($SQL.ADUsers).FirstName,
                                           $($SQL.ADUsers).LastName
                                FROM $($SQL.PaycomData), $($SQL.ADUsers)
                                WHERE $($SQL.PaycomData).Employee_Status <> 'Retired'
                                AND $($SQL.PaycomData).Employee_Status <> 'Deceased'
                                AND $($SQL.PaycomData).Employee_Status <> 'Terminated'
                                AND $($SQL.PaycomData).Employee_Email = $($SQL.ADUsers).Email
                                AND $($SQL.PaycomData).Company <> $($SQL.ADUsers).Company"


    $sqlCommand.CommandText = $sqlQuery_compare_Company
    $sqlResult = $sqlCommand.ExecuteReader()

    $sqlResult_update_Company = New-Object System.Data.DataTable
    $sqlResult_update_Company.Load($sqlResult)
    
    foreach($Employee in $sqlResult_update_Company)
    {

        try
        {
            Get-ADUser -Filter "EmailAddress -eq '$($Employee.Employee_Email)'" | 
            Set-ADUser -Company $Employee.Company

            $Total_Updates += 1

            $html_allUpdateRecords += BuildReport -FirstName $Employee.FirstName `
                                                         -LastName $Employee.LastName `
                                                         -Email $Employee.Employee_Email `
                                                         -Attribute "Company" `
                                                         -Value $Employee.Company

        $EventMessage = "The Company for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)) was updated to $($Employee.Company)"
        Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 2005 -EntryType "Information" -Message $EventMessage                                          
        }
        catch
        {
            $EventMessage = "There was an issue updating the Company for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)): " + $_.Exception.Message
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 9205 -EntryType "Error" -Message $EventMessage
            Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
        }

        
    }
    

    #===========================================================================================================
    # Check City
    #===========================================================================================================
    

    $sqlQuery_compare_City = "SELECT $($SQL.PaycomData).Employee_Email, 
                                     $($SQL.PaycomData).City,
                                     $($SQL.ADUsers).FirstName,
                                     $($SQL.ADUsers).LastName
                                FROM $($SQL.PaycomData), $($SQL.ADUsers)
                                WHERE $($SQL.PaycomData).Employee_Status <> 'Retired'
                                AND $($SQL.PaycomData).Employee_Status <> 'Deceased'
                                AND $($SQL.PaycomData).Employee_Status <> 'Terminated'
                                AND $($SQL.PaycomData).Employee_Email = $($SQL.ADUsers).Email
                                AND $($SQL.PaycomData).City <> $($SQL.ADUsers).City"


    $sqlCommand.CommandText = $sqlQuery_compare_City
    $sqlResult = $sqlCommand.ExecuteReader()

    $sqlResult_update_City = New-Object System.Data.DataTable
    $sqlResult_update_City.Load($sqlResult)
    
    foreach($Employee in $sqlResult_update_City)
    {

        try
        {
            Get-ADUser -Filter "EmailAddress -eq '$($Employee.Employee_Email)'" | 
            Set-ADUser -City $Employee.City

            $Total_Updates += 1

            $html_allUpdateRecords += BuildReport -FirstName $Employee.FirstName `
                                                         -LastName $Employee.LastName `
                                                         -Email $Employee.Employee_Email `
                                                         -Attribute "City" `
                                                         -Value $Employee.City
        
            $EventMessage = "The City for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)) was updated to $($Employee.City)"
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 2006 -EntryType "Information" -Message $EventMessage 
        }
        catch
        {
            $EventMessage = "There was an issue updating the City for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)): " + $_.Exception.Message
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 9206 -EntryType "Error" -Message $EventMessage
            Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
        }

        
    }
    

    #===========================================================================================================
    # Check State
    #===========================================================================================================
    

    $sqlQuery_compare_State = "SELECT $($SQL.PaycomData).Employee_Email, 
                                      $($SQL.PaycomData).State,
                                      $($SQL.ADUsers).FirstName,
                                      $($SQL.ADUsers).LastName                                     
                                FROM $($SQL.PaycomData), $($SQL.ADUsers)
                                WHERE $($SQL.PaycomData).Employee_Status <> 'Retired'
                                AND $($SQL.PaycomData).Employee_Status <> 'Deceased'
                                AND $($SQL.PaycomData).Employee_Status <> 'Terminated'
                                AND $($SQL.PaycomData).Employee_Email = $($SQL.ADUsers).Email
                                AND $($SQL.PaycomData).State <> $($SQL.ADUsers).State"

    $sqlCommand.CommandText = $sqlQuery_compare_State
    $sqlResult = $sqlCommand.ExecuteReader()

    $sqlResult_update_State = New-Object System.Data.DataTable
    $sqlResult_update_State.Load($sqlResult)
    
    foreach($Employee in $sqlResult_update_State)
    {

        try
        {
            Get-ADUser -Filter "EmailAddress -eq '$($Employee.Employee_Email)'" | 
            Set-ADUser -State $Employee.State

            $Total_Updates += 1

            $html_allUpdateRecords += BuildReport -FirstName $Employee.FirstName `
                                                         -LastName $Employee.LastName `
                                                         -Email $Employee.Employee_Email `
                                                         -Attribute "State" `
                                                         -Value $Employee.State

            $EventMessage = "The State for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)) was updated to $($Employee.State)"
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 2007 -EntryType "Information" -Message $EventMessage 
        }
        catch
        {
            $EventMessage = "There was an issue updating the State for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)): " + $_.Exception.Message
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 9207 -EntryType "Error" -Message $EventMessage
            Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
        }


    }
    

    #===========================================================================================================
    # Check Manager
    #===========================================================================================================


    $sqlQuery_compare_Manager = "SELECT $($SQL.PaycomData).Employee_Email, 
                                        $($SQL.PaycomData).Supervisor_Email,
                                        $($SQL.ADUsers).FirstName,
                                        $($SQL.ADUsers).LastName
                                FROM $($SQL.PaycomData), $($SQL.ADUsers)
                                WHERE $($SQL.PaycomData).Employee_Status <> 'Retired'
                                AND $($SQL.PaycomData).Employee_Status <> 'Deceased'
                                AND $($SQL.PaycomData).Employee_Status <> 'Terminated'
                                AND $($SQL.PaycomData).Employee_Email = $($SQL.ADUsers).Email
                                AND $($SQL.PaycomData).Supervisor_Email <> $($SQL.ADUsers).ManagerEmail"


    $sqlCommand.CommandText = $sqlQuery_compare_Manager
    $sqlResult = $sqlCommand.ExecuteReader()

    $sqlResult_update_Manager = New-Object System.Data.DataTable
    $sqlResult_update_Manager.Load($sqlResult)
    
    foreach($Employee in $sqlResult_update_Manager)
    {

        $employee_SAM = $Employee.Employee_Email.Split("@")[0]
        $manager_SAM = $Employee.Supervisor_Email.Split("@")[0]

        try
        {
            Get-ADUser $employee_SAM | Set-ADUser -Manager $manager_SAM

            $Total_Updates += 1

            $ManagerProfile = Get-ADUser -Identity $manager_SAM

            $html_allUpdateRecords += BuildReport -FirstName $Employee.FirstName `
                                                         -LastName $Employee.LastName `
                                                         -Email $Employee.Employee_Email `
                                                         -Attribute "Manager" `
                                                         -Value "$($ManagerProfile.Name) ($($Employee.Supervisor_Email))"

            $EventMessage = "The Manager for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)) was updated to $($Employee.Supervisor_Email)"
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 2008 -EntryType "Information" -Message $EventMessage 
        }
        catch
        {
            $EventMessage = "There was an issue updating the Manager for $($Employee.FirstName) $($Employee.LastName) ($($Employee.Employee_Email)): " + $_.Exception.Message
            Write-EventLog -LogName "AR User Update" -Source "User Update" -EventId 9208 -EntryType "Error" -Message $EventMessage
            Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
        }

    }





############################################################################################################
# Close connection to SQL server
############################################################################################################


    $conn.Close()
}




############################################################################################################
# End Script
############################################################################################################

$Script_Timer.Stop()
$ExecutionTime = $Script_Timer.Elapsed
$global:ExeTime = $ExecutionTime.Hours.ToString() + "h " + $ExecutionTime.Minutes.ToString() + "m " + $ExecutionTime.Seconds.ToString() + "s " + $ExecutionTime.Milliseconds.ToString() + "ms"

FinalizeReport -ReportData $html_allUpdateRecords -NameChange $NameChange

$EventMessage = "AR User Update completed in $ExeTime. `r$Total_Updates update(s) were made."
Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 2009 -EntryType "Information" -Message $EventMessage

Exit











