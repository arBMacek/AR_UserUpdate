

############################################################################################################
# Function for retrieving XML values for variables
############################################################################################################

Function Get-Settings ($Type)
{
    
    if(Test-Path -Path "$ScriptDirectory\UserUpdate_Settings.xml")
    {
        [xml]$XmlDoc = Get-Content -Path "$ScriptDirectory\UserUpdate_Settings.xml"
    }
    else
    {
        $EventMessage = "Couldn't find settings file." + $_.Exception.Message
        Write-EventLog -LogName "AR User Update" -Source "Script Environment" -EventId 9105 -EntryType "Error" -Message $EventMessage
        Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
    }

    
    if($Type -eq "SQL")
    {
        $SQL = @{
            Server = $XmlDoc.Settings.SQL.Server
            Cred = Get-SavedCredential -UserName $XmlDoc.Settings.SQL.UserName -KeyPath $XmlDoc.Settings.Environment.CredPath
            Database = $XmlDoc.Settings.SQL.Database
            ADUsers = $XmlDoc.Settings.SQL.Tables.ADUsers
            PaycomData = $XmlDoc.Settings.SQL.Tables.PaycomData
            ImportJobName = $XmlDoc.Settings.SQL.ImportJob.JobName          
        }

        Return $SQL
    }

    
    if($Type -eq "AD")
    {
        $AD = @{
            SearchBase = $XmlDoc.Settings.ActiveDirectory.SearchBase
        }

        Return $AD
    }


    if($Type -eq "Email")
    {
        $Email = @{
            SMTP_Server = $XmlDoc.Settings.Email.SMTP_Server
            From_Address = $XmlDoc.Settings.Email.From_Address
            To_Address_Error = $XmlDoc.Settings.Email.To_Address.Error
            To_Address_NameUpdate = $XmlDoc.Settings.Email.To_Address.NameUpdate
            To_Address_Report = $XmlDoc.Settings.Email.To_Address.Report
        }

        Return $Email
    }


    if($Type -eq "SFTP")
    {
        $SFTP = @{
            Hostname = $XmlDoc.Settings.Paycom_SFTP.Hostname
            Creds = Get-SavedCredential -UserName $XmlDoc.Settings.Paycom_SFTP.Username -KeyPath $XmlDoc.Settings.Environment.CredPath
            SSHKey = $XmlDoc.Settings.Paycom_SFTP.SSHKey
            SFTP_ExportFolder = $XmlDoc.Settings.Paycom_SFTP.ExportFolder
            Local_DataFolder = $XmlDoc.Settings.Environment.LocalDataFolder
            SQL_ImportFilePath = "\\" + $XmlDoc.Settings.SQL.Server + $XmlDoc.Settings.SQL.ImportFilePath
            WinSCP_Path = $XmlDoc.Settings.Environment.Includes.WinSCP
        }

        Return $SFTP
    }


}


############################################################################################################
# Function for downloading employee data from Paycom to the SQL server
############################################################################################################

function Import-PaycomSFTP-File()
{

    $SFTP = Get-Settings -Type 'SFTP'

    # Load WinSCP .NET assembly
    Add-Type -Path $SFTP.WinSCP_Path
    
    
    # Set up session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $SFTP.Hostname
        UserName = $SFTP.Creds.Username
        Password = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SFTP.Creds.Password))
        SshHostKeyFingerprint = $SFTP.SSHKey
    }
    
    $session = New-Object WinSCP.Session
    
    try
    {
        # Connect
        $session.Open($sessionOptions)
    
         # Download files
            $transferOptions = New-Object WinSCP.TransferOptions
            $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
     
            $remotePath = $SFTP.SFTP_ExportFolder
            $localPath = $SFTP.Local_DataFolder
    
    
            # Get list of files in the directory
            $directoryInfo = $session.ListDirectory($remotePath)
    
            # Select the most recent file
            $latest = $directoryInfo.Files |
                        Where-Object { -Not $_.IsDirectory } |
                        Sort-Object LastWriteTime -Descending |
                        Select-Object -First 1
    
    
            # Any file at all?
            if ($latest -eq $Null)
            {
                $EventMessage = "No files were found on the Paycom SFTP server."
                Write-EventLog -LogName "AR User Update" -Source "File Transfer" -EventId 9104 -EntryType "Error" -Message $EventMessage
                Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage

                exit 1
            }
     
            # Download the selected file
            $session.GetFileToDirectory($latest.FullName, $localPath) | Out-Null
         
    }
    finally
    {
        $session.Dispose()
    }
    
    
    $NewFile = Split-Path -Leaf $latest.FullName
    $PaycomFileSource = $localPath + "\" + $NewFile
    $PaycomFileDestination = $SFTP.SQL_ImportFilePath
    

    try 
    {
        Move-Item -Force -Path $PaycomFileSource -Destination $PaycomFileDestination

        $EventMessage = "Paycom employee file '$($NewFile)' was downloaded. The file was created on $($latest.LastWriteTime)"
        Write-EventLog -LogName "AR User Update" -Source "File Transfer" -EventId 1003 -EntryType "Information" -Message $EventMessage
    }
    catch 
    {
        $EventMessage = "There was an issue downloading the Paycom file '$($NewFile)' : " + $_.Exception.Message
        Write-EventLog -LogName "AR User Update" -Source "File Transfer" -EventId 9103 -EntryType "Error" -Message $EventMessage
        Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
    }

    

    $DownloadedFile = @{
        Name = $NewFile
        Date = $latest.LastWriteTime
    }

    Return $DownloadedFile

}



############################################################################################################
# Define function for sending email
############################################################################################################

function Send-Email ($type, $subject, $body)
{
    $Email = Get-Settings -Type "Email"

    switch ($type) {
        "Error" { 
                    $To = $Email.To_Address_Error
                }
        "NameUpdate" { 
                        $To = $Email.To_Address_NameUpdate
                    }
        "Report" { 
                    $To = $Email.To_Address_Report
                }        
        Default {}
    }

    Send-MailMessage -SmtpServer $($Email.SMTP_Server) `
    -To $To `
    -From $($Email.From_Address) `
    -Subject $subject `
    -Body $body `
    -Encoding "UTF8" `
    -BodyAsHtml

}


############################################################################################################
# Function for retrieving encrypted credentials
############################################################################################################

Function Get-SavedCredential([string]$UserName,[string]$KeyPath)
{
    If(Test-Path "$($KeyPath)\$($Username).cred") {
        $SecureString = Get-Content "$($KeyPath)\$($Username).cred" | ConvertTo-SecureString
        $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $SecureString
    }
    Else {
        $EventMessage = "Unable to retrieve credential for ""$($Username)"": " + $_.Exception.Message
        New-LogEvent -Event_ID "9011" -Type "Error" -Message $EventMessage
        Send-email -type "Error" -subject "AR User Update - Script Error" -body $EventMessage
    }
    Return $Credential
}



############################################################################################################
# Create function for adding items to an HTML email report
############################################################################################################

function BuildReport ($FirstName, $LastName, $Email, $Attribute, $Value)
{

    $html_UpdateRecord = '<div style="padding-bottom: 10px">
                            <table style="background-color: #f9f9f9; padding: 5px; width: 100%; border-bottom-style: solid; border-bottom-width: 1px; border-bottom-color: #dddddd;">
                                <tr>
                                    <td style="color: #0762c8; font-size: 17px;">
                                        <b>' + $FirstName + " " + $LastName +'</b> (' + $Email + ')</b>
                                    </td>
                                </tr>
                            </table>
                            <table style="font-size: 15px; padding: 10px;">                                
                                <tr>
                                    <td width="100">' + $Attribute + '</td>
                                    <td>' + $Value + '</td>
                                </tr>
                            </table>
                        </div>'

    return $html_UpdateRecord

}

function FinalizeReport($ReportData, $NameChange)
{

    if ($NameChange -eq $true)
    {
        $nameChangeText = '<tr>
                            <td style="font-size: 16px">
                                However, one or more employee names needs to be manually updated in AD. A separate email about each change will be sent to itsupport@alliedreliability.com.
                            </td>
                           </tr>' 
    }
    else
    {
        $nameChangeText = ""
    }

        $html_Report = Get-HTML -Block "email_start"

        if($Total_Updates -gt 0)
        {
            $html_Report += '<tr>
                            <td style="font-size: 16px; padding-top: 26px; padding-bottom: 30px;">
                                    The AR User Update script has synced all AD users with their Paycom employee records. 
                                    <br />
                                    <br />
                                    A total of <b>' + $Total_Updates + '</b> AD user profile attributes have been updated:
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <table style="width: 100%; padding: 15px; border-width: 1px; border-style: solid; border-color: #cccccc;">
                                    <tr>
                                        <td>' + $ReportData + '</td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        ' + $nameChangeText
           
        }
        else
        {

            
            $html_Report += '<tr>
                                <td style="font-size: 16px; padding-top: 26px; padding-bottom: 30px;">
                                        The AR User Update script has verified that the following attributes are synced across Paycom and AD:
                                        <ul>
                                            <li>Job Title</li>
                                            <li>Department</li>
                                            <li>Company</li>
                                            <li>City</li>
                                            <li>State</li>
                                            <li>Manager</li>
                                        </ul> 
                                        <br />
                                        No updates were necessary.
                                </td>
                            </tr>
                            ' + $nameChangeText
        }

        $html_Report += Get-HTML -Block "email_end"

        Send-Email -type "Report" -subject "User Update Report" -body $html_Report
    
}

############################################################################################################
# Function for creating the HTML for the employee name update email
############################################################################################################

function Notify_NameChange ($old_FirstName, $old_LastName, $new_FirstName, $new_LastName, $old_Email)
{
    $html_NameChange = '<html>
    <head></head>
    <body>
        <br />
        <table style="width: 650px;">
            <tr>
                <td>
                    <table style="border-collapse: collapse; width: 650px;">
                        <tr>
                            <td style="text-align: left; padding: 0; margin: 0; width: 500px;">
                                <span style="color: #21af3e; font-size: 28px;">
                                    <b>AR User Name Change Notification</b>
                                </span>
                            </td>
                            <td style="text-align: right; padding: 0; margin: 0;">

                                <img width="48px" heigth="48px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAABYlAAAWJQFJUiTwAAAGx2lUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNS42LWMxNDggNzkuMTY0MDM2LCAyMDE5LzA4LzEzLTAxOjA2OjU3ICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6cGhvdG9zaG9wPSJodHRwOi8vbnMuYWRvYmUuY29tL3Bob3Rvc2hvcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RFdnQ9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZUV2ZW50IyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ0MgMjAxOCAoV2luZG93cykiIHhtcDpDcmVhdGVEYXRlPSIyMDE4LTExLTA4VDExOjU3OjQ1LTA2OjAwIiB4bXA6TW9kaWZ5RGF0ZT0iMjAyMC0xMC0xOFQxODoxMDo0OS0wNTowMCIgeG1wOk1ldGFkYXRhRGF0ZT0iMjAyMC0xMC0xOFQxODoxMDo0OS0wNTowMCIgZGM6Zm9ybWF0PSJpbWFnZS9wbmciIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiIHBob3Rvc2hvcDpJQ0NQcm9maWxlPSJzUkdCIElFQzYxOTY2LTIuMSIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDozZDdmYWQ0My02MzQ4LTFmNDctYWY3Ni0wMDExZTdmMmFiMzAiIHhtcE1NOkRvY3VtZW50SUQ9ImFkb2JlOmRvY2lkOnBob3Rvc2hvcDplY2E5YzViYi04M2Y5LTc2NGYtYjk2Yi03OTBkNzhjM2I2ZDMiIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDo5ZmZjMjg3ZC0zNjIxLTcyNGMtOTE2Yi0zOGEyZmRiZjc1MTEiPiA8eG1wTU06SGlzdG9yeT4gPHJkZjpTZXE+IDxyZGY6bGkgc3RFdnQ6YWN0aW9uPSJjcmVhdGVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjlmZmMyODdkLTM2MjEtNzI0Yy05MTZiLTM4YTJmZGJmNzUxMSIgc3RFdnQ6d2hlbj0iMjAxOC0xMS0wOFQxMTo1Nzo0NS0wNjowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIENDIDIwMTggKFdpbmRvd3MpIi8+IDxyZGY6bGkgc3RFdnQ6YWN0aW9uPSJzYXZlZCIgc3RFdnQ6aW5zdGFuY2VJRD0ieG1wLmlpZDphZDlmMzEwMi0zNmE0LTllNGItYjExMi04NDVlNWM2ZTgyNzQiIHN0RXZ0OndoZW49IjIwMTgtMTEtMDhUMTI6MTM6NTAtMDY6MDAiIHN0RXZ0OnNvZnR3YXJlQWdlbnQ9IkFkb2JlIFBob3Rvc2hvcCBDQyAyMDE4IChXaW5kb3dzKSIgc3RFdnQ6Y2hhbmdlZD0iLyIvPiA8cmRmOmxpIHN0RXZ0OmFjdGlvbj0ic2F2ZWQiIHN0RXZ0Omluc3RhbmNlSUQ9InhtcC5paWQ6M2Q3ZmFkNDMtNjM0OC0xZjQ3LWFmNzYtMDAxMWU3ZjJhYjMwIiBzdEV2dDp3aGVuPSIyMDIwLTEwLTE4VDE4OjEwOjQ5LTA1OjAwIiBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZG9iZSBQaG90b3Nob3AgMjEuMCAoV2luZG93cykiIHN0RXZ0OmNoYW5nZWQ9Ii8iLz4gPC9yZGY6U2VxPiA8L3htcE1NOkhpc3Rvcnk+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiA8P3hwYWNrZXQgZW5kPSJyIj8+AEypYwAABPxJREFUaN7VmFtMHFUYx78ze4ZlLwPsBbog7NIEE02M+uCjjyYm3hJtG1HTC8aabkqbVkNrNL5UTbTRIja+2djWUC1oYwQTaKhaheiDbyomNo2mNYEtdpG6LLAzc47fDAUXgWXOGXa3zOR74bLn9zvnf25Lvc/+AC6eSqxfsSJO/4EoBOZmTOBjBHYng8bddw3BnaN3UE4pZJQsKPgKPP9QcPcQrDiWIgR/dRqO7LsXXt0ahfsuvQK1wZ/g9enn7b8RlNDcCnCsNFZUBL4jeTtDeIuS+W9U84Hqi56YN8zb01uJoITrEQDRnrfgj7YmyLw9B5+iQlSvhXPaRetnwhK0FPC5WYT/M5sPT/4bQg4qp1BrhKUkaNF73oL/YxpeTLYsg/+/xG2G+EjQoscGe/7gnhb29tPNK8K7laBFz/wejM1TiYLwbiRoiSYscb6siUnQWwleRoLeavCiErQUS+X6SRC+N72FWL/JKDO2BF3PpdJabdYLfiWJz7SviYIS9khgCxky404gf6l8AeHfeaZ5XeFXHolv7DglJx8ngO3TuVlFFh9ANxmksrwj2QIyPW91owchOHcm4eUqxIwI9GpfEQYM9k5u4bQxrEvhM2xVT+u5NoR/qzUB0j3vnwIlmwPOMBwKK9wmSlSgRKNRB59r32GYcATefzAt1W4ulwNfpW//ww8lIrLROHnpWzZyoQma7x8Hj5EFY8a/psRinPQoDAV+JNSvcqnGKWcH/V5+jDMDGxWfSh+ne3nb768RuPAhXPv7Hoi3Hrd/7lTCg281CwDN6lIjvy+nK8d4joBhclNV8NMEnu7xL9iuy4eURHAzuV6X5anBJ4D4p3jTY6eJiIQ1h2RWoXas9xZjKfhY8DtGD5NEVT1UmzXwF8VNKZKC8fP2JiUkIbORHcDqlM18d6rPhm/y1pMQ9YJumIiMk7diFtRARkqClgr+k9SXbPvPh0i8EuHVajD5dF4ecE9VUSJyTViClgp+xy/Y85Wxm/Dm8oOxpAQtemww89vt2MRIRK0BA+FXn5niErQUmY9752NTEL6QxKMfEY6bljnnA0KYY4F295nvwMw3LMbG+eEnT2JwG0FoeyTm0nVgzvqWjMRqAvuxumThz6T62a7RlzDz9eLwyyRSMHZ+mz0SjY90E30qvCROtKyZdyDhwSUWQhMwNvAkxkhZNieoW3ied5Rc2KSEMr/m5+NIWPtEdOXNjrqBNzmDCo+6BL7J6yI2EqvTgsAbWC8LHafxDXh8OMaE9U4OgiS8Ii2BqxJK2F/u7rwJn7FOyU5vSF58NysN/NTVc1ryypEKFz1/3fFdIl9iCG9khN+wBIaxQiLw1nm8kdTBGTLgffPK6d82KZFojadKeKkk1PhAUfUDwBQuOhIV4Qky8f0DzBK4LHI3pdxj3037tGF4N9ST3WRGeBUPyPT8CazdsnMC5UFV55wf5mx4PPbHjRj0BYehM3TWulBUajxATPFTtQX/nLuJjalTuDOBhdg0GFHoR/iucA9U4W1Iw2LlgBc5zOV/pdEfHIHOyFmoMgP2dc4sM/yaAvnwVua7Qj14i5rveQn4k+sNX1BgJXgXsTmF1QZFeGgJ4E8Uo+dXFdhI8MsENhr8EoGNCL8osFHh578hdA/PygVvC7iEtw5AvnLB2wL1RsTeYSVjY8FrWH3lgLcF+rUROB76VDbzMayjWIehTM+/EH9Svy5vEVsAAAAASUVORK5CYII=
">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="font-size: 16px; padding-top: 26px; padding-bottom: 30px;">
                The AR User Update script has detected a difference between an employee&#39;s name in Paycom and AD.
                <br />
                <br />
                Pleae update the users name in AD and complete all related maintenance tasks (listed at the bottom).
                </td>
            </tr>
            <tr>
                <td>
                    <table style="width: 100%; padding: 15px; border-width: 1px; border-style: solid; border-color: #cccccc;">
                        <tr>
                            <td>


                                <div style="padding-bottom: 10px">
                                    <table style="background-color: #f9f9f9; padding: 5px; width: 100%; border-bottom-style: solid; border-bottom-width: 1px; border-bottom-color: #dddddd;">
                                        <tr>
                                            <td style="color: #0762c8; font-size: 17px;">
                                                <b>' + $old_FirstName + ' ' + $old_LastName + '</b> (' + $old_Email + ')</b>
                                            </td>
                                        </tr>
                                    </table>
                                    <table style="font-size: 15px; padding: 10px;">
                                        <tr>
                                            <td width="100">New Name:</td>
                                            <td>' + $new_FirstName + ' ' + $new_LastName + '</td>
                                        </tr>
                                    </table>
                                </div>


                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="font-size: 16px; padding-top: 26px;">
                    <b>Name Change Maintenance:</b>
                    <ul>
                        <li>Update user SAM</li>
                        <li>Update user UPN</li>
                        <li>Update user email</li>
                        <li>Update user proxy addresses to include old email</li>
                        <li>Delete user&#39;s Office activations in Office 365 admin</li>
                        <li>Assist user with signing into Office apps with new UPN</li>
                        <li>Notify HR of user&#39;s new email address</li>
                    </ul>
                </td>
            </tr>
            <tr>
                <td>
                    <br />
                    <table style="background-color: #f9f9f9; width: 100%; padding: 5px;">
                        <tr>
                            <td style="width: 33%">AD users evaluated: ' + $Total_ADUsers + '</td>
                            <td style="width: 33%; text-align: center;">Paycom records evaluated: ' + $Total_PaycomData + '</td>
                            <td style="width: 33%; text-align: right;">Execution time: ' + $ExeTime + '</td>
                        </tr>

                        <tr>
                            <td colspan="3">This email was generated by the PowerShell script UserUpdate_Lib.ps1 on HDC-MIS02.ar.local.</td>
                        </tr>
                       
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <br />
                    <br />
                </td>
            </tr>
        </table>
        <br />
        <br />

    </body>
</html>'

Send-Email -type "NameUpdate" -subject "Name Change Notification: $old_FirstName $old_LastName" -body $html_NameChange

}


############################################################################################################
# Function for providing reusable blocks of HTML
############################################################################################################

function Get-HTML($Block)
{
    $HTML_Block = ''

    switch ($Block)
    {
        "email_start" {
            $HTML_Block = '<html>
                            <head></head>
                            <body>
                                <br />
                                <table style="width: 650px;">
                                    <tr>
                                        <td>
                                            <table style="border-collapse: collapse; width: 650px;">
                                                <tr>
                                                    <td style="text-align: left; padding: 0; margin: 0; width: 500px;">
                                                        <span style="color: #21af3e; font-size: 28px;">
                                                            <b>AR User Update Report</b>
                                                        </span>
                                                    </td>
                                                    <td style="text-align: right; padding: 0; margin: 0; height: 48px"> &nbsp;
                        
                                                        <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAABYlAAAWJQFJUiTwAAAGx2lUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNS42LWMxNDggNzkuMTY0MDM2LCAyMDE5LzA4LzEzLTAxOjA2OjU3ICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6cGhvdG9zaG9wPSJodHRwOi8vbnMuYWRvYmUuY29tL3Bob3Rvc2hvcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RFdnQ9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZUV2ZW50IyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ0MgMjAxOCAoV2luZG93cykiIHhtcDpDcmVhdGVEYXRlPSIyMDE4LTExLTA4VDExOjU3OjQ1LTA2OjAwIiB4bXA6TW9kaWZ5RGF0ZT0iMjAyMC0xMC0xOFQxODoxMDo0OS0wNTowMCIgeG1wOk1ldGFkYXRhRGF0ZT0iMjAyMC0xMC0xOFQxODoxMDo0OS0wNTowMCIgZGM6Zm9ybWF0PSJpbWFnZS9wbmciIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiIHBob3Rvc2hvcDpJQ0NQcm9maWxlPSJzUkdCIElFQzYxOTY2LTIuMSIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDozZDdmYWQ0My02MzQ4LTFmNDctYWY3Ni0wMDExZTdmMmFiMzAiIHhtcE1NOkRvY3VtZW50SUQ9ImFkb2JlOmRvY2lkOnBob3Rvc2hvcDplY2E5YzViYi04M2Y5LTc2NGYtYjk2Yi03OTBkNzhjM2I2ZDMiIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDo5ZmZjMjg3ZC0zNjIxLTcyNGMtOTE2Yi0zOGEyZmRiZjc1MTEiPiA8eG1wTU06SGlzdG9yeT4gPHJkZjpTZXE+IDxyZGY6bGkgc3RFdnQ6YWN0aW9uPSJjcmVhdGVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjlmZmMyODdkLTM2MjEtNzI0Yy05MTZiLTM4YTJmZGJmNzUxMSIgc3RFdnQ6d2hlbj0iMjAxOC0xMS0wOFQxMTo1Nzo0NS0wNjowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIENDIDIwMTggKFdpbmRvd3MpIi8+IDxyZGY6bGkgc3RFdnQ6YWN0aW9uPSJzYXZlZCIgc3RFdnQ6aW5zdGFuY2VJRD0ieG1wLmlpZDphZDlmMzEwMi0zNmE0LTllNGItYjExMi04NDVlNWM2ZTgyNzQiIHN0RXZ0OndoZW49IjIwMTgtMTEtMDhUMTI6MTM6NTAtMDY6MDAiIHN0RXZ0OnNvZnR3YXJlQWdlbnQ9IkFkb2JlIFBob3Rvc2hvcCBDQyAyMDE4IChXaW5kb3dzKSIgc3RFdnQ6Y2hhbmdlZD0iLyIvPiA8cmRmOmxpIHN0RXZ0OmFjdGlvbj0ic2F2ZWQiIHN0RXZ0Omluc3RhbmNlSUQ9InhtcC5paWQ6M2Q3ZmFkNDMtNjM0OC0xZjQ3LWFmNzYtMDAxMWU3ZjJhYjMwIiBzdEV2dDp3aGVuPSIyMDIwLTEwLTE4VDE4OjEwOjQ5LTA1OjAwIiBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZG9iZSBQaG90b3Nob3AgMjEuMCAoV2luZG93cykiIHN0RXZ0OmNoYW5nZWQ9Ii8iLz4gPC9yZGY6U2VxPiA8L3htcE1NOkhpc3Rvcnk+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiA8P3hwYWNrZXQgZW5kPSJyIj8+AEypYwAABPxJREFUaN7VmFtMHFUYx78ze4ZlLwPsBbog7NIEE02M+uCjjyYm3hJtG1HTC8aabkqbVkNrNL5UTbTRIja+2djWUC1oYwQTaKhaheiDbyomNo2mNYEtdpG6LLAzc47fDAUXgWXOGXa3zOR74bLn9zvnf25Lvc/+AC6eSqxfsSJO/4EoBOZmTOBjBHYng8bddw3BnaN3UE4pZJQsKPgKPP9QcPcQrDiWIgR/dRqO7LsXXt0ahfsuvQK1wZ/g9enn7b8RlNDcCnCsNFZUBL4jeTtDeIuS+W9U84Hqi56YN8zb01uJoITrEQDRnrfgj7YmyLw9B5+iQlSvhXPaRetnwhK0FPC5WYT/M5sPT/4bQg4qp1BrhKUkaNF73oL/YxpeTLYsg/+/xG2G+EjQoscGe/7gnhb29tPNK8K7laBFz/wejM1TiYLwbiRoiSYscb6siUnQWwleRoLeavCiErQUS+X6SRC+N72FWL/JKDO2BF3PpdJabdYLfiWJz7SviYIS9khgCxky404gf6l8AeHfeaZ5XeFXHolv7DglJx8ngO3TuVlFFh9ANxmksrwj2QIyPW91owchOHcm4eUqxIwI9GpfEQYM9k5u4bQxrEvhM2xVT+u5NoR/qzUB0j3vnwIlmwPOMBwKK9wmSlSgRKNRB59r32GYcATefzAt1W4ulwNfpW//ww8lIrLROHnpWzZyoQma7x8Hj5EFY8a/psRinPQoDAV+JNSvcqnGKWcH/V5+jDMDGxWfSh+ne3nb768RuPAhXPv7Hoi3Hrd/7lTCg281CwDN6lIjvy+nK8d4joBhclNV8NMEnu7xL9iuy4eURHAzuV6X5anBJ4D4p3jTY6eJiIQ1h2RWoXas9xZjKfhY8DtGD5NEVT1UmzXwF8VNKZKC8fP2JiUkIbORHcDqlM18d6rPhm/y1pMQ9YJumIiMk7diFtRARkqClgr+k9SXbPvPh0i8EuHVajD5dF4ecE9VUSJyTViClgp+xy/Y85Wxm/Dm8oOxpAQtemww89vt2MRIRK0BA+FXn5niErQUmY9752NTEL6QxKMfEY6bljnnA0KYY4F295nvwMw3LMbG+eEnT2JwG0FoeyTm0nVgzvqWjMRqAvuxumThz6T62a7RlzDz9eLwyyRSMHZ+mz0SjY90E30qvCROtKyZdyDhwSUWQhMwNvAkxkhZNieoW3ied5Rc2KSEMr/m5+NIWPtEdOXNjrqBNzmDCo+6BL7J6yI2EqvTgsAbWC8LHafxDXh8OMaE9U4OgiS8Ii2BqxJK2F/u7rwJn7FOyU5vSF58NysN/NTVc1ryypEKFz1/3fFdIl9iCG9khN+wBIaxQiLw1nm8kdTBGTLgffPK6d82KZFojadKeKkk1PhAUfUDwBQuOhIV4Qky8f0DzBK4LHI3pdxj3037tGF4N9ST3WRGeBUPyPT8CazdsnMC5UFV55wf5mx4PPbHjRj0BYehM3TWulBUajxATPFTtQX/nLuJjalTuDOBhdg0GFHoR/iucA9U4W1Iw2LlgBc5zOV/pdEfHIHOyFmoMgP2dc4sM/yaAvnwVua7Qj14i5rveQn4k+sNX1BgJXgXsTmF1QZFeGgJ4E8Uo+dXFdhI8MsENhr8EoGNCL8osFHh578hdA/PygVvC7iEtw5AvnLB2wL1RsTeYSVjY8FrWH3lgLcF+rUROB76VDbzMayjWIehTM+/EH9Svy5vEVsAAAAASUVORK5CYII=
                        ">
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>'
        }

        "email_end" {
            $HTML_Block = '<tr>
                            <td>
                                <br />
                                <table style="background-color: #f9f9f9; width: 100%; padding: 5px;">
                                    <tr>
                                        <td style="width: 33%">AD users evaluated: ' + $Total_ADUsers + '</td>
                                        <td style="width: 33%; text-align: center;">Paycom records evaluated: ' + $Total_PaycomData + '</td>
                                        <td style="width: 33%; text-align: right;">Execution time: ' + $ExeTime + '</td>
                                    </tr>

                                    <tr>
                                        <td colspan="3">This email was generated by the PowerShell script UserUpdate_Lib.ps1 on HDC-MIS02.ar.local.</td>
                                    </tr>
                                
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <br />
                                <br />
                            </td>
                        </tr>
                    </table>
                    <br />
                    <br />

                </body>
                </html>'
        }
    }

    return $HTML_Block

}



