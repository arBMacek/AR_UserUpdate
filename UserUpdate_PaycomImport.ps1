

$global:ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
. ("$ScriptDirectory\UserUpdate_Lib.ps1")

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
            $latest =
                $directoryInfo.Files |
                Where-Object { -Not $_.IsDirectory } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1
    
    
            # Any file at all?
            if ($latest -eq $Null)
            {
                Write-Host "No file found"
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
    
    Move-Item -Force -Path $PaycomFileSource -Destination $PaycomFileDestination

}






