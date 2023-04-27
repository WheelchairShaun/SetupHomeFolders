<#
.SYNOPSIS
    Sets up the 3D Objects, Desktop, Documents, Pictures, Music, and Videos folders to a redirected home folder for a new user profile on Windows 10 local PC

    The Home folder must already be created and set up with the following permissions:
        CREATOR OWNER - Full Control (Apply onto: Subfolders and Files Only)
        System - Full Control (Apply onto: This Folder, Subfolders and Files)
        Administrators - Full Control (Apply onto: This Folder, Subfolders and Files)
        Everyone - Create Folder/Append Data (Apply onto: This Folder Only)
        Everyone - List Folder/Read Data (Apply onto: This Folder Only)
        Everyone - Read Attributes (Apply onto: This Folder Only)
        Everyone - Traverse Folder/Execute File (Apply onto: This Folder Only)
    
    Edit the $TargetPath if you would like a Differentt location.
.EXAMPLE
    PS> ./SetupHomeFolders
#>

$TargetPath = "D:\Users\$env:UserName"

# Folders to redirect
$FoldersToRedirect = 
@(
    '3DObjects',
    'Desktop',
    'Documents',
    'Downloads',
    'Music',
    'Pictures',
    'Videos'
)

Write-Output "Attempting to move home folders to $TargetPath...`n"

$MajorVersion = $PSVersionTable.PSVersion.Major
Write-Output "Importing Module for PSVersion $MajorVersion..."
if ($MajorVersion -eq 5)
{
    Import-Module ./KnownFolderPathPS5.ps1 -Force
}
elseif ($MajorVersion -eq 7) 
{
    Import-Module ./KnownFolderPathPS7.ps1 -Force
}
else
{
    Write-Error -Message "This version of PowerShell is not supported." -Category OperationStopped
    exit
}

Write-Output "Module successfully imported!`n"

foreach($Folder in $FoldersToRedirect)
{
    Write-Output "Processing $Folder folder..."
    $TargetPathFolder = "$TargetPath\$Folder"
    $CurrentPathFolder = ""
    $output = Get-KnownFolderPath $Folder ([ref]$CurrentPathFolder)

    if ($output -ne 0)
    {
        Write-Error -Category InvalidResult "$Folder does not exist in this user profile"
        continue
    }

    if ($CurrentPathFolder -eq $TargetPathFolder)
    {
        Write-Output "The $Folder folder is already at the location: $TargetPathFolder`n"
    }
    else 
    { 
        # Validate the path
        if(!(Test-Path $TargetPathFolder -PathType Container))
        {
            Write-Output "Creating new folder: $TargetPathFolder..."
            New-Item -Path $TargetPathFolder -type Directory -Force
            Write-Output "`n"
        }

        Write-Output "Relocating $Folder from $CurrentPathFolder to $TargetPathFolder..."
        $output = Set-KnownFolderPath -KnownFolder $Folder -Path $TargetPathFolder
        
        if ($output -ne 0)
        {
            Write-Error -Category InvalidResult "$Folder was not set"
            continue
        }

        Write-Output "Moving contents of $Folder..."
        Move-Item "$CurrentPathFolder\*" $TargetPathFolder -Force
        Write-Output "Removing $CurrentPathFolder..."
        Remove-Item $CurrentPathFolder -Recurse -Force

        Write-Output "$Folder folder relocated successfully!`n"
    }
    
}

Write-Output "`nScript finished successfully!`n"