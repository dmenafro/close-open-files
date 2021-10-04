<# The purpose of this script is to do the following:
    1. Run through a CSV file which contains server names, and share names
    2. Get the file path of the share
    3. Convert the file path to a usable format
    4. Check for files open under the path, and close them

    Import CSV
    Headers for the CSV should be the following
    Server,Share

    Server: Should contain just the server name

    Share: Should contain just the name of the share
#>


# Getting the desktop path of the user launching the script. Just as a default starting path
$DesktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)

# Prompt the user to select the CSV file for the script
Function Get-FileName($initialDirectory){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $initialDirectory
    $OpenFileDialog.Filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.Title = "Select GET SHARE ACCCESS CSV"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}

$FilePath  = Get-FileName -initialDirectory $DesktopPath
$csv      = @()
$csv      = Import-Csv -Path $FilePath 
$results = @()

# Set count for activity status
$i = 0

#Loop through all items in the CSV 
ForEach ($item In $csv) 
{

    # Put the objects into string variables, because new-dfsnfoldertarget likes strings
    [string]$server = $item.server
    [string]$share = $item.share
    
    # This little diddy will provide a progress bar!
    $i = $i+1
    Write-Progress -Activity "Checking \\$server\$share. Ignore any errors you see on the screen" -Status "Progress:" -PercentComplete ($i/$csv.Count*100)
  
    # Doing the work
    # Get's the file path of the share from the server
    $fp = Get-SmbShare -CimSession $server -Name $share | Select-Object -Property path

    # convert the file path because powershell fails on single slash paths in this situation
    $efp = [regex]::Escape($fp.path)

    # Find files open under the share path and close them!
    Get-SmbOpenFile -CimSession $server | Where-Object -Property Path -match $efp | Close-SmbOpenFile -Force
    
}

write-host "All done!" -BackgroundColor Green
