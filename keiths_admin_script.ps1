Start-Transcript "C:\Users\$env:UserName\PSLOG\kaslog.txt"
$connectedToExchange = $false
$connectedToGraph = $false
$activeDirectoryImported = $false

# IMPORTANT!!!! The first time you run this, run it as administrator! It will install necessary module
# Afterwards, not needed.

# It's now easier than ever to add and move around menu items. 
# IF you want to add new functionality to the script:
#   1. Write a function somewhere above the menu function
#   2. Add its menu entry to the MenuItems array
#   3. Add a new switch statement in the main menu function which calls the function you made



# Below code block will create a settings directory and json if one does not already exist
$settingsPath = "C:\PowerShell\Keith's Admin Script"
$settingsFile = "C:\PowerShell\Keith's Admin Script\settings.json"
$defaultSettings = @{
    firstTime = $true
    showStatus = $true
    showWelcome = $true
    domainController = "hostname-goes-here"
    exchangeDomain = "example.com"

}
if (-not (Test-Path -PathType Container $settingsPath)) {
    # Create the folder structure
    New-Item -ItemType Directory -Path $settingsPath
    Write-Host "Created folder structure: $settingsPath"
}
if (-not (Test-Path $settingsFile)) {
    # Create the file if it doesn't exist
    New-Item -Path $settingsFile -ItemType File
    $defaultSettings | ConvertTo-Json -Depth 100 | Set-Content -Path $settingsFile
    $settings = Get-Content $settingsFile -ErrorAction Stop | Out-String | ConvertFrom-Json
    Write-Host "Created new file and added initial text content"
} else {
    $settings = Get-Content $settingsFile -ErrorAction Stop | Out-String | ConvertFrom-Json
}

$MenuHeader = "Enter selection`n----------"
$SubMenuHeader = "----------"

$MenuItems = @(
    "Disable clutter on user's mailbox",
    "Block email address",
    "Add subscriber to StatusPage (Email)",
    "Add subscriber to StatusPage (Phone #)",
    "Add PowerShell Script as Windows app`n`t->Requires Powershell 7 to be installed",
    "Add anything as app (pin to start/taskbar)",
    "Enable the in-place archive for a user",
    "Adjust email group membership`n`t->Bulk add users to group, or groups to users",
    "List members of an email group`n`t->Only works for DL and mail-enabled security",
    "Pipe (|) Delimiter Filter Prep",
    "Dell service tag lookup",
    "Pipe (|) Delimiter Filter Prep (Inverse)",
    "FedEx Tracking Number Lookup"
)

$SubMenuItems  = @(
    $SubMenuHeader,
    "a. AD commands",
    "e. Connect to ExchangeOnline (needed for any exchange commands)",
    "c. Chocolatey commands",
    "l. View changelog",
    "s. Change settings"
    "q. quit"
)

$changelog = @(
    "8/17/23`n`t~Added service tag lookup`n`t~Added exchange connection status indicator to top of menu items`n`t~Added changelog to script.",
    "8/18/23`n`t~Added disclaimer to email group operations",
    "2/5/24`n`t~Added ability to bulk add users to groups and bulk add groups to users",
    "2/9/24`n`t~Added settings menu, welcome message, first time welcome message for new users, and ability to`n`t toggle that message alongside Microsoft Exchange connection indicator.",
    "2/12/24`n`t~Added ability to lookup fedex tracking information",
    "4/19/24`n`t~Removed the ability to block emails by uploading the .eml file since it was broken`n`t~Added the ability to disable the clutter feature for a user's mailbox",
    "10/24/24`n`tRemoved any ties to former employers, and added a setting to adjust the exchange domain as well as AD domain directly from the script"
)

$firstTimeWelcome = "Welcome to Keith's Admin Script. You can use option 5 on the menu to add this script as an app on your computer.`nBefore running any Exchange commands, use the 'e' option on the main menu to connect to exchange."

$welcomeMessage = "Welcome! There have been updates, see the changelog.`nYou can now disable clutter for a user's mailbox"

$s = $settings.domainController # Preferred AD server -- change this variable at any time
$d = $settings.exchangeDomain # Domain

#----------DEPENDENCY CHECKER----------#
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) { # needed for any email-related functions
    Write-Host "ExchangeOnlineManagement module is already installed."
} else {
    # Install the ExchangeOnlineManagement module
    Install-Module -Name ExchangeOnlineManagement -Force
}

if (Get-Module -ListAvailable -Name Microsoft.Graph) { # needed for any email-related functions
    Write-Host "Graph module is already installed."
} else {
    # Install the ExchangeOnlineManagement module
    Install-Module -Name Microsoft.Graph -Force
}

# Check if the ActiveDirectory module is installed
if (Get-Module -ListAvailable -Name ActiveDirectory) { # Needed for AD functions
    Write-Host "ActiveDirectory module is already installed."
} else {
    # Install the ExchangeOnlineManagement module
    Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
    # Write-Host "ActiveDirectory module not installed. Please install the RSAT: AD LDS"
}

Import-Module ExchangeOnlineManagement
#--------------------------------------#

function Save-Settings{
    $settings | ConvertTo-Json -Depth 100 | Set-Content -Path $settingsFile
}

function Disable-Clutter{
    $username = Read-Host "Please enter the username (before the @) of the user you wish to disable clutter for"
    Set-Clutter -Identity "$username@$($settings.exchangeDomain)" -Enable $false
}

function Block-SenderByEML{
    Import-Module ExchangeOnlineManagement

    # Import the Windows Forms assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Create a new open file dialog
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog

    # Set the filter to only show .eml files
    $openFileDialog.Filter = "Email files (*.eml)|*.eml"

    # Show the dialog and get the result
    $result = $openFileDialog.ShowDialog()

    # If the user selected a file, rename it to .txt and read it
    if ($result -eq "OK") {
        # Get the full path of the selected file
        $emlFile = $openFileDialog.FileName

        # Replace the .eml extension with .txt
        $txtFile = $emlFile -replace "\.eml$", ".txt"

        # Rename the file
        Rename-Item $emlFile $txtFile

        # Display a message box to confirm the operation
        # [System.Windows.Forms.MessageBox]::Show("Renamed $emlFile to $txtFile", "Success")

        # Read the text file and filter the lines that contain "From:"
        $lines = Get-Content $txtFile | Where-Object {$_ -match "@"} # originally "From:"

        # Get the last line that contains "From:"
        $lastLine = $lines[0] # originally -1

        # Get the index of the last line in the text file
        $index = [array]::IndexOf((Get-Content $txtFile), $lastLine)

        # Initialize an empty array to store the output lines
        $outputLines = @()

        # Loop through the text file from the index of the last line until the end or until a line with ">" is found
        for ($i = $index; $i -lt (Get-Content $txtFile).Count; $i++) {
            # Get the current line
            $currentLine = (Get-Content $txtFile)[$i]

            # Add the current line to the output array
            $outputLines += $currentLine

            # If the current line contains ">", break the loop
            if ($currentLine -match ">") {
                break
            }
        }

        # Display the output lines as a single string
        Write-Output ($outputLines -join "`n")

        $email = ($outputLines -join "`n")

        # Prompt the user to continue
        $continue = Read-Host -Prompt "Do you want to block this sender? (y/n)"

        # If the user answers no, exit the script
        if ($continue -ne "y") {
            Write-Host "Exiting..."
            break
        }

        $len = $email.IndexOf(">") - $email.IndexOf("<")

        $email = $email.Substring($email.IndexOf("<") + 1, $len - 1)


        # Write-Output $email

        

        # Delete the text file
        Remove-Item $txtFile

        # Display a message box to confirm the deletion
        # [System.Windows.Forms.MessageBox]::Show("Deleted $txtFile", "Success")

        New-TenantAllowBlockListItems -ListType Sender -Block -Entries $email -NoExpiration

        # Pause until a key is pressed
        Read-Host -Prompt "Press the Enter key to exit"


        # Open the URL in the default browser
        # Start-Process https://admin.exchange.microsoft.com/#/transportrules

    }
}

function Block-SenderByAddress{
    Import-Module ExchangeOnlineManagement
    # Connect-ExchangeOnline

    $email = Read-Host -Prompt "Email"

    # Prompt the user to continue
    $continue = Read-Host -Prompt "Do you want to block this sender? (y/n)"

    # If the user answers no, exit the script
    if ($continue -ne "y") {
        Write-Host "Exiting..."
        break
    }

    # $len = $email.IndexOf(">") - $email.IndexOf("<")

    # $email = $email.Substring($email.IndexOf("<") + 1, $len - 1)


    # Write-Output $email

    

    # Delete the text file
    # Remove-Item $txtFile

    # Display a message box to confirm the deletion
    # [System.Windows.Forms.MessageBox]::Show("Deleted $txtFile", "Success")

    New-TenantAllowBlockListItems -ListType Sender -Block -Entries $email -NoExpiration

    # Pause until a key is pressed
    Read-Host -Prompt "Press the Enter key to exit"
}


function Add-ScriptAsApp {
    # Add reference to System.Windows.Forms namespace
    Add-Type -AssemblyName System.Windows.Forms

    # Create an OpenFileDialog object
    $dialog = New-Object System.Windows.Forms.OpenFileDialog

    # Set some properties of the dialog
    $dialog.InitialDirectory = "C:\"
    $dialog.Filter = "All files (*.*)|*.*"
    $dialog.Title = "Select a file to add to start menu / taskbar"

    # Show the dialog and get the user's input
    if ($dialog.ShowDialog() -eq "OK")
    {
        # Assign the selected file path to $target
        $target = $dialog.FileName
    }
    else
    {
        # Exit the script if the user cancels
        Write-Host "No file selected. Exiting script."
        Exit
    }

    # Ask for what the name will show up as in the start menu / taskbar
    $name = Read-Host -Prompt "Name to show in Start Menu"

    # Defines the path to make a shortcut, then creates a shortcut there. 
    $toPath = "C:\Users\$env:UserName\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\$name.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $shortcut = $WScriptShell.CreateShortcut($toPath)
    $shortcut.TargetPath = "C:\Program Files\PowerShell\7\pwsh.exe" # Change this line to target explorer.exe
    $shortcut.Arguments = "`"$target`"" # Add this line to pass the original target as an argument
    $shortcut.Save()
}

function Add-AnythingAsApp {
    # Add reference to System.Windows.Forms namespace
    Add-Type -AssemblyName System.Windows.Forms

    # Create an OpenFileDialog object
    $dialog = New-Object System.Windows.Forms.OpenFileDialog

    # Set some properties of the dialog
    $dialog.InitialDirectory = "C:\"
    $dialog.Filter = "All files (*.*)|*.*"
    $dialog.Title = "Select a file to add to start menu / taskbar"

    # Show the dialog and get the user's input
    if ($dialog.ShowDialog() -eq "OK")
    {
        # Assign the selected file path to $target
        $target = $dialog.FileName
    }
    else
    {
        # Exit the script if the user cancels
        Write-Host "No file selected. Exiting script."
        Exit
    }

    # Ask for what the name will show up as in the start menu / taskbar
    $name = Read-Host -Prompt "Name to show in Start Menu"

    # Defines the path to make a shortcut, then creates a shortcut there. 
    $toPath = "C:\Users\$env:UserName\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\$name.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $shortcut = $WScriptShell.CreateShortcut($toPath)
    $shortcut.TargetPath = "explorer.exe" # Change this line to target explorer.exe
    $shortcut.Arguments = "`"$target`"" # Add this line to pass the original target as an argument
    $shortcut.Save()
}
function Enable-Archive {
    Import-Module ExchangeOnlineManagement
    # Connect-ExchangeOnline

    $username = Read-Host -Prompt "Username"

    Enable-Mailbox -Identity $username -Archive
    Read-Host "Press enter to continue"
}

function Add-UserToEmailGroup {
    Write-Host "Bulk add users or groups?"
    $choice = Read-Host -Prompt "Enter USER or GROUP"

    switch($choice.Substring(0,1).ToLower()){
        "u" {
            $group = Read-Host -Prompt "Group (before the @)"
            $user = ""
            
            Write-Host "Begin entering users. Only use the name before the @ of the user. Submit q to exit."
            while($user -ne "q"){
                $user = Read-Host -Prompt "User to add"
                if($user -ne "q"){
                    $groupObj = Get-UnifiedGroup -Identity "$group@$($settings.exchangeDomain)" -ErrorAction SilentlyContinue
                    if($groupObj){
                        Write-Host "This is an M365 group, updating membership..."
                        Add-UnifiedGroupLinks -Identity "$group@$($settings.exchangeDomain)" -LinkType Member -Links "$user@$($settings.exchangeDomain)"
                    } else {
                        Write-Host "This is an Exchange group, updating membership..."
                        Add-DistributionGroupMember -Identity "$group@$($settings.exchangeDomain)" -Member "$user@$($settings.exchangeDomain)" -BypassSecurityGroupManagerCheck
                    }
                }
            }
        }
        "g" {
            $group = ""
            $user = Read-Host -Prompt "Username"
            
            Write-Host "Begin entering groups. Only use the name before the @ of the group. Submit q to exit."
            while($group -ne "q"){
                $group = Read-Host -Prompt "Group to add"
                if($group -ne "q"){
                    $groupObj = Get-UnifiedGroup -Identity "$group@$($settings.exchangeDomain)" -ErrorAction SilentlyContinue
                    if($groupObj){
                        Write-Host "This is an M365 group, updating membership..."
                        Add-UnifiedGroupLinks -Identity "$group@$($settings.exchangeDomain)" -LinkType Member -Links "$user@$($settings.exchangeDomain)"
                    } else {
                        Write-Host "This is an Exchange group, updating membership..."
                        Add-DistributionGroupMember -Identity "$group@$($settings.exchangeDomain)" -Member "$user@$($settings.exchangeDomain)" -BypassSecurityGroupManagerCheck
                    }
                }
            }
        }
    }

    
}

function Get-EmailGroupMembers {
    $group = Read-Host -Prompt "Group (before the @)"

    $groupObj = Get-UnifiedGroup -Identity "$group@$($settings.exchangeDomain)" -ErrorAction SilentlyContinue
    if($groupObj){
        Write-Host "This is an M365 group, getting membership..."
        $members = Get-UnifiedGroupLinks -Identity $groupObj.Id -LinkType members
        foreach ($member in $members) {
            Write-Host $member.DisplayName
        }
    } else {
        Write-Host "This is an Exchange group, getting membership..."
        $distGroupObj = Get-DistributionGroup -Identity "$group@$($settings.exchangeDomain)"
        $members = Get-DistributionGroupMember -Identity $distGroupObj.Identity
        foreach($member in $members){
            Write-Host $member.DisplayName
        }
    }
    Read-Host -Prompt "Press enter to continue"
}


function Get-ADUserLoc {
    $username= Read-Host -Prompt "Username"
    Get-ADUser -Server $s -Identity $username -Properties Office | Select-Object -ExpandProperty Office
    Read-Host -Prompt "Press Enter to continue..."
}

function Get-ADUserManager {
    $username= Read-Host -Prompt "Username"
    $user = Get-ADUser -Server $s -Identity $username -Properties Manager
    $manager = Get-ADUser -Server $s -Identity $user.Manager
    $mgrname = $manager.Name
    Write-Output "$mgrname"
}

function Get-ADUserLastLogin {
    $username= Read-Host -Prompt "Username"
    Get-ADUser -Server $s -Identity $username -Properties * | select *logon*
}



function Convert-NewlineToPipe () {
    $array = @()
    $item = "ERROR"
    while ($item -ne "") {
        $item = Read-Host -Prompt "Enter an item (press Enter to finish)"
        if ($item -ne "") {
            $array += $item
        }
    }
    $new_string = $array -join "|"
    Write-Output $new_string
    Read-Host -Prompt "Press Enter to finish"
}

function Convert-NewlineToPipeNOT () {
    $array = @()
    $item = "ERROR"
    while ($item -ne "") {
        $item = Read-Host -Prompt "Enter an item (press Enter to finish)"
        if ($item -ne "") {
            $array += "<>$item"
        }
    }
    $new_string = $array -join "&"
    Write-Output $new_string
    Read-Host -Prompt "Press Enter to finish"
}

function Lookup-ST{
    param(
        [string]$servicetag
    )

    Start-Process "https://www.dell.com/support/home/en-us/product-support/servicetag/$servicetag/overview"
}

function Get-Changelog{
    Clear-Host
    Render-Menu -header $changelog
    Read-Host -Prompt "Press Enter to continue..."
}

function Get-FedExTracking{
    $trackingnum = Read-Host -Prompt "Enter tracking number"
    $baseURL = "https://www.bing.com/search?q=fedex+tracking+$trackingnum"
    Start-Process $baseURL
}


#----------ALL FUNCTIONS ABOVE HERE----------#

function Render-Menu{
    param(
        [array]$header,
        [array]$entries
    )


    foreach($line in $header){
        Write-Host $line
    }

    $counter = 1
    foreach($entry in $entries){
        Write-Host "$counter. $entry`n"
        $counter++
    }
}

#----------ALL MENUS BELOW HERE----------#

function ChocoMenu{
    Read-Host -Prompt "This functionality has been removed"
}

function ADmenu{
    $ADMenuHeader = @("AD Commands","----------")
    $ADMenuItems = @(
        "Get user location",
        "Get user manager",
        "Get user last login time"
    )
    $ADSubMenuItems = @(
        "a. Load AD module"
        "q. Quit"
    )
    $adcom = ""
    while($adcom -ne "q"){
        clear
        # Write-Host "`nAD Commands`n----------`n1. Get user location`n2. Get user manager`n3. Get user last login time`n----------`nq. Quit"
        Render-Menu -header $ADMenuHeader -entries $ADMenuItems
        Render-Menu -header $ADSubMenuItems
        $adcom = Read-Host -Prompt "Selection"
        switch($adcom){
            1 {
                Get-ADUserLoc
            }
            2 {
                Get-ADUserManager
            }
            3 {
                Get-ADUserLastLogin
            }
            "a" {
                Import-Module ActiveDirectory
                $activeDirectoryImported = $true
            }
        }
    }
}

function SettingsMenu{
    $SettingsMenuHeader = @("SETTINGS","----------")
    $settingsMenuItems = @(
        "Toggle connection status display`n`tCurrently: $($settings.showStatus)",
        "Show welcome message`n`tCurrently: $($settings.showWelcome)",
        "Adjust domain for email operations`n`tCurrently: $($settings.exchangeDomain)",
        "Adjust domain controller hostname for AD operations`n`tCurrently: $($settings.domainController)"
    )
    $SettingsSubMenuItems = @(
        "q. Quit"
    )
    $setchoice = ""
    $err = $false
    $errmsg = ""
    while($setchoice -ne "q"){
        clear
        if($err){
            Write-Host $errmsg -ForegroundColor Red
        }
        # Write-Host "`nAD Commands`n----------`n1. Get user location`n2. Get user manager`n3. Get user last login time`n----------`nq. Quit"
        Render-Menu -header $SettingsMenuHeader -entries $settingsMenuItems
        Render-Menu -header $SettingsSubMenuItems
        $setchoice = Read-Host -Prompt "Selection"
        switch($setchoice){
            1 {
                if($settings.showStatus){
                    $settings.showStatus = $false
                }else{
                    $settings.showStatus = $true
                }
            }
            2 {
                Write-Host "Enter SHOW or HIDE to show/hide the welcome message"
                $sh = Read-Host -Prompt "SHOW/HIDE"
                switch($sh){
                    "SHOW" {
                        $settings.showWelcome = $true
                        $err = $false
                    }
                    "HIDE" {
                        $settings.showWelcome = $false
                        $err = $false
                    }
                    default {
                        $err = $true
                        $errmsg = "Invalid choice"

                    }
                }
            }
            3 {
                Write-Host "Current domain is $($settings.exchangeDomain)"
                $newDomain = Read-Host -Prompt "New domain (q to quit)"
                if($newDomain -eq "q"){
                    break
                }
                $settings.exchangeDomain = $newDomain
                Save-Settings
                Read-Host -Prompt "Press Enter to continue"
            }
            4 {
                Write-Host "Current domain controller is $($settings.domainController)"
                $newDC = Read-Host -Prompt "New domain controller (q to quit)"
                if($newDC -eq "q"){
                    break
                }
                $settings.domainController = $newDC
                Save-Settings
                Read-Host -Prompt "Press Enter to continue"
            }
        }
    }
}


function menuBlock {
    $choice = ""
    while($choice -ne "q"){
        Clear-Host
        if($connectedToExchange -and $settings.showStatus){
            Write-Host "Connected to Exchange: $connectedToExchange" -ForegroundColor Green
        }elseif($settings.showStatus){
            Write-Host "Connected to Exchange: $connectedToExchange" -ForegroundColor Red
        }
        if($connectedToGraph -and $settings.showStatus){
            Write-Host "Connected to Graph: $connectedToGraph" -ForegroundColor Green
        }elseif($settings.showStatus){
            Write-Host "Connected to Graph: $connectedToGraph" -ForegroundColor Red
        }
        if($activeDirectoryImported -and $settings.showStatus){
            Write-Host "AD Module Imported: $activeDirectoryImported" -ForegroundColor Green
        }elseif($settings.showStatus){
            Write-Host "AD Module Imported: $activeDirectoryImported" -ForegroundColor Red
        }
        Render-Menu -header $MenuHeader -entries $MenuItems
        Render-Menu -header $SubMenuItems
        $choice = Read-Host -Prompt "Selection"

        switch($choice){ 
            1 { # Choice 1: block sender by EML file
                Disable-Clutter
            }
            2 { # Choice 2: block sender by email address
                Block-SenderByAddress
            }
            3 { # Choice 3: add user to status page by email address
                Add-UserToStatusPage
            }
            4 { # Choice 4: add user to status page by phone number
                $num = Read-Host -Prompt "Phone number (ex: 1234567890 or 123-456-7890)"
                Add-NumberToStatusPage($num)
            }
            5 { # Choice 5: add a powershell script as an app
                Add-ScriptAsApp
            }
            6 { # Choice 6: add anything as a windows app
                Add-AnythingAsApp
            }
            7 { # Choice 7: enable the in-place archive for a user's 365 account
                Enable-Archive
            }
            8 { # Choice 8: add a user to an email group
                Add-UserToEmailGroup
            }
            9 { # Choice 9: Get a list of member of a given email group
                Get-EmailGroupMembers
            }
            10 { # Choice 10: Nav filter prep
                # $rawobjs = Read-Host -Prompt "Object numbers"
                Convert-NewlineToPipe 
            }
            11 { # Choice 11: Dell service tag lookup
                $st = Read-Host -Prompt "Service Tag"
                Lookup-ST -servicetag $st
            }
            12 {
                Convert-NewlineToPipeNOT
            }
            13 {
                Get-FedExTracking
            }
            #----------SUBMENU COMMANDS----------#
            "e" { # Allows executing exchange commands
                Import-Module ExchangeOnlineManagement
                Connect-ExchangeOnline
                $connectedToExchange = $true
            }
            "g" {

            }
            "a" {
                ADmenu
            }
            "c" {
                ChocoMenu
            }
            "l" {
                Get-Changelog
            }
            "q" {
                Save-Settings
            }
            "s" {
                SettingsMenu
            }
        }
        
    }

}

Clear-Host

#----------Message of the day----------#
if($settings.firstTime){
    Write-Host $firstTimeWelcome
    $settings.firstTime = $false
    Save-Settings
    Read-Host -Prompt "Hit Enter to continue"
}elseif($settings.showWelcome){
    Write-Host $welcomeMessage
    Read-Host -Prompt "Hit Enter to continue"
}
#---------------------------------------#

menuBlock
Stop-Transcript
