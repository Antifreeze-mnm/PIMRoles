<# Filename: 		Activate-PIMRole.0.6.ps1
    =============================================================================
    .SYNOPSIS
        Activate an Azure AD Privileged Identity Management (PIM) role with PowerShell.
    .DESCRIPTION
        Presents the user with the PIM Roles available to activate, to select one
        or more roles, provide a reason and duration, and activate the role(s).
        Every activation is saved in a history file, which can be re-used.
        If a Role is already activate it is greyed out and cannot be selected.

    .INPUTS
        None
    .OUTPUTS
        None

    .REFERENCES
        - GitHub:   https://github.com/VitalProject/Show-LoadingScreen/

#>

#region script change log
# originally written by Mark Jackson
#
# Version 0.5   - Initial release Jan 2025
#         0.5.1 - Bug fix: Added User.read permission to the Graph API call
#         0.5.2 - Bug fix: remove numbers in output & fixed issue with repeated entries on the form
#                 when selecting previous history with a duplicate reason
#         0.5.3 - Fixed activation message duration to end time in hh:mm tt format
#         0.6   - Added logging, displayed on the form
#endregion

(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region Loading Screen
# https://github.com/VitalProject/Show-LoadingScreen/
function show-LoadingScreen() {
    param(
        [ValidateLength(0, 15)][String]$note
    )

    $script:LoadingScreen = New-Object system.Windows.Forms.Form
    $LoadingScreen.ClientSize = '150,170'
    $LoadingScreen.TopMost = $false
    $LoadingScreen.BackColor = "#4a4a4a"
    $LoadingScreen.ShowIcon = $false
    $LoadingScreen.FormBorderStyle = "none"
    $LoadingScreen.StartPosition = "CenterScreen"
    $LoadingScreen.TopMost = $True
    $LabelPercent = New-Object system.Windows.Forms.Label
    $LabelPercent.Name = "LoadPercent"
    $LabelPercent.TextAlign = "MiddleCenter"
    $LabelPercent.text = "0%"
    $LabelPercent.AutoSize = $false
    $LabelPercent.width = 150
    $LabelPercent.height = 130
    $LabelPercent.location = New-Object System.Drawing.Point(0, 5)
    $LabelPercent.Font = 'Microsoft Sans Serif,40,style=Bold'
    $LabelPercent.ForeColor = "#D3D6D6"
    $LabelPercent.visible = $false
    $LoadingScreen.controls.Add($LabelPercent)
#Region Image
    $base64Load = "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"
#endregion
    $LoadingData = [Convert]::FromBase64String($base64Load)
    $ms = New-Object IO.MemoryStream($LoadingData, 0, $LoadingData.Length)
    $ms.Write($LoadingData, 0, $LoadingData.Length);
    $LoadingImage = [System.Drawing.Image]::FromStream($ms, $true)
    $PictureBox1 = New-Object system.Windows.Forms.PictureBox
    $PictureBox1.width = 150
    $PictureBox1.height = 130
    $PictureBox1.location = New-Object System.Drawing.Point(0, 5)
    $PictureBox1.Image = $LoadingImage
    $PictureBox1.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::zoom
    $LoadingScreen.controls.Add($PictureBox1)
    $LabelLoading = New-Object system.Windows.Forms.Label
    $LabelLoading.TextAlign = "MiddleCenter"
    $LabelLoading.text = "Loading..."
    $LabelLoading.AutoSize = $false
    $LabelLoading.width = 140
    $LabelLoading.height = 35
    $LabelLoading.location = New-Object System.Drawing.Point(5, 135)
    $LabelLoading.Font = 'Microsoft Sans Serif,15,style=Bold'
    $LabelLoading.ForeColor = "#D3D6D6"
    $LoadingScreen.controls.Add($LabelLoading)
    $LabelNote = New-Object system.Windows.Forms.Label
    $LabelNote.Name = "LoadNote"
    $LabelNote.TextAlign = "MiddleCenter"
    $LabelNote.text = $note
    $LabelNote.AutoSize = $false
    $LabelNote.width = 140
    $LabelNote.height = 55
    $LabelNote.location = New-Object System.Drawing.Point(5, 170)
    $LabelNote.Font = 'Segoe UI Emoji,15,style=Bold'
    $LabelNote.ForeColor = "#D3D6D6"
    #$LabelNote.visible                  = $false
    $LoadingScreen.controls.Add($LabelNote)

    if ($note) {
        $LoadingScreen.ClientSize = '150,225'
        $LabelNote.visible = $true
    }
    $runspace = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("LoadingScreen", $LoadingScreen)
    $data = [hashtable]::Synchronized(@{text = "" })
    $runspace.SessionStateProxy.SetVariable("data", $data)
    $pipeLine = $runspace.CreatePipeline({ $LoadingScreen.ShowDialog() })
    $pipeLine.Input.Close()
    $pipeLine.InvokeAsync()


    $LS = [pscustomobject]@{
        runspace      = $runspace
        pipeline      = $pipeLine
        LoadingScreen = $LoadingScreen
    }

    $LS | Add-Member -MemberType ScriptMethod -Name "close" -Force -Value {
        $this.LoadingScreen.close();
        $this.runspace.close();
    } 

    $LS | Add-Member -MemberType ScriptMethod -Name "updateNote" -Force -Value {
        param(
            [Parameter(Mandatory = $true)][ValidateLength(0, 15)][String]$note
        )
        if ($note) {
            $noteLabel = $this.LoadingScreen.controls | where { $_.name -eq "LoadNote" }
            if ($noteLabel.visible) {
                $this.LoadingScreen.ClientSize = '150,225'
                $noteLabel.text = $note
            }
            else {
                $this.LoadingScreen.ClientSize = '150,225'
                $noteLabel.visible = $true
                $noteLabel.text = $note

            }
        }
        $this.LoadingScreen.Refresh();

    } 
    $LS | Add-Member -MemberType ScriptMethod -Name "updatePercent" -Force -Value {
        param(
            [Parameter(Mandatory = $true)][ValidateRange(0, 99)]$percent
        )

        if ($percent) {
            $percentLabel = $this.LoadingScreen.controls | where { $_.name -eq "LoadPercent" }
            if ($percentLabel.visible) {
                $percentLabel.text = "$($percent)%"
            }
            else {
                $percentLabel.visible = $true
                $percentLabel.text = "$($percent)%"
            }
        }
    
        $this.LoadingScreen.Refresh();

    } 
    $LS | Add-Member -MemberType ScriptMethod -Name "update" -Force -Value {
        param(
            [Parameter(Mandatory = $true)][ValidateRange(0, 99)]$percent,
            [Parameter(Mandatory = $true)][ValidateLength(0, 15)][String]$note
        )
        $this.updateNote($note);
        $this.updatePercent($percent);
    } 

    return $LS
}

# Show the loading screen
$loadingWindow = show-LoadingScreen
#endregion

#region Selection Form
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PIM Role Activation 0.6" Height="600" Width="950">
    <Window.Resources>
        <DataTemplate x:Key="ListBoxItemTemplate">
            <TextBlock Text="{Binding DisplayName}" Foreground="{Binding Foreground}" />
        </DataTemplate>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="350"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <!-- Header Section -->
        <StackPanel Grid.Row="0" Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center" Margin="10">
        </StackPanel>

        <!-- Main Content -->
        <Grid Grid.Row="1" Grid.Column="0">
            <Label Content="Select Roles:" FontSize="14" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,06,0,0" FontWeight="Bold" Foreground="Navy"/>
            <ListBox Name="RoleListBox" HorizontalAlignment="Left" Height="200" VerticalAlignment="Top" Width="560" Margin="10,30,0,0" SelectionMode="Multiple" ItemTemplate="{StaticResource ListBoxItemTemplate}"/>
            <Label Content="Selected Roles:" FontSize="14" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,236,0,0" FontWeight="Bold" Foreground="Navy"/>
            <TextBlock Name="SelectedRolesTextBlock" HorizontalAlignment="Left" VerticalAlignment="Top" Width="560" Margin="10,260,0,0" TextWrapping="Wrap" Foreground="Blue" FontWeight="Bold"/>
            <Label Content="Reason:" FontSize="14" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,286,0,0" FontWeight="Bold" Foreground="Navy"/>
            <TextBox Name="ReasonTextBox" HorizontalAlignment="Left" Height="23" VerticalAlignment="Top" Width="560" Margin="10,310,0,0"/>
            <Label Content="Duration (hours):" FontSize="14" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,336,0,0" FontWeight="Bold" Foreground="Navy"/>
            <TextBox Name="DurationTextBox" HorizontalAlignment="Left" Height="23" VerticalAlignment="Top" Width="100" Margin="10,360,0,0"/>
            <Label Content="Previous Selections:" FontSize="14" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,386,0,0" FontWeight="Bold" Foreground="Navy"/>
            <ComboBox Name="HistoryComboBox" HorizontalAlignment="Left" VerticalAlignment="Top" Width="560" Margin="10,410,0,0"/>
            <Image Name="ActivateImage" Grid.Row="2" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Bottom" Width="75" Height="75" Margin="10,0,10,10"/>
            <Image Name="ClearImage" Grid.Row="2" Grid.Column="1" HorizontalAlignment="Right" VerticalAlignment="Bottom" Width="75" Height="75" Margin="0,0,10,10"/>
        </Grid>

        <!-- Log Window -->
        <RichTextBox Name="LogTextBox" Grid.Row="0" Grid.RowSpan="3" Grid.Column="1" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="10" VerticalScrollBarVisibility="Auto" IsReadOnly="True" Foreground="Yellow" Background="Black"/>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$RoleListBox = $window.FindName("RoleListBox")
$SelectedRolesTextBlock = $window.FindName("SelectedRolesTextBlock")
$ReasonTextBox = $window.FindName("ReasonTextBox")
$DurationTextBox = $window.FindName("DurationTextBox")
$HistoryComboBox = $window.FindName("HistoryComboBox")
$LogTextBox = $window.FindName("LogTextBox")
$ActivateImage = $window.FindName("ActivateImage")
$ClearImage = $window.FindName("ClearImage")
#endregion

function Write-Log {
    param (
        [string]$message,
        [string]$color = "Black"
    )
    $range = New-Object System.Windows.Documents.TextRange($LogTextBox.Document.ContentEnd, $LogTextBox.Document.ContentEnd)
    $range.Text = "$message`r"
    $range.ApplyPropertyValue([System.Windows.Documents.TextElement]::ForegroundProperty, $color)
    $LogTextBox.ScrollToEnd()
}

function Convert-Base64ToImage {
    param (
        [string]$base64String
    )
    try {
        if ([string]::IsNullOrEmpty($base64String)) {
            throw "Base64 string is null or empty."
        }

        Write-Log "Converting Base64 string to byte array..." "Yellow"
        $bytes = [Convert]::FromBase64String($base64String)
        $stream = New-Object System.IO.MemoryStream
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null

        Write-Log "Initializing BitmapImage..." "Yellow"
        $image = New-Object System.Windows.Media.Imaging.BitmapImage
        $image.BeginInit()
        $image.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $image.StreamSource = $stream
        $image.EndInit()
        $image.Freeze()  # To make the image thread-safe
        Write-Log "BitmapImage successfully created." "Green"
        return $image
    } catch {
        Write-Log "Error converting Base64 string to image: $_" "Red"
        return $null
    }
}

#Region Activate Button Image
# Base64 string of the image
$ActivateBase64String ="iVBORw0KGgoAAAANSUhEUgAAANwAAADfCAMAAACqEE0dAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAMAUExURVy0Pmy/P3O/P2zAP3bAP33DP1+mR12tRV6kTF6tSl+4Sl6zRV65RF6zSV6pUmapTGGuRmKmTWKsS2mxSmG1RWS6RWy+RWGzSWS6SGm9SHi/RHK/Q2apWWWlU2mmVWSqUmqqVWylW2yqWnCmXnGrXmmxVHKyXW6pYnmqaHKnYnSsZHWraXutbG2yY3q4bXWyY3myZnuybH2ucXGwcH6zcmzARHXAQnvCQoXFP4vGPo3IPp7NNpLHPZ3GPZPJPZzLO5/QOKHGPKHNNqDNOqLQO5u+R5+8WYCuboCzb4KvdYe5eYKzdIS7dYa1eYu2fIu6fJG7foPFQYvGQI3IQJ/LSZHHQJ3FRZLIQJ7KQp3FSZ/QQJ7GWZ7DUqnKT6LGRKLKQ6LFS6PJS6TRQ6fJWKPFU6XKU6rMU6TFWqnGW6zLXLDNXqzTW4fAeqjGaKrGY6vKZKzKarLNY7LNbK/QbbPVZ7LVa63McLjOebPNcqzQcbrWdLXSdLbSerrUfL3dfb/ieMLdf4u6ho27gpq9jZG8g5W+ip6+mZu+ko/BhprIipPBhJXCipvDjZ/MnpzEkp7Kkb7ahbzVhL3ZiqXJmaHElKLKlaXLm6nMnaTVmavSnabJoazOorPNqa3Spq3To7rYrrLSpbTTq7XZq7jPsbfWtLvVs7zas77bub7huMLWh8HbhMPci8zemMXcksnhjNDmjsrklMzklczjm9LnnMPauMLWvMPcvM/hpNbpqNPmpNXrpNrtrNnxqN3yrcnjvMTivNzuudvts9/xu93ytOLyrOHuuOT5tuL1s+PzvOb6vej5vsrVxcXcwsrdw87dytDdxtPezc/c0dfX2NfX1tbc1Nnb1tzX29fa2dra2uDZ3MXiwcvixM3iy9PhydLjzdTpzdXg2NXi09ni09Tq0tnh2eLvwun1zOT1w+n1xOb6w+r7xOv7yvH/zeHh3u362e320+z80vP50/P/1PX/3Pr/3N3Z4eDd4u744/b/4vr/5Pb+6/v/7ff/8f3/9P///gAAADe2qO0AAAEAdFJOU////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////wBT9wclAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAAGHRFWHRTb2Z0d2FyZQBQYWludC5ORVQgNS4xLjL7vAO2AAAAtmVYSWZJSSoACAAAAAUAGgEFAAEAAABKAAAAGwEFAAEAAABSAAAAKAEDAAEAAAACAAAAMQECABAAAABaAAAAaYcEAAEAAABqAAAAAAAAAAx3AQDoAwAADHcBAOgDAABQYWludC5ORVQgNS4xLjIAAwAAkAcABAAAADAyMzABoAMAAQAAAAEAAAAFoAQAAQAAAJQAAAAAAAAAAgABAAIABAAAAFI5OAACAAcABAAAADAxMDAAAAAA7uotGMbMAAIAAEX3SURBVHhe3b0JfFzVmeZtY0AplcoqlUoqeS3bmCxAJ8AXG++2lnwsYbW8KOqgksEEmkkwkE6HL+DuDoFJBmgIYZs0acAETDqETEIHvgQS6IReIMEg7CqV5EVWyepMOrGTtG25bWzBPM/znnvr3pJky8bM/H7zVqkk1Xr+993PPffWmHf/L5b/bXC79unyrm7+N8n7CgeO7i5f8u63k+5d3bv2uSe+T/J+wTmCwwlhc/h5/xT5fsBp6F0BnQ0jgUc3bsRNt3vtcZXjDdfFkUJ22a8jMQLN/UCR7i2OnxxXOJ/Il3342QVA/g7IiMTujY6THD+4EFl31wCV1+/+1e+h6EOl+7h64HGCGzJw4gz0U2W78NtQRyH2LPem71mOC5xGFBBYHbE8vXlyRL6itbo3fo9yHOBogsMI1eWcjeobpep8OR7m+V7hOAzLzuEg4VCovF1DVDhKcR9x7PLe4EqIwkI/0+PQIPBG0Fyn+62/iv9A8AL3Mccq7wWOIyhVSp41B3icz3kP57pym7I5PJjLbdpkT5F0dnZ1ehuoszs3dFu9pwLt2OHcp5t4o8K4c7LSAXJl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"
#endregion

#Region Clear Selections Button Image
# Base64 string of the image
$ClearBase64String ="iVBORw0KGgoAAAANSUhEUgAAAOQAAADkCAMAAAC/iXi/AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAMAUExURQs+ABBWDilXDwdjHBxkGy5vCyR3BDl6AyhnHDNpHRJfKQRdJjFaKBVsMAliIQJ4ODNyLihlJDVrKEZeAEt3B0J7GFJwMUdoODtqRVFwTkhxSWl6TGx8Z3J/aS2CGDeEHAaDPBiIOQWUPBmYPjWGKiONOCqVOzeaOy6iPTukPVKMHUaCAFmGBlmIE22TCWOKCWOLE0+GMUSdN1qbNWqULlWzNkinPFasO1qyO26oMHWlLme2LGSmOGi2PXi8PX3BPQ+NRwKKRQSUSxiaSQqaUxScVTaeUiSeRzWdSCadVxiiRQ6gWBakWyuiRzimRienVjerVjyyWSufYByrYj6vcSWtZTOtaymxajOxbjuzc0eVRWyPU0erRFOtQ02xQ1iyQ0ioVFesVki1WVe4WXiqSmWtSGe3Qna7QmeuWGa5WXa7WFuAYXiQbmyHZnWEaVm0YUm2elG3e3azdGq6ZHW9ZF7BXH3BQWjDXHbHW3vLbmnCYnnGZoWtFoupLIO9O5e7OozFJ5nILYjDPJbIO6fOL6bMOrDOPK3RPLHSPYWeTI67UoW+R6S+TYubc5KydYe7Z53QUojCQ5fIRIfIWpnKV7TPX6bNRrLPRavSRrTUSKbOV6nSV7jXWIzRcojKZ5nNaYrRaZnSa4fNcpjMcpvSc7HPdqPNaqbUa7vaZqbOc6nTeLfadq/gbLXha7ric8HdXMPgXsLda8PddcfjacniedDpfFG6gG+/i3O9im/Ai3nCjIeThZmylJeniJemlKezl5u/oLCzq625p7O7p5rOiobGiKrOj67WgrTYiazVmrvclrPijb3hm6rTrLTEqqzSorraprjHs7ndub7iorzhtsfehsLem9rwn8zkhtLoisrjmNbpmNDcv8PeosLLu8Tbucfjptvrp97wqcnktdnqtt/xs+DvreHxq+HutuTxub7XwsTNxsjTx9Pcy9fd19bx0cnjx9rrx9vxytvl2eHuyOj0yOPq2ez21fL52t3q4Obt5uv26PT65u3+8v7//QAAAODVVgEAAAEAdFJOU////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////wBT9wclAAAACXBIWXMAAA7DAAAOwwHHb6hkAAAAGHRFWHRTb2Z0d2FyZQBQYWludC5ORVQgNS4xLjL7vAO2AAAAtmVYSWZJSSoACAAAAAUAGgEFAAEAAABKAAAAGwEFAAEAAABSAAAAKAEDAAEAAAACAAAAMQECABAAAABaAAAAaYcEAAEAAABqAAAAAAAAAAp3AQDoAwAACncBAOgDAABQYWludC5ORVQgNS4xLjIAAwAAkAcABAAAADAyMzABoAMAAQAAAAEAAAAFoAQAAQAAAJQAAAAAAAAAAgABAAIABAAAAFI5OAACAAcABAAAADAxMDAAAAAAauiJOT9thO8AAESNSURBVHhevZ0JQFXXue/zOtzbIWmadEjHd402Vt5te5ukAdIo6gFEiYDGaJOUSY2g3rS5vRo14YU0TnhL08igIhAGGQ4HnADtu+/dtrlNegMFm6Y14BRSCDjdJmk1GhMH8v7/71tr733gMGhM//ucw+Fwzt7rt//f+ta39t7ANe/9bXWqW9Rrvv3b6AOH7Ovr7evt7erc14lbm2eBurq78MO+vtPmvR+UPkjIU6f79nft29fW1t7WCrXzIUjtre1tbW379nV1dXef+gBJPzDI3u7uLvAJXouo3lVNfX1Lvb7Kn+NtbW37u7rPmM9ebX0QkH29XYhM4SOGYAVCCrDAdVG7uno/CNCrDtnbhQh18YSlVlWMxSt8F6iVNwRgLqwl6b59+0+ZVV01XV3Ivu4uWEhC4PkD/ura2nLegFFTU9PS0NrQ2tTawYeGhpaaGrwsb1BWxnBLSzs4u7qvLudVhOztAyE9bGiBgX51C3AtHdBh6ggW+dptnhzsaG9oqKlXq/212C0SvO3k7DOrvQq6apAYJZSQMVpbxUbXNDQA7+BhJRSBjsMkHxSTtNwJtFVQGbro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"
#endregion

# Convert Base64 string to image
$activateImage64 = Convert-Base64ToImage -base64String $ActivateBase64String
# Convert Base64 string to image for Clear Selections
$clearImage64 = Convert-Base64ToImage -base64String $ClearBase64String
$ActivateImage.Source = $ActivateImage64
$ClearImage.Source = $ClearImage64

#region Generate Roles
# Connect to Microsoft Graph using device code authentication
Connect-MgGraph -Scopes "RoleManagement.Read.Directory", "RoleManagement.ReadWrite.Directory", "User.Read" -UseDeviceAuthentication -NoWelcome

# Get the current user's ID
$CurrentAccountId = (Get-AzContext).Account.Id
$CurrentUser = Get-MgUser -UserId $CurrentAccountId
$CurrentAccountId = $CurrentUser.Id

# Retrieve role eligibility schedules for the current user
$PimRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "principalId eq '$CurrentAccountId'" -ExpandProperty "roleDefinition"

# Retrieve all built-in roles
$allBuiltInRoles = Get-MgRoleManagementDirectoryRoleDefinition -All

# Retrieve role management policy assignments and policies
$assignments = Get-MgPolicyRoleManagementPolicyAssignment -Filter "scopeId eq '/' and scopeType eq 'Directory'"
$policies = Get-MgPolicyRoleManagementPolicy -Filter "scopeId eq '/' and scopeType eq 'Directory'"

# Get all active PIM role assignments for the current user
$ActivePimRoles = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$CurrentAccountId'" -ExpandProperty "roleDefinition"
# Populate the RoleListBox with roles, sorted alphabetically
$SortedRoles = $PimRoles | Sort-Object { $_.RoleDefinition.DisplayName }
foreach ($Role in $SortedRoles) {
    $RoleDisplayName = $Role.RoleDefinition.DisplayName
    #Write-Output "Processing Role: $RoleDisplayName"  # Debugging output
    $IsActive = $ActivePimRoles | Where-Object { $_.RoleDefinition.DisplayName -eq $RoleDisplayName }
    $Foreground = if ($IsActive) { "Gainsboro" } else { "Black" }
    $IsSelectable = if ($IsActive) { $false } else { $true }
    $RoleListBox.Items.Add([PSCustomObject]@{
            DisplayName  = $RoleDisplayName
            Foreground   = $Foreground
            IsSelectable = $IsSelectable
        }) | Out-Null
}
#endregion

#region Load History
# Load previous selections
$savePath = "$env:USERPROFILE\Documents\PIMRoleSelections.json"
$history = [System.Collections.Generic.List[PSObject]]::new()

try {
    if (Test-Path $savePath) {
        $historyContent = Get-Content -Path $savePath -ErrorAction Stop
        $historyArray = $historyContent | ConvertFrom-Json -ErrorAction Stop

        # Ensure $historyArray is always an array
        if ($historyArray -isnot [array]) {
            $historyArray = @($historyArray)
        }
        # Populate the ComboBox with history entries
        $HistoryComboBox.Items.Clear()
        foreach ($entry in $historyArray) {
            $history.Add([PSCustomObject]$entry)
            $HistoryComboBox.Items.Add([PSCustomObject]@{
                Id     = $entry.Id
                Reason = $entry.Reason
            }) | Out-Null
        }
        # Display only the reason in the ComboBox
        $HistoryComboBox.DisplayMemberPath = "Reason"
    }
}
catch {
    Write-Error "Failed to load history: $_"
}
#endregion

#region populate Selection form
# Handle selection to prevent selecting unselectable items
$RoleListBox.Add_SelectionChanged({
        $SelectedItems = @($RoleListBox.SelectedItems)
        foreach ($item in $SelectedItems) {
            if (-not $item.IsSelectable) {
                $RoleListBox.SelectedItems.Remove($item)
                [System.Windows.MessageBox]::Show("The role '$($item.DisplayName)' is already activated and cannot be selected.")
            }
        }
    })

# Handle the SelectionChanged event
$HistoryComboBox.Add_SelectionChanged({
    $selectedItem = $HistoryComboBox.SelectedItem
    $selectedId = $selectedItem.Id
    $selectedHistory = $history | Where-Object { $_.Id -eq $selectedId }
    if ($selectedHistory) {
        $ReasonTextBox.Text = $selectedHistory.Reason
        $DurationTextBox.Text = $selectedHistory.Duration.TotalHours
        $RoleListBox.SelectedItems.Clear()
        foreach ($role in $selectedHistory.SelectedRoles) {
            $RoleItem = $RoleListBox.Items | Where-Object { $_.DisplayName -eq $role.RoleDisplayName }
            if ($RoleItem -and $RoleItem.IsSelectable) {
                $RoleListBox.SelectedItems.Add($RoleItem)
            }
        }
        # Update the SelectedRolesTextBlock
        $SelectedRolesTextBlock.Text = ($RoleListBox.SelectedItems | ForEach-Object { $_.DisplayName }) -join ", "
    }
})
#endregion

#region Process Form selections
$ClearImage.MouseLeftButtonUp.Add({
        Write-Log "Clearing selections..." "White"
        $RoleListBox.SelectedItems.Clear()
        $SelectedRolesTextBlock.Text = ""
        $ReasonTextBox.Clear()
        $DurationTextBox.Clear()
        $HistoryComboBox.SelectedIndex = -1
    })

$ActivateImage.MouseLeftButtonUp.Add({
        Write-Log "Activate button clicked. Processing selections..." "White"
        $SelectedRoles = @()
        foreach ($SelectedItem in $RoleListBox.SelectedItems) {
            $Role = $PimRoles | Where-Object { $_.RoleDefinition.DisplayName -eq $SelectedItem.DisplayName }
            $SelectedRoles += [PSCustomObject]@{
                RoleDisplayName  = $Role.RoleDefinition.DisplayName
                RoleDefinitionId = $Role.RoleDefinition.Id
            }
        }

        $Reason = $ReasonTextBox.Text
        $DurationHours = [double]$DurationTextBox.Text
        $userRequestedDuration = [TimeSpan]::FromHours($DurationHours)

        # Activate selected roles
        foreach ($SelectedRole in $SelectedRoles) {
            # Retrieve the policy for the role
            Write-Log "   Processing role: $($SelectedRole.RoleDisplayName)..." "Yellow"
            $roleDefinitionId = $SelectedRole.RoleDefinitionId
            $roleDefinition = $allBuiltInRoles | Where-Object { $_.Id -eq $roleDefinitionId }

            if ($roleDefinition) {
                $assignment = $assignments | Where-Object { $_.RoleDefinitionId -eq $roleDefinitionId }
                $policy = $policies | Where-Object { $_.Id -eq $assignment.PolicyId }

                if ($policy) {
                    $policyId = $policy.Id
                    #Write-Output "Role: $($roleDefinition.DisplayName)"
                    #Write-Output "Policy ID: $policyId"
                    # Retrieve the policy rule
                    $ruleId = "Expiration_EndUser_Assignment"
                    $rolePolicyRule = Get-MgPolicyRoleManagementPolicyRule -UnifiedRoleManagementPolicyId $policyId -UnifiedRoleManagementPolicyRuleId $ruleId

                    # Access the "AdditionalProperties" field and get the "maximumDuration" value
                    $maximumDurationIso = $rolePolicyRule.AdditionalProperties.maximumDuration

                    # Convert the ISO 8601 duration to a TimeSpan object
                    $maximumDuration = [System.Xml.XmlConvert]::ToTimeSpan($maximumDurationIso)

                    # Compare the user's requested duration with the maximum allowed duration
                    Write-Log "   Checking duration does not exceed allowed maximum..." "Yellow"
                    if ($userRequestedDuration -le $maximumDuration) {
                        Write-Log "Requested duration $DurationHours hours, within allowed maximum." "White"
                        $Duration = $DurationHours
                    }
                    else {
                        Write-Log "Requested duration $DurationHours hours exceeds maximum. Adjusted to $($maximumDuration.TotalHours) hours." "Red"
                        $Duration = $maximumDuration.TotalHours
                    }
                }
                else {
                    Write-Log "No policy found for role: $($roleDefinition.DisplayName)" "Red"
                    $Duration = $DurationHours
                }
            }
            else {
                Write-Log "Role definition for role ID $roleDefinitionId not found:" "Red"
                $Duration = $DurationHours
            }

            # Activate PIM role.
            Write-Log "Activating PIM role '$($SelectedRole.RoleDisplayName)'..." "White"
            # Create activation schedule based on the current role limit.
            $Schedule = @{
                StartDateTime = (Get-Date).ToUniversalTime()
                Expiration    = @{
                    Type     = "AfterDuration"
                    Duration = "PT${Duration}H"
                }
            }

            # Setup parameters for activation
            $params = @{
                Action           = "selfActivate"
                PrincipalId      = $CurrentAccountId
                RoleDefinitionId = $SelectedRole.RoleDefinitionId
                DirectoryScopeId = "/"
                Justification    = $Reason
                ScheduleInfo     = $Schedule
            }
            #New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params | Out-Null
            # Calculate the expiration time in UTC
            $startDateTimeUtc = [datetime]::Parse($Schedule.StartDateTime)
            $expirationTimeUtc = $startDateTimeUtc.AddHours($Duration)
            # Convert the expiration time to local time
            $expirationTimeLocal = $expirationTimeUtc.ToLocalTime()
            # Format the expiration time in local time
            $formattedExpirationTime = $expirationTimeLocal.ToString("hh:mm tt")
            Write-Log "   $($SelectedRole.RoleDisplayName) has been activated until $formattedExpirationTime!" "Yellow"
        }

        # Save the current selection to history
        Write-Log "   Saving current selection to history..." "Yellow"
        $newEntry = [PSCustomObject]@{
            Id            = [guid]::NewGuid().ToString()
            Reason        = $Reason
            Duration      = $userRequestedDuration
            SelectedRoles = $SelectedRoles
        }

        # Check for duplicates before adding to history
        $existingEntry = $history | Where-Object {
            $_.Reason -eq $newEntry.Reason -and
            $_.Duration -eq $newEntry.Duration -and
    ($_.SelectedRoles | ForEach-Object { $_.RoleDisplayName }) -join "," -eq ($newEntry.SelectedRoles | ForEach-Object { $_.RoleDisplayName }) -join ","
        }

        if (-not $existingEntry) {
            $history.Add($newEntry)
            $HistoryComboBox.Items.Add($newEntry.Reason)
            $history | ConvertTo-Json -Depth 10 | Set-Content -Path $savePath
        }

        # Load existing history
        if (Test-Path $savePath) {
            $history = Get-Content -Path $savePath | ConvertFrom-Json -Depth 10
            if ($history -isnot [System.Collections.ArrayList]) {
                $history = @($history)
            }
        }
        else {
            $history = @()
        }

        [System.Windows.MessageBox]::Show("Roles activated successfully!")
        $window.Close()
    })
#endregion

# Close loading window and show main window
$loadingWindow.Close()
$window.ShowDialog()