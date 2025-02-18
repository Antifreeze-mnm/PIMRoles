<# Filename: 		Activate-PIMRole.0.7.ps1
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
#endregion

(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region Selection Form
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PIMRoles-GUI"
        xmlns:Themes="clr-namespace:Microsoft.Windows.Themes;assembly=PresentationFramework.Aero2"
        xmlns:dxe="http://schemas.devexpress.com/winfx/2008/xaml/editors"
        xmlns:dxg="http://schemas.devexpress.com/winfx/2008/xaml/grid"
        xmlns:col="clr-namespace:System.Collections;assembly=mscorlib"
        xmlns:sys="clr-namespace:System;assembly=mscorlib"

        x:Class="System.Windows.Window"
        Title="PIM Roles Activation GUI"
        Width="600"
        MinWidth="600"
        Height="540"
        MinHeight="540"
        Name="ActivationWindow"
        AllowsTransparency="False"
        BorderThickness="0"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        WindowStyle="None">
    <Window.Resources>
        <DataTemplate x:Key="ListBoxItemTemplate">
            <TextBlock Text="{Binding DisplayName}" Foreground="{Binding Foreground}" />
        </DataTemplate>
    </Window.Resources>
    <!-- <WindowChrome.WindowChrome>
    <WindowChrome CaptionHeight="{StaticResource TitleBarHeight}"
                    ResizeBorderThickness="{x:Static SystemParameters.WindowResizeBorderThickness}"
                    CornerRadius="8" />
    </WindowChrome.WindowChrome -->
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- Header Section -->
        <StackPanel Grid.Row="0" Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center" Margin="10">
        </StackPanel>

        <!-- Main Content -->
        <Grid Grid.Row="1" Grid.ColumnSpan="2">
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
            <Button Content="Activate" HorizontalAlignment="Left" VerticalAlignment="Bottom" Width="75" Margin="10,0,10,10" Name="ActivateButton"/>
            <Button Content="Clear Selections" HorizontalAlignment="Right" VerticalAlignment="Bottom" Width="100" Margin="0,0,10,10" Name="ClearButton"/>
        </Grid>
    </Grid>
</Window>
"@

#Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$xaml.SelectNodes("//*[@Name]") | % { Set-Variable -Name ($_.Name) -Value $window.FindName($_.Name) }
#endregion

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
            IsChecked    = $false
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
        $SelectedItems = @($RoleListBox.Items | Where-Object { $_.IsChecked })    
        # Concatenate the display names of the selected roles
        $selectedRoles = $selectedItems | ForEach-Object { $_.DisplayName }
        # Update the SelectedRolesTextBlock
        $SelectedRolesTextBlock.Text = $selectedRoles -join ", "
        foreach ($item in $SelectedItems) {
            if (-not $item.IsSelectable) {
                $item.IsChecked = $false
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
$ClearButton.Add_Click({
        $RoleListBox.SelectedItems.Clear()
        $SelectedRolesTextBlock.Text = ""
        $ReasonTextBox.Clear()
        $DurationTextBox.Clear()
        $HistoryComboBox.SelectedIndex = -1
    })

$ActivateButton.Add_Click({
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
            $roleDefinitionId = $SelectedRole.RoleDefinitionId
            $roleDefinition = $allBuiltInRoles | Where-Object { $_.Id -eq $roleDefinitionId }

            if ($roleDefinition) {
                $assignment = $assignments | Where-Object { $_.RoleDefinitionId -eq $roleDefinitionId }
                $policy = $policies | Where-Object { $_.Id -eq $assignment.PolicyId }

                if ($policy) {
                    $policyId = $policy.Id
                    Write-Output "Role: $($roleDefinition.DisplayName)"
                    Write-Output "Policy ID: $policyId"

                    # Retrieve the policy rule
                    $ruleId = "Expiration_EndUser_Assignment"
                    $rolePolicyRule = Get-MgPolicyRoleManagementPolicyRule -UnifiedRoleManagementPolicyId $policyId -UnifiedRoleManagementPolicyRuleId $ruleId

                    # Access the "AdditionalProperties" field and get the "maximumDuration" value
                    $maximumDurationIso = $rolePolicyRule.AdditionalProperties.maximumDuration

                    # Convert the ISO 8601 duration to a TimeSpan object
                    $maximumDuration = [System.Xml.XmlConvert]::ToTimeSpan($maximumDurationIso)

                    # Compare the user's requested duration with the maximum allowed duration
                    if ($userRequestedDuration -le $maximumDuration) {
                        Write-Output "The requested duration of $DurationHours hours is within the allowed maximum duration of $($maximumDuration.TotalHours) hours."
                        $Duration = $DurationHours
                    }
                    else {
                        Write-Output "The requested duration of $DurationHours hours exceeds the allowed maximum duration of $($maximumDuration.TotalHours) hours. It has been adjusted to $($maximumDuration.TotalHours) hours."
                        $Duration = $maximumDuration.TotalHours
                    }
                }
                else {
                    Write-Output "No policy found for role: $($roleDefinition.DisplayName)"
                    $Duration = $DurationHours
                }
            }
            else {
                Write-Output "No role definition found for role ID: $roleDefinitionId"
                $Duration = $DurationHours
            }

            # Activate PIM role.
            Write-Host "Activating PIM role '$($SelectedRole.RoleDisplayName)'..." -ForegroundColor Blue
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

            New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params | Out-Null
            # Calculate the expiration time in UTC
            $startDateTimeUtc = [datetime]::Parse($Schedule.StartDateTime)
            $expirationTimeUtc = $startDateTimeUtc.AddHours($Duration)
            # Convert the expiration time to local time
            $expirationTimeLocal = $expirationTimeUtc.ToLocalTime()
            # Format the expiration time in local time
            $formattedExpirationTime = $expirationTimeLocal.ToString("hh:mm tt")
            Write-host "$($SelectedRole.RoleDisplayName) has been activated until $formattedExpirationTime!" -ForegroundColor Green
        }

        # Save the current selection to history
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

#show main window
$window.ShowDialog()