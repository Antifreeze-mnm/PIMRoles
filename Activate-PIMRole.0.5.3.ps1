<# Filename: 		Activate-PIMRole.0.5.3.ps1
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
        Title="PIM Role Activation 0.5" Height="540" Width="600">
    <Window.Resources>
        <DataTemplate x:Key="ListBoxItemTemplate">
            <TextBlock Text="{Binding DisplayName}" Foreground="{Binding Foreground}" />
        </DataTemplate>
    </Window.Resources>
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

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$RoleListBox = $window.FindName("RoleListBox")
$SelectedRolesTextBlock = $window.FindName("SelectedRolesTextBlock")
$ReasonTextBox = $window.FindName("ReasonTextBox")
$DurationTextBox = $window.FindName("DurationTextBox")
$HistoryComboBox = $window.FindName("HistoryComboBox")
$ActivateButton = $window.FindName("ActivateButton")
$ClearButton = $window.FindName("ClearButton")
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
        # Concatenate the display names of the selected roles
        $selectedRoles = $selectedItems | ForEach-Object { $_.DisplayName }
        # Update the SelectedRolesTextBlock
        $SelectedRolesTextBlock.Text = $selectedRoles
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

# Close loading window and show main window
$loadingWindow.Close()
$window.ShowDialog()