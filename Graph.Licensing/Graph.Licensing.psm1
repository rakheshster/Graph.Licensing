#Requires -Modules Microsoft.PowerShell.ConsoleGuiTools

New-Variable -Name subscribedSkuIdHashTable -Scope Script -Force
New-Variable -Name subscribedSkuNameHashTable -Scope Script -Force
New-Variable -Name skuNameHashTable -Scope Script -Force
New-Variable -Name skuIdHashTable -Scope Script -Force
New-Variable -Name planNameHashTable -Scope Script -Force
New-Variable -Name planIdHashTable -Scope Script -Force

function Update-MgLicensingData {
    <#
    .DESCRIPTION

    Run this one-time (and preferably periodically) to pull the latest info on the license part numbers, GUIDS, etc.

    You must be logged in to Graph before running this cmdlet.

    .EXAMPLE
    Update-MgLicensingData
    #>

    begin {
        $scopes = (Get-MgContext).Scopes

        if ($scopes -contains "Directory.ReadWrite.All" -or ($scopes -contains "User.ReadWrite.All" -and $scopes -contains "Group.ReadWrite.All")) {
            Write-Output "✔ You seem to have the required permissions."
    
        } elseif ($scopes -contains "User.ReadWrite.All") {
            Write-Output "👉 Warning: You will only be able to make changes to direct user license assignments."
    
        } elseif ($scopes -contains "User.ReadWrite.All") {
            Write-Output "👉 Warning: You will only be able to make changes to group license assignments."
    
        } elseif ($scopes -contains "Directory.Read.All" -or ($scopes -contains "User.Read.All" -and $scopes -contains "Group.Read.All")) {
            Write-Output "👉 Warning: You will only be able to read info, not make any changes."
    
        } elseif ($scopes -contains "User.Read.All") {
            Write-Output "👉 Warning: You will only be able to read user info, not make any changes."
    
        } elseif ($scopes -contains "Group.Read.All") {
            Write-Output "👉 Warning: You will only be able to read group info, not make any changes."
    
        } else {
            Write-Output "🚫 Missing the required permissions. Are you connected to Graph and have the right permissions?"
            Write-Output "For example: Connect-MgGraph -Scopes 'Directory.Read.All'"
            return
        }
    }

    process {
        $subscribedSkuIdHashTable = @{}
        $subscribedSkuNameHashTable = @{}
        Write-Output "〰 Creating a list of available licenses from Graph"
        Get-MgSubscribedSku | Where-Object { $_.AppliesTo -eq "User" } | ForEach-Object {
            $subscribedSkuIdHashTable[$_.SkuId] = @{
                "PlansIncludedFriendlyName" = $_.ServicePlans
                "SkuPartNumber" = $_.SkuPartNumber
            }
    
            $subscribedSkuNameHashTable[$_.SkuPartNumber] = @{
                "PlansIncludedFriendlyName" = $_.ServicePlans
                "SkuId" = $_.SkuId
            }
        }
    
        Write-Output "〰 Downloading latest licenses and plan details"
        # From https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
        $licenseCsvURL = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    
        $skuNameHashTable = @{}
        $skuIdHashTable = @{}
        $planNameHashTable = @{}
        $planIdHashTable = @{}
    
        (Invoke-WebRequest -Uri $licenseCsvURL).ToString() -replace "\?","" | ConvertFrom-Csv | ForEach-Object {
            # Maps "Office 365 E5" to its SkuId
            $skuNameHashTable[$_.Product_Display_Name] = @{
                "SkuId" = $_.GUID
                "SkuPartNumber" = $_.String_Id
                "PlansIncludedFriendlyName" = if ($skuNameHashTable[$_.Product_Display_Name].PlansIncludedFriendlyName.Length -eq 0 ) {
                    @($_.Service_Plans_Included_Friendly_Names) 
                } else { 
                    @($skuNameHashTable[$_.Product_Display_Name].PlansIncludedFriendlyName) + @($_.Service_Plans_Included_Friendly_Names)
                }
    
                "PlansIncludedName" = if ($skuNameHashTable[$_.Product_Display_Name].PlansIncludedName.Length -eq 0 ) {
                    @($_.Service_Plan_Name) 
                } else { 
                    @($skuNameHashTable[$_.Product_Display_Name].PlansIncludedName) + @($_.Service_Plan_Name)
                }
    
                "PlansIncludedIds" = if ($skuNameHashTable[$_.Product_Display_Name].PlansIncludedIds.Length -eq 0 ) {
                    @($_.Service_Plan_Id) 
                } else { 
                    @($skuNameHashTable[$_.Product_Display_Name].PlansIncludedIds) + @($_.Service_Plan_Id)
                }
            }
    
            # Maps SkuId to "Office 365 E5" 
            $skuIdHashTable[$_.GUID] = @{
                "SkuPartNumber" = $_.String_Id
                "DisplayName" = $_.Product_Display_Name
                "PlansIncludedFriendlyName" = if ($skuIdHashTable[$_.GUID].PlansIncludedFriendlyName.Length -eq 0 ) {
                    @($_.Service_Plans_Included_Friendly_Names) 
                } else { 
                    @($skuIdHashTable[$_.GUID].PlansIncludedFriendlyName) + @($_.Service_Plans_Included_Friendly_Names)
                }
    
                "PlansIncludedName" = if ($skuIdHashTable[$_.GUID].PlansIncludedName.Length -eq 0 ) {
                    @($_.Service_Plan_Name) 
                } else { 
                    @($skuIdHashTable[$_.GUID].PlansIncludedName) + @($_.Service_Plan_Name)
                }
    
                "PlansIncludedIds" = if ($skuIdHashTable[$_.GUID].PlansIncludedIds.Length -eq 0 ) {
                    @($_.Service_Plan_Id) 
                } else { 
                    @($skuIdHashTable[$_.GUID].PlansIncludedIds) + @($_.Service_Plan_Id)
                }
            }
    
            # Maps a plan name its Id and all the SKUs it is a part both (both SKU Ids and Names)
            $planNameHashTable[$_.Service_Plans_Included_Friendly_Names] = @{
                "PlanId" = $_.Service_Plan_Id
                "Skus" = if ($planNameHashTable[$_.Service_Plans_Included_Friendly_Names].Skus.Length -eq 0 ) {
                    @($_.Product_Display_Name) 
                } else { 
                    @($planNameHashTable[$_.Service_Plans_Included_Friendly_Names].Skus) + @($_.Product_Display_Name)
                }
    
                "SkuPartNumbers" = if ($planNameHashTable[$_.Service_Plans_Included_Friendly_Names].SkuPartNumbers.Length -eq 0 ) {
                    @($_.String_Id) 
                } else { 
                    @($planNameHashTable[$_.Service_Plans_Included_Friendly_Names].SkuPartNumbers) + @($_.String_Id)
                }
                
                "SkuIds" = if ($planNameHashTable[$_.Service_Plans_Included_Friendly_Names].SkuIds.Length -eq 0 ) {
                    @($_.GUID) 
                } else { 
                    @($planNameHashTable[$_.Service_Plans_Included_Friendly_Names].SkuIds) + @($_.GUID)
                }
            }
    
            # Maps a plan Id to its friendly name and all the SKUs it is a part both (both SKU Ids and Names)
            $planIdHashTable[$_.Service_Plan_Id] = @{
                "DisplayName" = $_.Service_Plans_Included_Friendly_Names
                "Skus" = if ($planIdHashTable[$_.Service_Plan_Id].Skus.Length -eq 0 ) {
                        @($_.Product_Display_Name) 
                } else { 
                        @($planIdHashTable[$_.Service_Plan_Id].Skus) + @($_.Product_Display_Name)
                }
    
                "SkuPartNumbers" = if ($planIdHashTable[$_.Service_Plan_Id].SkuPartNumbers.Length -eq 0 ) {
                    @($_.String_Id) 
                } else { 
                    @($planIdHashTable[$_.Service_Plan_Id].SkuPartNumbers) + @($_.String_Id)
                }
                
                "SkuIds" = if ($planIdHashTable[$_.Service_Plan_Id].SkuIds.Length -eq 0 ) {
                    @($_.GUID) 
                } else { 
                    @($planIdHashTable[$_.Service_Plan_Id].SkuIds) + @($_.GUID)
                }
            }
        }
    }

    end {
        $script:subscribedSkuIdHashTable = $subscribedSkuIdHashTable
        $script:subscribedSkuNameHashTable = $subscribedSkuNameHashTable
        $script:skuNameHashTable = $skuNameHashTable
        $script:skuIdHashTable = $skuIdHashTable
        $script:planNameHashTable = $planNameHashTable
        $script:planIdHashTable = $planIdHashTable
    
        Write-Output "✔ All done!"
    }
}

function Get-MgAssignedLicenses {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "Group")]
        [string]$GroupName,

        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "User")]
        [Alias("UPN")]
        [string]$UserPrincipalName,

        [Switch]$ShowPlansOnly,
        [Switch]$SortPlansByState
    )

    <#
    .DESCRIPTION
    SHow the licenses assigned to a group or user. By default it shows all the licenses and you can select one or more by pressing the SPACE key and then ENTER to see the plans of the licenses you selected.

    .PARAMETER GroupName
    The Group you'd like to see the license assignments of. 

    .PARAMETER UserPrincipalName
    The User you'd like to see the license assignments of. 

    .PARAMETER ShowPlansOnly
    The default output shows the licenses and you must select one or more to see the plan details. Use this switch to skip that and show all the plans assigned to the user or group across all licenses assigned to it.

    .PARAMETER SortPlansByState
    By default plans are sorted alphabetically. Use this to sort them by On/ Off state. 
    #>

    begin {
        $subscribedSkuIdHashTable = $script:subscribedSkuIdHashTable
        $subscribedSkuNameHashTable = $script:subscribedSkuNameHashTable
        $skuNameHashTable = $script:skuNameHashTable
        $skuIdHashTable = $script:skuIdHashTable
        $planNameHashTable = $script:planNameHashTable
        $planIdHashTable = $script:planIdHashTable

        if ($subscribedSkuIdHashTable.Count -eq 0 -or $subscribedSkuNameHashTable -eq 0 -or 
                $skuNameHashTable -eq 0 -or $skuIdHashTable -eq 0 -or 
                $planNameHashTable  -eq 0 -or $planIdHashTable -eq 0) {
            
            Update-MgLicensingData

            $subscribedSkuIdHashTable = $script:subscribedSkuIdHashTable
            $subscribedSkuNameHashTable = $script:subscribedSkuNameHashTable
            $skuNameHashTable = $script:skuNameHashTable
            $skuIdHashTable = $script:skuIdHashTable
            $planNameHashTable = $script:planNameHashTable
            $planIdHashTable = $script:planIdHashTable
        }

        if ($PSCmdlet.ParameterSetName -eq "Group") {
            try {
                $groupObj = Get-MgGroup -Filter "DisplayName eq '$GroupName'" -Property assignedLicenses,Id,LicenseAssignmentStates -ErrorAction Stop
            
                if ($null -eq $groupObj) {
                    Write-Output "⛔ Couldn't find $GroupName"
                }
            
            } catch {
                Write-Output "⛔ Error searching for ${GroupName} - $($_.Exception.Message)"
            }

            $assignedLicenses = $groupObj.AssignedLicenses

            $assignmentPaths = @{}
            foreach ($skuIdTemp in $assignedLicenses.SkuId) {
                $assignmentPaths[$skuIdTemp] = "Direct"
            }
        }

        if ($PSCmdlet.ParameterSetName -eq "User") {
            try {
                $userObj = Get-MgUser -Filter "UserPrincipalName eq '$UserPrincipalName'" -Property assignedLicenses,Id,LicenseAssignmentStates -ErrorAction Stop
            
                if ($null -eq $userObj) {
                    Write-Output "⛔ Couldn't find $GroupName"
                }
            
            } catch {
                Write-Output "⛔ Error searching for ${GroupName} - $($_.Exception.Message)"
            }

            $assignedLicenses = $userObj.AssignedLicenses

            $assignmentPaths = @{}
            foreach ($skuAssignment in $userObj.LicenseAssignmentStates) {
                $assignmentPaths[$($skuAssignment.SkuId)] = if ($skuAssignment.AssignedByGroup) { "Group" } else { "Direct" }
            }
        }
    }

    process {
        # If we have to show only the plans, it's straight forward...
        if ($ShowPlansOnly) {
            $assignedLicenses | ForEach-Object {
                $skuAssignedToObject = $_.SkuId
                # The plans that are disabled for this license assignments
                $disabledPlans = $_.DisabledPlans
    
                # All the plans that are actually available for this license SKU
                $skuApplicablePlans = ($subscribedSkuIdHashTable[$skuAssignedToObject].PlansIncludedFriendlyName | Where-Object { $_.AppliesTo -eq "User" }).ServicePlanId
    
                $planStates = foreach ($planId in $skuIdHashTable[$skuAssignedToObject].PlansIncludedIds) {
                    if ($planId -notin $skuApplicablePlans)  { continue }
    
                    [pscustomobject][ordered]@{
                        "PlanName" = $planIdHashTable[$planId].DisplayName
                        "PlanId" = $planId
                        "State" = if ($disabledPlans -contains $planId) { "Off" } else { "On" }
                        # Yes, I repeat this info here and below. That's coz I might be viewing the plans only via -ShowPlansOnly
                        "SkuName" = $skuIdHashTable[$skuAssignedToObject].DisplayName
                        "SkuId" = $skuAssignedToObject
                    }
                }

                if ($SortPlansByState) {
                    $planStates | Sort-Object -Descending -Property { $_.State }

                } else {
                    $planStates | Sort-Object -Descending -Property { $_.PlanName }
                }

            } | Out-ConsoleGridView

        } else {
        # If we are showing licenses to, I want to capture the user selections and then loop into a plan only view for each of their selections :)
            $userSelections = $assignedLicenses | ForEach-Object {
                $skuAssignedToObject = $_.SkuId
                # The plans that are disabled for this license assignments
                $disabledPlans = $_.DisabledPlans
    
                # All the plans that are actually available for this license SKU
                $skuApplicablePlans = ($subscribedSkuIdHashTable[$skuAssignedToObject].PlansIncludedFriendlyName | Where-Object { $_.AppliesTo -eq "User" }).ServicePlanId
    
                $planStates = foreach ($planId in $skuIdHashTable[$skuAssignedToObject].PlansIncludedIds) {
                    if ($planId -notin $skuApplicablePlans)  { continue }
    
                    [pscustomobject][ordered]@{
                        "PlanName" = $planIdHashTable[$planId].DisplayName
                        "PlanId" = $planId
                        "State" = if ($disabledPlans -contains $planId) { "Off" } else { "On" }
                        # Yes, I repeat this info here and below. That's coz I might be viewing the plans only via -ShowPlansOnly
                        "SkuName" = $skuIdHashTable[$skuAssignedToObject].DisplayName
                        "SkuId" = $skuAssignedToObject
                    }
                }
    
    
                $totalCount = $skuApplicablePlans.Count
                $enabledCount = $skuApplicablePlans.Count - $_.DisabledPlans.Count
    
                [pscustomobject][ordered]@{
                    "SkuName" = $skuIdHashTable[$skuAssignedToObject].DisplayName
                    "SkuId" = $skuAssignedToObject
                    "AssignmentPath" = $assignmentPaths[$skuAssignedToObject]
                    "Plans" = $planStates
                    "Count" = "${enabledCount}/${totalCount}"
                }
    
            } | Out-ConsoleGridView

            # If the user made selection, expand into the plans
            foreach ($selection in $userSelections) {
                $assignedLicenses | Where-Object { $_.SkuId -eq $selection.SkuId } | ForEach-Object {
                    $skuAssignedToObject = $_.SkuId
                    # The plans that are disabled for this license assignments
                    $disabledPlans = $_.DisabledPlans
        
                    # All the plans that are actually available for this license SKU
                    $skuApplicablePlans = ($subscribedSkuIdHashTable[$skuAssignedToObject].PlansIncludedFriendlyName | Where-Object { $_.AppliesTo -eq "User" }).ServicePlanId
        
                    $planStates = foreach ($planId in $skuIdHashTable[$skuAssignedToObject].PlansIncludedIds) {
                        if ($planId -notin $skuApplicablePlans)  { continue }
        
                        [pscustomobject][ordered]@{
                            "PlanName" = $planIdHashTable[$planId].DisplayName
                            "PlanId" = $planId
                            "State" = if ($disabledPlans -contains $planId) { "Off" } else { "On" }
                            # Yes, I repeat this info here and below. That's coz I might be viewing the plans only via -ShowPlansOnly
                            "SkuName" = $skuIdHashTable[$skuAssignedToObject].DisplayName
                            "SkuId" = $skuAssignedToObject
                        }
                    }

                    if ($SortPlansByState) {
                        $planStates | Sort-Object -Descending -Property { $_.State } | 
                            # Don't allow any selections
                            Out-ConsoleGridView -Title $skuIdHashTable[$skuAssignedToObject].DisplayName -OutputMode None
    
                    } else {
                        $planStates | Sort-Object -Descending -Property { $_.PlanName } | 
                            # Don't allow any selections
                            Out-ConsoleGridView -Title $skuIdHashTable[$skuAssignedToObject].DisplayName -OutputMode None
                    }
        
                } 
            }

            # Output the selections too just in case. Easy to copy paste for other cmdlets.
            # When doing this add the group or user name; remove the plans.
            foreach ($selection in $userSelections) {
                if ($PSCmdlet.ParameterSetName -eq "User") {
                    $selection | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $UserPrincipalName

                } else {
                    $selection | Add-Member -MemberType NoteProperty -Name "GroupName" -Value $GroupName
                }

                $selection | Select-Object -Property * -ExcludeProperty "Plans"
            }
        }
    }
}

function Update-MgAssignedLicensePlans {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "Group-SkuName")]
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "Group-SkuId")]
        [string]$GroupName,

        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "User-SkuName")]
        [Parameter(Position=1,Mandatory=$false,ParameterSetName = "User-SkuId")]
        [Alias("UPN")]
        [string]$UserPrincipalName,

        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "Group-SkuName")]
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "User-SkuName")]
        [ArgumentCompleter( { $skuNameHashTable.Keys | Sort-Object })]
        [string]$SkuName,

        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "Group-SkuId")]
        [Parameter(Position=1,Mandatory=$false,ParameterSetName = "User-SkuId")]
        [string]$SkuId,

        [Switch]$SortPlansByState
    )

    <#
    .DESCRIPTION
    SHow the license **plans** assigned to a group or user. You can select one or more by pressing the SPACE key and then ENTER to toggle the state (disable the plan if enabled; enable the plan if disabled).

    .PARAMETER GroupName
    The Group you'd like to see the license assignments of. 

    .PARAMETER UserPrincipalName
    The User you'd like to see the license assignments of. 

    .PARAMETER SkuName
    The license SKU Name (assigned to the user or group) whose plan details you wish to modify.

    You can enter the SKU Id too instead.

    .PARAMETER SkuId
    The license SKU Id (assigned to the user or group) whose plan details you wish to modify. 

    You can enter the SKU Name too instead.

    .PARAMETER SortPlansByState
    By default plans are sorted alphabetically. Use this to sort them by On/ Off state. 
    #>

    begin {
        $subscribedSkuIdHashTable = $script:subscribedSkuIdHashTable
        $subscribedSkuNameHashTable = $script:subscribedSkuNameHashTable
        $skuNameHashTable = $script:skuNameHashTable
        $skuIdHashTable = $script:skuIdHashTable
        $planNameHashTable = $script:planNameHashTable
        $planIdHashTable = $script:planIdHashTable

        if ($subscribedSkuIdHashTable.Count -eq 0 -or $subscribedSkuNameHashTable -eq 0 -or 
                $skuNameHashTable -eq 0 -or $skuIdHashTable -eq 0 -or 
                $planNameHashTable  -eq 0 -or $planIdHashTable -eq 0) {
            
            Update-MgLicensingData

            $subscribedSkuIdHashTable = $script:subscribedSkuIdHashTable
            $subscribedSkuNameHashTable = $script:subscribedSkuNameHashTable
            $skuNameHashTable = $script:skuNameHashTable
            $skuIdHashTable = $script:skuIdHashTable
            $planNameHashTable = $script:planNameHashTable
            $planIdHashTable = $script:planIdHashTable
        }


        if ($PSCmdlet.ParameterSetName -match "Group") {
            try {
                $groupObj = Get-MgGroup -Filter "DisplayName eq '$GroupName'" -Property assignedLicenses,Id,LicenseAssignmentStates -ErrorAction Stop
            
                if ($null -eq $groupObj) {
                    Write-Output "⛔ Couldn't find $GroupName"
                }
            
            } catch {
                Write-Output "⛔ Error searching for ${GroupName} - $($_.Exception.Message)"
            }

            $assignedLicenses = $groupObj.AssignedLicenses

            $assignmentPaths = @{}
            foreach ($skuIdTemp in $assignedLicenses.SkuId) {
                $assignmentPaths[$skuIdTemp] = "Direct"
            }
        }

        if ($PSCmdlet.ParameterSetName -match "User") {
            try {
                $userObj = Get-MgUser -Filter "UserPrincipalName eq '$UserPrincipalName'" -Property assignedLicenses,Id,LicenseAssignmentStates -ErrorAction Stop
            
                if ($null -eq $userObj) {
                    Write-Output "⛔ Couldn't find $GroupName"
                }
            
            } catch {
                Write-Output "⛔ Error searching for ${GroupName} - $($_.Exception.Message)"
            }

            $assignedLicenses = $userObj.AssignedLicenses

            $assignmentPaths = @{}
            foreach ($skuAssignment in $userObj.LicenseAssignmentStates) {
                $assignmentPaths[$($skuAssignment.SkuId)] = if ($skuAssignment.AssignedByGroup) { "Group" } else { "Direct" }
            }
        }

        if ($PSCmdlet.ParameterSetName -match "SkuName") {
            $SkuId = $skuNameHashTable[$SKuName].SkuId
        }
    }

    process {
        if ($assignmentPaths[$SkuId] -eq "Group") {
            Write-Output "🚫 Unable to make changes as this is a group assignment"
            return
        }

        $userSelections = $assignedLicenses | Where-Object { $_.SkuId -eq "$SkuId" } | ForEach-Object {
            # NOTE: There will only be one object, so the ForEach-Object is kind of redundant. Keeping it this way to keep the code similar to others.
            
            # Again, same. Using this variable to keep it similar to the other code. 
            $skuAssignedToObject = $_.SkuId
            # The plans that are disabled for this license assignments
            $disabledPlans = $_.DisabledPlans

            # All the plans that are actually available for this license SKU
            $skuApplicablePlans = ($subscribedSkuIdHashTable[$skuAssignedToObject].PlansIncludedFriendlyName | Where-Object { $_.AppliesTo -eq "User" }).ServicePlanId

            foreach ($planId in $skuIdHashTable[$skuAssignedToObject].PlansIncludedIds) {
                if ($planId -notin $skuApplicablePlans)  { continue }

                [pscustomobject][ordered]@{
                    "PlanName" = $planIdHashTable[$planId].DisplayName
                    "PlanId" = $planId
                    "State" = if ($disabledPlans -contains $planId) { "Off" } else { "On" }
                    # Yes, I repeat this info here and below. I add it here coz it makes it easier to pass this info to other cmdlets later.
                    "SkuName" = $skuIdHashTable[$skuAssignedToObject].DisplayName
                    "SkuId" = $skuAssignedToObject
                }
            }
        } | Sort-Object -Descending -Property { if ($SortPlansByState) { $_.State } else { $_.PlanName } } | Out-ConsoleGridView -Title $skuIdHashTable[$SkuId].DisplayName

        if ($userSelections.Count -ne 0) {
            Write-Output "Please confirm the following actions:"

            foreach ($selection in $userSelections) {
                $planName = $planIdHashTable[$selection.PlanId].DisplayName

                if ($selection.State -eq "Off") { 
                    $currentState = "Disabled"
                    $newState = "Enabled" 

                    # Need to turn this on - so we remove it from the disabled plans
                    $disabledPlans = @($disabledPlans | Where-Object { $_ -ne $selection.PlanId })

                    Write-Host -ForegroundColor Green "$planName | $currentState => $newState"

                } else {
                    $currentState = "Enabled" 
                    $newState = "Disabled"

                    if ($disabledPlans.Count -eq 0) {
                        $disabledPlans = @($selection.PlanId)
                    } else {
                        $disabledPlans = @($disabledPlans + $selection.PlanId)
                    }

                    Write-Host -ForegroundColor Red "$planName | $currentState => $newState"
                }
            }

            do {
                $confirmation = Read-Host "Ok to proceed? [y/n]"
            } while ($confirmation -notin "y","n")

            if ($confirmation -eq "n") {
                Write-Output "Not actioning any changes"
                return
            }

            # https://learn.microsoft.com/en-us/graph/api/group-assignlicense?view=graph-rest-1.0&tabs=powershell
            $params = @{
                addLicenses = @(
                    @{
                        disabledPlans = $disabledPlans
                        skuId = $SkuId
                    }
                )
                removeLicenses = @()
            }

            try {
                if ($PSCmdlet.ParameterSetName -match "Group") {
                    Set-MgGroupLicense -GroupId $groupObj.Id -BodyParameter $params -ErrorAction Stop | Out-Null
                }
        
                if ($PSCmdlet.ParameterSetName -match "User") {
                    Set-MgUserLicense -UserId $userObj.Id -BodyParameter $params -ErrorAction Stop | Out-Null
                }

                Write-Output "✔ All done!"
            } catch {
                Write-Output "⛔ Something went wrong: $($_.Exception.Message)"
            }
        }
    }   
}