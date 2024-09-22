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
            throw "Missing the required permissions. Are you connected to Graph and have the right permissions? For example: Connect-MgGraph -Scopes 'Directory.Read.All'"
        }
    }

    process {
        $subscribedSkuIdHashTable = @{}
        $subscribedSkuNameHashTable = @{}
        Write-Output "〰 Creating a list of available licenses from Graph"
        Get-MgSubscribedSku -All | Where-Object { $_.AppliesTo -eq "User" } | ForEach-Object {
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
        if ($subscribedSkuIdHashTable.Count -eq 0 -or $subscribedSkuNameHashTable -eq 0 -or 
                $skuNameHashTable -eq 0 -or $skuIdHashTable -eq 0 -or 
                $planNameHashTable  -eq 0 -or $planIdHashTable -eq 0) {

            throw "Missing the required info. Something went wrong."

        } else {
            $script:subscribedSkuIdHashTable = $subscribedSkuIdHashTable
            $script:subscribedSkuNameHashTable = $subscribedSkuNameHashTable
            $script:skuNameHashTable = $skuNameHashTable
            $script:skuIdHashTable = $skuIdHashTable
            $script:planNameHashTable = $planNameHashTable
            $script:planIdHashTable = $planIdHashTable
        
            Write-Output "✔ All done!"
        }
    }
}

function Get-MgAssignedLicenses {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "Group-SkuName")]
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "Group-SkuId")]
        [string]$GroupName,

        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "User-SkuName")]
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "User-SkuId")]
        [Alias("UPN")]
        [string]$UserPrincipalName,

        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "Group-SkuName")]
        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "User-SkuName")]
        [ArgumentCompleter( { $skuNameHashTable.Keys | Sort-Object })]
        [string]$SkuName,

        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "Group-SkuId")]
        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "User-SkuId")]
        [string]$SkuId,
        
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

    .PARAMETER SkuId
    Show only this SkuId. Useful with the -ShowPlansOnly switch. 

    .PARAMETER SkuName
    Show only this SkuName. Useful with the -ShowPlansOnly switch. 
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

        if ($subscribedSkuIdHashTable.Count -eq 0 -or $subscribedSkuNameHashTable -eq 0 -or 
                $skuNameHashTable -eq 0 -or $skuIdHashTable -eq 0 -or 
                $planNameHashTable  -eq 0 -or $planIdHashTable -eq 0) {

            throw "Missing the required info. Did Update-MgLicensingData successfully?"

        }

        if ($PSCmdlet.ParameterSetName -match "Group") {
            try {
                $groupObj = Get-MgGroup -Filter "DisplayName eq '$GroupName'" -Property assignedLicenses,Id,LicenseAssignmentStates -ErrorAction Stop
            
                if ($null -eq $groupObj) {
                    throw "Couldn't find group '$GroupName'"
                }
            
            } catch {
                throw "Error searching for group '${GroupName}' - $($_.Exception.Message)"
            }

            $assignedLicenses = $groupObj.AssignedLicenses

            $assignmentPaths = @{}
            foreach ($skuIdTemp in $assignedLicenses.SkuId) {
                $assignmentPaths[$skuIdTemp] = "Direct"
            }

            $targetSnippet = "Group '${GroupName}'"
        }

        if ($PSCmdlet.ParameterSetName -match "User") {
            try {
                $userObj = Get-MgUser -Filter "UserPrincipalName eq '$UserPrincipalName'" -Property assignedLicenses,Id,LicenseAssignmentStates -ErrorAction Stop
            
                if ($null -eq $userObj) {
                    throw "Couldn't find user '$UserPrincipalName'"
                }
            
            } catch {
                throw " Error searching for user '${UserPrincipalName}' - $($_.Exception.Message)"
            }

            $assignedLicenses = $userObj.AssignedLicenses

            $assignmentPaths = @{}
            foreach ($skuAssignment in $userObj.LicenseAssignmentStates) {
                $assignmentPaths[$($skuAssignment.SkuId)] = if ($skuAssignment.AssignedByGroup) { "Group" } else { "Direct" }
            }

            $targetSnippet = "User '${UserPrincipalName}'"
        }

        if ($PSCmdlet.ParameterSetName -match "SkuName") {
            $SkuId = $skuNameHashTable[$SKuName].SkuId
        }

        if ($PSCmdlet.ParameterSetName -match "SkuId") {
            $SkuName = $skuIdHashTable[$SkuId].DisplayName
        }
    }

    process {
        # If a SkuId is given focus only on that license
        if ($SkuId) {
            $assignedLicenses = $assignedLicenses | Where-Object { $_.SkuId -eq "$SkuId" }
        }

        # If we have to show only the plans, it's straight forward...
        if ($ShowPlansOnly) {
            # If a SkuId is given modify the output title accordingly
            # I don't need to do it this way if I am not doing -ShowPlansOnly coz then I set the title within the loop and that picks up the SkuName. 
            # I use a different variable $targetSnippet2 there coz I need to keep some parts the same across the loop.
            if ($SkuId) {
                $SkuName = $skuIdHashTable[$SkuId].DisplayName
                $targetSnippet = $targetSnippet + " for license '$SkuName'"
            }

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

            } | Out-ConsoleGridView -OutputMode None -Title "Plans assigned to $targetSnippet"

        } else {
        # If we are showing licenses to, I want to capture the user selections and then loop into a plan only view for each of their selections :)
            $userSelections = $assignedLicenses | ForEach-Object {
                $skuAssignedToObject = $_.SkuId
                # The plans that are disabled for this license assignments
                $disabledPlans = $_.DisabledPlans

                # If a SkuId is given modify the output title accordingly
                if ($SkuId) {
                    $SkuName = $skuIdHashTable[$SkuId].DisplayName
                    $targetSnippet = $targetSnippet + " for license '$SkuName'"
                }
    
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
    
            } | Out-ConsoleGridView -Title "Licenses assigned to $targetSnippet" 
            # Note I use $targetSnippet here as it's NOT Sku specific

            # If the user made selection, expand into the plans
            foreach ($selection in $userSelections) {
                $assignedLicenses | Where-Object { $_.SkuId -eq $selection.SkuId } | ForEach-Object {
                    $skuAssignedToObject = $_.SkuId
                    # The plans that are disabled for this license assignments
                    $disabledPlans = $_.DisabledPlans

                    # This is the text I show in the title. In case -SkuId is specified, I don't need to do a special case as the loop will only run once.
                    # $targetSnippet never changes, but $targetSnippet2 keeps varies each time
                    $SkuName = $skuIdHashTable[$skuAssignedToObject].DisplayName
                    $targetSnippet2 = $targetSnippet + " for license '$SkuName'"
        
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
                            Out-ConsoleGridView -OutputMode None -Title "Plans assigned to $targetSnippet2"
    
                    } else {
                        $planStates | Sort-Object -Descending -Property { $_.PlanName } | 
                            # Don't allow any selections
                            Out-ConsoleGridView -OutputMode None -Title "Plans assigned to $targetSnippet2"
                    }
                } 
            }

            # Output the selections too just in case. Easy to copy paste for other cmdlets.
            # When doing this add the group or user name; remove the plans.
            foreach ($selection in $userSelections) {
                if ($PSCmdlet.ParameterSetName -match "User") {
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
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "User-SkuId")]
        [Alias("UPN")]
        [string]$UserPrincipalName,

        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "Group-SkuName")]
        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "User-SkuName")]
        [ArgumentCompleter( { $skuNameHashTable.Keys | Sort-Object })]
        [string]$SkuName,

        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "Group-SkuId")]
        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "User-SkuId")]
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

        if ($subscribedSkuIdHashTable.Count -eq 0 -or $subscribedSkuNameHashTable -eq 0 -or 
                $skuNameHashTable -eq 0 -or $skuIdHashTable -eq 0 -or 
                $planNameHashTable  -eq 0 -or $planIdHashTable -eq 0) {

            throw "Missing the required info. Did Update-MgLicensingData successfully?"

        }

        if ($PSCmdlet.ParameterSetName -match "Group") {
            try {
                $groupObj = Get-MgGroup -Filter "DisplayName eq '$GroupName'" -Property assignedLicenses,Id,LicenseAssignmentStates -ErrorAction Stop
            
                if ($null -eq $groupObj) {
                    throw "Couldn't find group '$GroupName'"
                }
            
            } catch {
                throw "Error searching for group '${GroupName}' - $($_.Exception.Message)"
            }

            $assignedLicenses = $groupObj.AssignedLicenses

            $assignmentPaths = @{}
            foreach ($skuIdTemp in $assignedLicenses.SkuId) {
                $assignmentPaths[$skuIdTemp] = "Direct"
            }

            $targetSnippet = "Group '${GroupName}'"
        }

        if ($PSCmdlet.ParameterSetName -match "User") {
            try {
                $userObj = Get-MgUser -Filter "UserPrincipalName eq '$UserPrincipalName'" -Property assignedLicenses,Id,LicenseAssignmentStates -ErrorAction Stop
            
                if ($null -eq $userObj) {
                    throw "Couldn't find user '$UserPrincipalName'"
                }
            
            } catch {
                throw "Error searching for user '${UserPrincipalName}' - $($_.Exception.Message)"
            }

            $assignedLicenses = $userObj.AssignedLicenses

            $assignmentPaths = @{}
            foreach ($skuAssignment in $userObj.LicenseAssignmentStates) {
                $assignmentPaths[$($skuAssignment.SkuId)] = if ($skuAssignment.AssignedByGroup) { "Group" } else { "Direct" }
            }

            $targetSnippet = "User '${UserPrincipalName}'"
        }

        if ($PSCmdlet.ParameterSetName -match "SkuName") {
            $SkuId = $skuNameHashTable[$SKuName].SkuId
        }

        if ($PSCmdlet.ParameterSetName -match "SkuId") {
            $SkuName = $skuIdHashTable[$SkuId].DisplayName
        }
    }

    process {
        if ($assignmentPaths[$SkuId] -eq "Group") {
            throw "Unable to make changes as this is a group assignment"
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
        } | Sort-Object -Descending -Property { if ($SortPlansByState) { $_.State } else { $_.PlanName } } | Out-ConsoleGridView -Title "Plans of license '$($skuIdHashTable[$SkuId].DisplayName)' assigned to $targetSnippet - select to toggle the state"

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
                throw "Something went wrong: $($_.Exception.Message)"
            }
        }
    }
}

function Add-MgAssignedLicense {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "Group-SkuName")]
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "Group-SkuId")]
        [string]$GroupName,

        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "User-SkuName")]
        [Parameter(Position=0,Mandatory=$true,ParameterSetName = "User-SkuId")]
        [Alias("UPN")]
        [string]$UserPrincipalName,

        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "Group-SkuName")]
        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "User-SkuName")]
        [ArgumentCompleter( { $skuNameHashTable.Keys | Sort-Object })]
        [string]$SkuName,

        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "Group-SkuId")]
        [Parameter(Position=1,Mandatory=$true,ParameterSetName = "User-SkuId")]
        [string]$SkuId,

        [Switch]$SortPlansByState
    )

    <#
    .DESCRIPTION
    Add (assign) a license SKU to a user or group. While adding you can select the plans too.

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

        if ($subscribedSkuIdHashTable.Count -eq 0 -or $subscribedSkuNameHashTable -eq 0 -or 
                $skuNameHashTable -eq 0 -or $skuIdHashTable -eq 0 -or 
                $planNameHashTable  -eq 0 -or $planIdHashTable -eq 0) {

            throw "Missing the required info. Did Update-MgLicensingData successfully?"

        }

        if ($PSCmdlet.ParameterSetName -match "Group") {
            try {
                $groupObj = Get-MgGroup -Filter "DisplayName eq '$GroupName'" -Property assignedLicenses,Id,LicenseAssignmentStates -ErrorAction Stop
            
                if ($null -eq $groupObj) {
                    throw "Couldn't find group '$GroupName'"
                }
            
            } catch {
                throw "Error searching for group '${GroupName}' - $($_.Exception.Message)"
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
                    throw "Couldn't find user '$UserPrincipalName'"
                }
            
            } catch {
                throw "Error searching for user '${UserPrincipalName}' - $($_.Exception.Message)"
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

        if ($PSCmdlet.ParameterSetName -match "SkuId") {
            $SkuName = $skuIdHashTable[$SkuId].DisplayName
        }
    }

    process {
        $disabledPlans = @()

        # Using this variable to keep it similar to the other code.
        $skuAssignedToObject = $SkuId

        # All the plans that are actually available for this license SKU
        $skuApplicablePlans = ($subscribedSkuIdHashTable[$skuAssignedToObject].PlansIncludedFriendlyName | Where-Object { $_.AppliesTo -eq "User" }).ServicePlanId

        $planStates = foreach ($planId in $skuIdHashTable[$skuAssignedToObject].PlansIncludedIds) {
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
        
        $userSelections = $planStates | Sort-Object -Descending -Property { if ($SortPlansByState) { $_.State } else { $_.PlanName } } | Out-ConsoleGridView -Title "Plans of license '$($skuIdHashTable[$SkuId].DisplayName)' - select plans you wish to disable" 

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

        } else {
            do {
                $confirmation = Read-Host "Please confirm you'd like to assign the '$SkuName' license SKU [y/n]"
            } while ($confirmation -notin "y","n")

            if ($confirmation -eq "n") {
                Write-Output "Not actioning any changes"
                return
            }
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
            throw "Something went wrong: $($_.Exception.Message)"
        }
    }
}

function Get-MgAvailableLicenses {
    <#
        .DESCRIPTION
        List the available licenses in the tenant.
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

        if ($subscribedSkuIdHashTable.Count -eq 0 -or $subscribedSkuNameHashTable -eq 0 -or 
                $skuNameHashTable -eq 0 -or $skuIdHashTable -eq 0 -or 
                $planNameHashTable  -eq 0 -or $planIdHashTable -eq 0) {

            throw "Missing the required info. Did Update-MgLicensingData successfully?"

        }
    }

    process {
        $availableLicenses = Get-MgSubscribedSku -All | Where-Object { $_.CapabilityStatus -eq "Enabled" -and $_.AppliesTo -eq "User" } | 
        Select-Object @{"Label" = "Name"; "Expression" = {$skuIdHashTable[$_.SkuId].DisplayName}}, @{"Label" = "Consumed"; "Expression" = { $_.ConsumedUnits }}, @{"Label" = "Available"; "Expression" = { $_.PrepaidUnits.Enabled }}, @{"Label" = "Warning"; "Expression" = { $_.PrepaidUnits.Warning }}, @{"Label" = "LockedOut"; "Expression" = { $_.PrepaidUnits.LockedOut }}, @{"Label" = "Suspended"; "Expression" = { $_.PrepaidUnits.Suspended }}, SkuId |
        Sort-Object -Property Name 
    
        $userSelections = $availableLicenses | Out-ConsoleGridView -Title "Available Licenses (Count: $($availableLicenses.Count))"
    
        foreach ($selection in $userSelections) {
            $SkuId = $selection.SkuId
            $SkuName = $skuIdHashTable[$SkuId].DisplayName
            $filterClause = "assignedLicenses/any(x:x/SkuId eq $SkuId)"
            $userSelections2 = @()
    
            try {
                $userSelections2 = Get-MgUser -All -Filter $filterClause -ConsistencyLevel Eventual -CountVariable userCount -ErrorAction Stop | 
                Select-Object UserPrincipalName, Id, @{"Label" = "SkuName"; Expression = { $SkuName }}, @{"Label" = "SkuId"; Expression = { $SkuId }} | 
                Out-ConsoleGridView -Title "Users assigned the '$SkuName' license (Count: $userCount)"
        
            } catch {
                throw "Error retrieving users assigned this license: $($_.Exception.Message)"
            }
    
            foreach ($selection2 in $userSelections2) {
                Get-MgAssignedLicenses -UserPrincipalName $selection2.UserPrincipalName -SkuId $selection2.SkuId -ShowPlansOnly
            }

            # re-initialize this
            $userSelections2 = @()
    
            try {
                $userSelections2 = Get-MgGroup -All -Filter $filterClause -ConsistencyLevel Eventual -CountVariable groupCount -ErrorAction Stop | 
                Select-Object DisplayName, Id, @{"Label" = "SkuName"; Expression = { $SkuName }}, @{"Label" = "SkuId"; Expression = { $SkuId }} | 
                Out-ConsoleGridView -Title "Groups assigned the '$SkuName' license (Count: $groupCount)"
        
            } catch {
                throw "Error retrieving groups assigned this license: $($_.Exception.Message)"
            }
    
            foreach ($selection2 in $userSelections2) {
                Get-MgAssignedLicenses -GroupName $selection2.DisplayName -SkuId $selection2.SkuId -ShowPlansOnly
            }
        }
    }
}