# Graph Licensing
Source code a module I created for managing Entra ID licensing via Graph. It isn't complete in any way. 

You can install the module [from PowerShell Gallery](https://www.powershellgallery.com/packages/Graph.Licensing/).

Tested on macOS, Linux, and Windows with PowerShell 7.4. Should work on other OSes too with PowerShell 7.x and above. 

## Pre-requisite modules
- "Microsoft.Graph.Authentication"
- "Microsoft.Graph.Groups"
- "Microsoft.Graph.Users"
- "Microsoft.Graph.Identity.DirectoryManagement"
- "Microsoft.PowerShell.ConsoleGuiTools"

## Screenshots

`Get-MgAvailableLicenses`

![image-20240922013047361](./assets/image-20240922013047361.png)

Selecting one or more of the licenses above (use `SPACE` key to make selections, followed by the `ENTER` key to confirm) will go through each of the selections and first show the users directly assigned to it. 

![image-20240922013145072](./assets/image-20240922013145072.png)

Followed by the groups directly assigned to it. 

![image-20240922013401551](./assets/image-20240922013401551.png)

In either screen (users or groups), selecting one or more entries (use `SPACE` key to make selections, followed by the `ENTER` key to confirm) will then go through each of the selections and show the plans. 

![image-20240922013606905](./assets/image-20240922013606905.png)

There are cmdlets to directly query a group or user rather than go the route above. 

It is also possible to modify the plan assignments. 

`Update-MgAssignedLicensePlans -UserPrincipalName $UPN -SkuName $SkuName` or `Update-MgAssignedLicensePlans -GroupName $GroupName -SkuName $SkuName`

![image-20240922014348917](./assets/image-20240922014348917.png)

After making the selections:

![image-20240922014422154](./assets/image-20240922014422154.png)

Confirm. Via `Get-MgAssignedLicenses -GroupName $GroupName -SkuName $SkuName` or `Get-MgAssignedLicenses -UserPrincipalName $UPN -SkuName $SkuName`

![image-20240922015413294](./assets/image-20240922015413294.png)
