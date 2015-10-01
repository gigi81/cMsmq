<#
Author  : Serge Nikalaichyk (https://www.linkedin.com/in/nikalaichyk)
Version : 1.0.0
Date    : 2015-09-30
#>


function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateSet('Absent', 'Present')]
        [System.String]
        $Ensure = 'Present',

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Name,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Principal,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $AccessRights = @('GenericRead')
    )
    begin
    {
        try
        {
            $Service = Get-Service -Name MSMQ -ErrorAction Stop

            if ($Service.Status -ne 'Running')
            {
                throw "Please ensure that the Message Queuing (MSMQ) service is running."
            }
        }
        catch
        {
            throw $_.Exception.Message
        }

        Import-cMsmqType
    }
    process
    {
        $QueuePath = '.\private$\{0}' -f $Name

        if (-not [System.Messaging.MessageQueue]::Exists($QueuePath))
        {
            Write-Error -Message "Queue '$Name' could not be found at the specified path: '$QueuePath'."

            return
        }

        $CurrentPermission = Get-cMsmqQueuePermission -Name $Name -Principal $Principal -ErrorAction SilentlyContinue

        if ($CurrentPermission)
        {
            "An existing permission entry was found for principal '$Principal' on queue '$Name':",
            "Currently assigned permissions: '$CurrentPermission'." |
            Write-Verbose

            if ($Ensure -eq 'Present')
            {
                $DesiredPermission = [System.Messaging.MessageQueueAccessRights]$AccessRights

                if ($CurrentPermission -eq $DesiredPermission)
                {
                    $EnsureResult = 'Present'
                }
                else
                {
                    $EnsureResult = 'Absent'
                }
            }
            elseif ($Ensure -eq 'Absent')
            {
                $EnsureResult = 'Present'
            }

            $AccessRightsResult = @($CurrentPermission.ToString() -split ', ')
        }
        else
        {
            Write-Verbose -Message "There is no existing permission entry found for principal '$Principal' on queue '$Name'."

            $EnsureResult = 'Absent'
            $AccessRightsResult = @()
        }

        $ReturnValue = @{
                Ensure = $EnsureResult
                Name = $Name
                Principal = $Principal
                AccessRights = $AccessRightsResult
            }

        return $ReturnValue
    }
}


function Set-TargetResource
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateSet('Absent', 'Present')]
        [System.String]
        $Ensure = 'Present',

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Name,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Principal,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $AccessRights = @('GenericRead')
    )
    begin
    {
        $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    process
    {
        if (-not $PSCmdlet.ShouldProcess($Name))
        {
            return
        }

        $QueuePath = '.\private$\{0}' -f $Name

        if (-not [System.Messaging.MessageQueue]::Exists($QueuePath))
        {
            Write-Error -Message "Queue '$Name' could not be found at the specified path: '$QueuePath'."

            return
        }

        $Queue = New-Object -TypeName System.Messaging.MessageQueue -ArgumentList $QueuePath

        $DesiredPermission = [System.Messaging.MessageQueueAccessRights]$AccessRights

        Write-Verbose -Message "Testing if the current user has the permission necessary to perform the operation."

        $CurrentUserPermission = Get-cMsmqQueuePermission -Name $Name -Principal $CurrentUser -ErrorAction SilentlyContinue
        $PermissionToTest = [System.Messaging.MessageQueueAccessRights]::ChangeQueuePermissions

        if (-not $CurrentUserPermission -or -not $CurrentUserPermission.HasFlag($PermissionToTest))
        {
            "User '{0}' does not have the '{1}' permission on queue '{2}'." -f $CurrentUser, $PermissionToTest, $Name |
            Write-Verbose

            Reset-cMsmqQueueSecurity -Name $Name -Confirm:$false -Verbose:$VerbosePreference
        }

        if ($Ensure -eq 'Absent')
        {
            Write-Verbose -Message "Revoking all existing permissions for principal '$Principal' on queue '$Name'."

            $Queue.SetPermissions($Principal, $DesiredPermission, [System.Messaging.AccessControlEntryType]::Revoke)
        }
        else
        {
            Write-Verbose -Message "Setting permissions for principal '$Principal' on queue '$Name'."

            $Queue.SetPermissions($Principal, $DesiredPermission, [System.Messaging.AccessControlEntryType]::Set)
        }
    }
}


function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateSet('Absent', 'Present')]
        [System.String]
        $Ensure = 'Present',

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Name,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Principal,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $AccessRights = @('GenericRead')
    )

    $PSBoundParameters.GetEnumerator() |
    ForEach-Object -Begin {
        $Width = $PSBoundParameters.Keys.Length | Sort-Object -Descending | Select-Object -First 1
    } -Process {
        "{0,-$($Width)} : '{1}'" -f $_.Key, ($_.Value -join ', ') |
        Write-Verbose
    }

    $TargetResource = Get-TargetResource @PSBoundParameters

    $InDesiredState = $Ensure -eq $TargetResource.Ensure

    if ($InDesiredState -eq $true)
    {
        Write-Verbose -Message "The target resource is already in the desired state. No action is required."
    }
    else
    {
        Write-Verbose -Message "The target resource is not in the desired state."
    }

    return $InDesiredState

}


Export-ModuleMember -Function Get-TargetResource, Set-TargetResource, Test-TargetResource


#region Helper Functions

function Import-cMsmqType
{
    <#
    .SYNOPSIS
        Imports custom and native MSMQ types.
    .DESCRIPTION
        The Import-cMsmqType function imports custom and native MSMQ types.
    #>

    $DllFilePath = Split-Path -Path $PSScriptRoot -Parent |
        Split-Path -Parent |
        Join-Path -ChildPath 'cMsmq.dll'

    if ([System.AppDomain]::CurrentDomain.GetAssemblies().Location -notcontains $DllFilePath)
    {
        Add-Type -Path $DllFilePath -ErrorAction Stop
    }

    if ([System.AppDomain]::CurrentDomain.GetAssemblies().ManifestModule.Name -notcontains 'System.Messaging.dll')
    {
        Add-Type -AssemblyName System.Messaging -ErrorAction Stop
    }
}

Import-cMsmqType


function Get-cMsmqQueuePermission
{
    <#
    .SYNOPSIS
        Gets the access rights of the specified principal on the specified private MSMQ queue.
    .DESCRIPTION
        The Get-cMsmqQueuePermission function gets the access rights that have been granted 
        to the specified security principal on the specified private MSMQ queue.
    .PARAMETER Name
        Specifies the name of the queue.
    .PARAMETER Principal
        Specifies the identity of the principal.
    #>
    [CmdletBinding()]
    [OutputType([System.Messaging.MessageQueueAccessRights])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Principal
    )
    begin
    {
        Import-MsmqType
    }
    process
    {
        try
        {
            Write-Verbose -Message "Getting permissions for principal '$Principal' on queue '$Name'."

            $AccessMask = [cMsmq.Security]::GetAccessMask($Name, $Principal)
            $OutputObject = [System.Messaging.MessageQueueAccessRights]$AccessMask.value__

            return $OutputObject
        }
        catch
        {
            Write-Error -Message $_.Exception.Message

            return
        }
    }
}


function Reset-cMsmqQueueSecurity
{
    <#
    .SYNOPSIS
        Resets the security settings on the specified private MSMQ queue.
    .DESCRIPTION
        The Reset-cMsmqQueueSecurity function performs the following actions:
            - Grants ownership of the queue to the SYSTEM account (DSC runs as SYSTEM);
            - Resets the permission list to the operating system's default values.
    .PARAMETER Name
        Specifies the name of the queue.
    #>
    [CmdletBinding(ConfirmImpact = 'High', SupportsShouldProcess = $true)]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Name
    )
    begin
    {
        Import-cMsmqType

        $DefaultSecurity = 'Security=010007801c0000002800000000000000140000000200080000000000' +
            '010100000000000512000000010500000000000515000000e611610036157811027bc60001020000'
    }
    process
    {
        if (-not $PSCmdlet.ShouldProcess($Name, 'Reset Queue Security'))
        {
            return
        }

        $QueuePath = '.\private$\{0}' -f $Name

        $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $QueueOwner = [cMsmq.Security]::GetOwner($Name)

        Write-Verbose -Message "Queue Owner : '$QueueOwner'"

        if ($CurrentUser -ne $QueueOwner)
        {
            Write-Verbose -Message "Taking ownership of queue '$Name'."

            $FilePath = Get-ChildItem -Path "$Env:SystemRoot\System32\msmq\storage\lqs" -Force |
                Select-String -Pattern "Label=private`$\$($Name)" -SimpleMatch |
                Select-Object -ExpandProperty Path

            if (-not $FilePath)
            {
                Write-Error -Message "Could not find a corresponding .ini file for queue '$Name'."

                return
            }

            (Get-Content -Path $FilePath) |
            ForEach-Object {$_ -replace '^Security=.+', $DefaultSecurity} |
            Set-Content -Path $FilePath
        }

        Write-Verbose -Message "Resetting permissions on queue '$Name'."

        $Queue = New-Object -TypeName System.Messaging.MessageQueue
        $Queue.Path = $QueuePath
        $Queue.ResetPermissions()
    }
}


#endregion

