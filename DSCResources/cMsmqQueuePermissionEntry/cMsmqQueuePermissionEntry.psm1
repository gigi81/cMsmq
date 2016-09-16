#requires -Version 4.0

Import-Module cMsmq

function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([Hashtable])]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateSet('Absent', 'Present')]
        [String]
        $Ensure = 'Present',

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Principal,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $AccessRights = @('GenericRead'),
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster
    )
    
    if ($null -eq (Get-cMsmqQueue -Name $Name -Cluster $Cluster -ErrorAction SilentlyContinue))
    {
        return
    }

    $CurrentPermission = Get-cMsmqQueuePermission -Name $Name -Cluster $Cluster -Principal $Principal -ErrorAction SilentlyContinue

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
        Ensure       = $EnsureResult
        Name         = $Name
        Principal    = $Principal
        AccessRights = $AccessRightsResult
    }

    return $ReturnValue
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([Boolean])]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateSet('Absent', 'Present')]
        [String]
        $Ensure = 'Present',

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Principal,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $AccessRights = @('GenericRead'),
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster
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

function Set-TargetResource
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateSet('Absent', 'Present')]
        [String]
        $Ensure = 'Present',

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Principal,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $AccessRights = @('GenericRead'),
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster
    )

    if (-not $PSCmdlet.ShouldProcess($Name))
    {
        return
    }

    $revoke = if ($Ensure -eq 'Absent') { $true } else { $false }
    
    #test if the user DSC is running on (should be SYSTEM) has permissions to change the queue permissions
    if(-not (Test-cMsmqPermissions -Name $Name -Cluster $Cluster -Permission 'ChangeQueuePermissions'))
    {
        #try to reset the permissions by changing the permissions on the queue file
        Reset-cMsmqQueueSecurity -Name $Name -Cluster $Cluster -Confirm:$false
    }
    
    Set-cMsmqPermissions -Name $Name -Cluster $Cluster -Principal $Principal -Permission $AccessRights -Revoke:$revoke
}

Export-ModuleMember -Function *-TargetResource
