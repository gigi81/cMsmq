#requires -Version 4.0

Import-Module cMsmq

function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster
    )
    
    $cMsmqQueue = Get-cMsmqQueue -Name $Name -Cluster $Cluster -ErrorAction SilentlyContinue

    if ($cMsmqQueue)
    {
        Write-Verbose -Message "Queue '$Name' was found."
        $EnsureResult = 'Present'
    }
    else
    {
        Write-Verbose -Message "Queue '$Name' could not be found."
        $EnsureResult = 'Absent'
    }

    $ReturnValue = @{
        Ensure        = $EnsureResult
        Name          = $Name
        Transactional = $cMsmqQueue.Transactional
        Authenticate  = $cMsmqQueue.Authenticate
        Journaling    = $cMsmqQueue.Journaling
        JournalQuota  = $cMsmqQueue.JournalQuota
        Label         = $cMsmqQueue.Label
        PrivacyLevel  = $cMsmqQueue.PrivacyLevel
        QueueQuota    = $cMsmqQueue.QueueQuota
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

        [Parameter(Mandatory = $false)]
        [Boolean]
        $Transactional = $false,

        [Parameter(Mandatory = $false)]
        [Boolean]
        $Authenticate = $false,

        [Parameter(Mandatory = $false)]
        [Boolean]
        $Journaling = $false,

        [Parameter(Mandatory = $false)]
        [UInt32]
        $JournalQuota = [UInt32]::MaxValue,

        [Parameter(Mandatory = $false)]
        [String]
        $Label = $null,

        [Parameter(Mandatory = $false)]
        [ValidateSet('None', 'Optional', 'Body')]
        [String]
        $PrivacyLevel = 'Optional',

        [Parameter(Mandatory = $false)]
        [UInt32]
        $QueueQuota = [UInt32]::MaxValue,
        
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

    $TargetResource = Get-TargetResource -Name $Name -Cluster $Cluster

    if ($Ensure -eq 'Absent')
    {
        if ($TargetResource.Ensure -eq 'Absent')
        {
            $InDesiredState = $true
        }
        else
        {
            $InDesiredState = $false
        }
    }
    else
    {
        if ($TargetResource.Ensure -eq 'Absent')
        {
            $InDesiredState = $false
        }
        else
        {
            $InDesiredState = $true

            if ($PSBoundParameters.ContainsKey('Transactional'))
            {
                if ($TargetResource.Transactional -ne $Transactional)
                {
                    $InDesiredState = $false

                    if ($TargetResource.Transactional -eq $true)
                    {
                        $CurrentQueueTypeString = 'transactional'
                    }
                    else
                    {
                        $CurrentQueueTypeString = 'non-transactional'
                    }

                    if ($Transactional -eq $true)
                    {
                        $DesiredQueueTypeString = 'transactional'
                    }
                    else
                    {
                        $DesiredQueueTypeString = 'non-transactional'
                    }

                    $ErrorMessage = "Queue '{0}' is {1} and cannot be converted to {2}." -f $Name, $CurrentQueueTypeString, $DesiredQueueTypeString

                    throw $ErrorMessage
                }
            }

            $PSBoundParameters.GetEnumerator() |
            Where-Object {$_.Key -in @('Authenticate', 'Journaling', 'JournalQuota', 'Label', 'PrivacyLevel', 'QueueQuota')} |
            ForEach-Object {

                $PropertyName = $_.Key

                if ($TargetResource."$PropertyName" -cne $_.Value)
                {
                    $InDesiredState = $false

                    "Property '{0}': Current value '{1}'; Desired value: '{2}'." -f $PropertyName, $TargetResource."$PropertyName", $_.Value |
                    Write-Verbose
                }
            }
        }
    }

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

        [Parameter(Mandatory = $false)]
        [Boolean]
        $Transactional = $false,

        [Parameter(Mandatory = $false)]
        [Boolean]
        $Authenticate = $false,

        [Parameter(Mandatory = $false)]
        [Boolean]
        $Journaling = $false,

        [Parameter(Mandatory = $false)]
        [UInt32]
        $JournalQuota = [UInt32]::MaxValue,

        [Parameter(Mandatory = $false)]
        [String]
        $Label = $null,

        [Parameter(Mandatory = $false)]
        [ValidateSet('None', 'Optional', 'Body')]
        [String]
        $PrivacyLevel = 'Optional',

        [Parameter(Mandatory = $false)]
        [UInt32]
        $QueueQuota = [UInt32]::MaxValue,
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster
    )

    if (-not $PSCmdlet.ShouldProcess($Name))
    {
        return
    }

    if ($Ensure -eq 'Absent')
    {
        Test-cMsmqPermissions -Name $Name -Cluster $Cluster -Permission [System.Messaging.MessageQueueAccessRights]::DeleteQueue
    
        $PSBoundParameters.GetEnumerator() |
        Where-Object {$_.Key -in (Get-Command -Name Remove-cMsmqQueue).Parameters.Keys} |
        ForEach-Object -Begin {$RemoveParameters = @{}} -Process {$RemoveParameters.Add($_.Key, $_.Value)}

        Remove-cMsmqQueue @RemoveParameters -Confirm:$false
    }
    else
    {
        $TargetResource = Get-TargetResource -Name $Name -Cluster $Cluster

        if ($TargetResource.Ensure -eq 'Absent')
        {
            $PSBoundParameters.GetEnumerator() |
            Where-Object {$_.Key -in (Get-Command -Name New-cMsmqQueue).Parameters.Keys} |
            ForEach-Object -Begin {$NewParameters = @{}} -Process {$NewParameters.Add($_.Key, $_.Value)}

            New-cMsmqQueue @NewParameters
        }
        else
        {
            Test-cMsmqPermissions -Name $Name -Cluster $Cluster -Permission [System.Messaging.MessageQueueAccessRights]::SetQueueProperties

            $PSBoundParameters.GetEnumerator() |
            Where-Object {$_.Key -in (Get-Command -Name Set-cMsmqQueue).Parameters.Keys} |
            ForEach-Object -Begin {$SetParameters = @{}} -Process {$SetParameters.Add($_.Key, $_.Value)}

            Set-cMsmqQueue @SetParameters
        }
    }
}
