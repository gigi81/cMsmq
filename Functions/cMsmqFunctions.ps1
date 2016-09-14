
function Initialize-cMsmq
{
    <#
    .SYNOPSIS
        Initializes custom and native MSMQ types.
    .DESCRIPTION
        The Initialize-cMsmq function initializes custom and native MSMQ types.
    #>
    [CmdletBinding()]
    param(
        $Session
    )

    Invoke-Command -ScriptBlock {
        param(
            $Path
        )
        
        $dllFilePath = Join-Path -Path $Path -ChildPath 'cMsmq.dll'

        if ([AppDomain]::CurrentDomain.GetAssemblies().Location -notcontains $dllFilePath)
        {
            Add-Type -Path $dllFilePath -ErrorAction Stop
        }

        if ([AppDomain]::CurrentDomain.GetAssemblies().ManifestModule.Name -notcontains 'System.Messaging.dll')
        {
            Add-Type -AssemblyName System.Messaging -ErrorAction Stop
        }
    } -ArgumentList (Split-Path -Path $PSScriptRoot -Parent) -Session $Session
}

function New-cMsmqSession
{
    param(
        [String]
        $Cluster
    )
    
    $computerName = $Cluster
    if([String]::IsNullorEmpty($computerName))
    {
        $computerName = $env:computername
    }
    
    Write-Verbose "Creating cMsmq session on $computerName"
    $session = New-PSSession -ComputerName $computerName
    Initialize-cMsmq $session
    
    if(-Not [String]::IsNullorEmpty($Cluster))
    {
        Invoke-Command -ScriptBlock {
            param(
                $Cluster
            )
            
            # https://technet.microsoft.com/en-us/library/hh405007(v=vs.85).aspx
            $env:_Cluster_Network_Name_ = $Cluster
        } -Session $session -ArgumentList $Cluster
    }
    
    return $session
}

function Remove-cMsmqSession
{
    [CmdletBinding()]
    param(
        $Session
    )
    
    if($Session)
    {
        Write-Verbose "Removing cMsmq session"
        $Session | Remove-PSSession
    }
}

function Get-cMsmqQueue
{
    <#
    .SYNOPSIS
        Gets the specified private MSMQ queue by its name.
    .DESCRIPTION
        The Get-cMsmqQueue function gets the specified private MSMQ queue by its name.
    .PARAMETER Name
        Specifies the name of the queue.
    .PARAMETER Cluster
        Specifies the name of the cluster
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster
    )

    $session = New-cMsmqSession $Cluster

    try
    {
        Invoke-Command -ScriptBlock {
            param(
                $Name
            )
            
            $QueuePath = '.\private$\{0}' -f $Name
            
            if (-not [System.Messaging.MessageQueue]::Exists($QueuePath))
            {
                Write-Error -Message "Queue '$Name' could not be found at the specified path: '$QueuePath'."
                return
            }

            $Queue = New-Object -TypeName System.Messaging.MessageQueue -ArgumentList $QueuePath

            $OutputObject = [PSCustomObject]@{
                    Name          = $Name
                    Path          = $Queue.Path
                    Transactional = $Queue.Transactional
                    Authenticate  = $Queue.Authenticate
                    Journaling    = $Queue.UseJournalQueue
                    JournalQuota  = [UInt32]$Queue.MaximumJournalSize
                    Label         = $Queue.Label
                    PrivacyLevel  = [String]$Queue.EncryptionRequired
                    QueueQuota    = [UInt32]$Queue.MaximumQueueSize
                }

            return $OutputObject
        } -Session $session -ArgumentList $Name,$Cluster
    }
    finally
    {
        Remove-cMsmqSession $session | Out-Null
    }
}

function New-cMsmqQueue
{
    <#
    .SYNOPSIS
        Creates a new private MSMQ queue.
    .DESCRIPTION
        The New-cMsmqQueue function creates a new private MSMQ queue.
    .PARAMETER Name
        Specifies the name of the queue.
    .PARAMETER Transactional
        Specifies whether the queue is a transactional queue.
    .PARAMETER Authenticate
        Sets a value that indicates whether the queue accepts only authenticated messages.
    .PARAMETER Journaling
        Sets a value that indicates whether received messages are copied to the journal queue.
    .PARAMETER JournalQuota
        Sets the maximum size of the journal queue in KB.
    .PARAMETER Label
        Sets the queue description.
    .PARAMETER PrivacyLevel
        Sets the privacy level associated with the queue.
    .PARAMETER QueueQuota
        Sets the maximum size of the queue in KB.
    .PARAMETER Cluster
        Specifies the name of the cluster
    #>
    [CmdletBinding(ConfirmImpact = 'Medium', SupportsShouldProcess = $true)]
    param
    (
        [Parameter( Mandatory = $true, ValueFromPipeline = $true)]
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

    if (-not $PSCmdlet.ShouldProcess($Name, 'Create Queue'))
    {
        return
    }
    
    $session = New-cMsmqSession $Cluster

    try
    {
        Invoke-Command -ScriptBlock {
            param
            (
                $Name,
                $Transactional,
                $Parameters
            )
            
            $PropertyNames = @{
                Authenticate = 'Authenticate'
                Journaling   = 'UseJournalQueue'
                JournalQuota = 'MaximumJournalSize'
                Label        = 'Label'
                PrivacyLevel = 'EncryptionRequired'
                QueueQuota   = 'MaximumQueueSize'
            }

            $QueuePath = '.\private$\{0}' -f $Name

            try
            {
                $Queue = [System.Messaging.MessageQueue]::Create($QueuePath, $Transactional)
            }
            catch
            {
                Write-Error -Message $_.Exception.Message
                return
            }

            $Parameters.GetEnumerator() |
            Where-Object {$_.Key -in $PropertyNames.Keys} |
            ForEach-Object {

                $PropertyName = $PropertyNames.Item($_.Key)

                if ($Queue."$PropertyName" -cne $_.Value)
                {
                    "Setting property '{0}' to value '{1}'." -f $PropertyName, $_.Value |
                    Write-Verbose

                    $Queue."$PropertyName" = $_.Value
                }
            }
        } -ArgumentList $Name,$Transactional,$PSBoundParameters -Session $session
    }
    finally
    {
        Remove-cMsmqSession $session | Out-Null
    }
}

function Remove-cMsmqQueue
{
    <#
    .SYNOPSIS
        Removes the specified private MSMQ queue.
    .DESCRIPTION
        The Remove-cMsmqQueue function the specified private MSMQ queue.
    .PARAMETER Name
        Specifies the name of the queue.
    .PARAMETER Cluster
        Specifies the name of the cluster
    #>
    [CmdletBinding(ConfirmImpact = 'High', SupportsShouldProcess = $true)]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster
    )

    if (-not $PSCmdlet.ShouldProcess($Name, 'Remove Queue'))
    {
        return
    }

    $session = New-cMsmqSession $Cluster

    try
    {
        Invoke-Command -ScriptBlock {
            param
            (
                $Name
            )
            
            $QueuePath = '.\private$\{0}' -f $Name

            try
            {
                [Void][System.Messaging.MessageQueue]::Delete($QueuePath)
            }
            catch
            {
                Write-Error -Message $_.Exception.Message
                return
            }
        } -ArgumentList $Name -Session $session
    }
    finally
    {
        Remove-cMsmqSession $session | Out-Null
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
    .PARAMETER Cluster
        Specifies the name of the cluster
    #>
    [CmdletBinding(ConfirmImpact = 'High', SupportsShouldProcess = $true)]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster
    )
    
    if (-not $PSCmdlet.ShouldProcess($Name, 'Reset Queue Security'))
    {
        return
    }

    $session = New-cMsmqSession $Cluster

    try
    {
        Invoke-Command -ScriptBlock {
            param
            (
                $Name
            )

            $DefaultSecurity = 'Security=010007801c0000002800000000000000140000000200080000000000' +
                               '010100000000000512000000010500000000000515000000e611610036157811027bc60001020000'
            
            $QueuePath = '.\private$\{0}' -f $Name
            $QueueOwner = [cMsmq.Security]::GetOwner($Name)
            $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            
            Write-Verbose -Message "Queue '$Name' Owner is '$QueueOwner' while current user is '$CurrentUser'"

            if ($CurrentUser -ne $QueueOwner)
            {
                Write-Verbose -Message "Taking ownership of queue '$Name'."

                $FilePath = Get-ChildItem -Path "$Env:SystemRoot\System32\msmq\storage\lqs" -Force |
                    Select-String -Pattern "QueueName=\private`$\$($Name)" -SimpleMatch |
                    Select-Object -ExpandProperty Path

                if (-not $FilePath)
                {
                    Write-Error -Message "Could not find a corresponding .INI file for queue '$Name'."
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
        } -ArgumentList $Name -Session $session
    }
    finally
    {
        Remove-cMsmqSession $session | Out-Null
    }
}

function Set-cMsmqQueue
{
    <#
    .SYNOPSIS
        Sets properties on the specified private MSMQ queue.
    .DESCRIPTION
        The Set-cMsmqQueue function sets properties on the specified private MSMQ queue.
    .PARAMETER Name
        Specifies the name of the queue.
    .PARAMETER Authenticate
        Sets a value that indicates whether the queue accepts only authenticated messages.
    .PARAMETER Journaling
        Sets a value that indicates whether received messages are copied to the journal queue.
    .PARAMETER JournalQuota
        Sets the maximum size of the journal queue in KB.
    .PARAMETER Label
        Sets the queue description.
    .PARAMETER PrivacyLevel
        Sets the privacy level associated with the queue.
    .PARAMETER QueueQuota
        Sets the maximum size of the queue in KB.
    .PARAMETER Cluster
        Specifies the name of the cluster
    #>
    [CmdletBinding(ConfirmImpact = 'Medium', SupportsShouldProcess = $true)]
    param
    (
        [Parameter( Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,

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

    if (-not $PSCmdlet.ShouldProcess($Name, 'Set Queue'))
    {
        return
    }

    $session = New-cMsmqSession $Cluster

    try
    {
        Invoke-Command -ScriptBlock {
            param
            (
                $Name,
                $Parameters
            )

            $PropertyNames = @{
                Authenticate = 'Authenticate'
                Journaling   = 'UseJournalQueue'
                JournalQuota = 'MaximumJournalSize'
                Label        = 'Label'
                PrivacyLevel = 'EncryptionRequired'
                QueueQuota   = 'MaximumQueueSize'
            }

            $QueuePath = '.\private$\{0}' -f $Name

            if (-not [System.Messaging.MessageQueue]::Exists($QueuePath))
            {
                Write-Error -Message "Queue '$Name' could not be found at the specified path: '$QueuePath'."
                return
            }

            $Queue = New-Object -TypeName System.Messaging.MessageQueue -ArgumentList $QueuePath

            $Parameters.GetEnumerator() |
            Where-Object {$_.Key -in $PropertyNames.Keys} |
            ForEach-Object {

                $PropertyName = $PropertyNames.Item($_.Key)

                if ($Queue."$PropertyName" -ne $_.Value)
                {
                    "Setting property '{0}' to value '{1}'." -f $PropertyName, $_.Value |
                    Write-Verbose

                    $Queue."$PropertyName" = $_.Value
                }
            }
        } -ArgumentList $Name,$PSBoundParameters -Session $session
    }
    finally
    {
        Remove-cMsmqSession $session | Out-Null
    }
}

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
    .PARAMETER Cluster
        Specifies the name of the cluster
    #>
    [CmdletBinding()]
    [OutputType([System.Messaging.MessageQueueAccessRights])]
    param
    (
        [Parameter(Mandatory = $true)]
        [String]
        $Name,

        [Parameter(Mandatory = $false)]
        [String]
        $Principal = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name),
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster
    )

    $session = New-cMsmqSession $Cluster

    try
    {
        Invoke-Command -ScriptBlock {
            [CmdletBinding()]
            param(
                $Name,
                $Principal
            )
            
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
        } -Session $session -ArgumentList $Name,$Principal
    }
    finally
    {
        Remove-cMsmqSession $session | Out-Null
    }
}

function Test-cMsmqPermissions
{
    <#
    .SYNOPSIS
        Tests the specified permission on the specified private MSMQ queue for the specified user
    .DESCRIPTION
        The Set-cMsmqQueue function sets properties on the specified private MSMQ queue.
    .PARAMETER Name
        Specifies the name of the queue.
    .PARAMETER Principal
        Specifies the identity of the principal.
    .PARAMETER Permission
        Specifies the Permission to test
    .PARAMETER Cluster
        Specifies the name of the cluster
    #>
    [CmdletBinding()]
    param
    (
        [Parameter( Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,
        
        [Parameter(Mandatory = $true)]
        [System.Messaging.MessageQueueAccessRights]
        $Permission,
        
        [Parameter(Mandatory = $false)]
        [String]
        $Principal = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name),
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster
    )
    
    Write-Verbose -Message "Testing if the user '$Principal' has the permission necessary to perform the operation ($Permission)"
    $CurrentUserPermission = Get-cMsmqQueuePermission -Name $Name -Principal $Principal -Cluster $Cluster -ErrorAction SilentlyContinue

    if (-not $CurrentUserPermission -or -not $CurrentUserPermission.HasFlag($Permission))
    {
        Write-Verbose "User '$Principal' does not have the '$Permission' permission on queue '$Name'."
        return $false
    }
    
    return $true
}

function Set-cMsmqPermissions
{
    <#
    .SYNOPSIS
        Set ore Revoke the specified permission on the specified private MSMQ queue for the specified user
    .DESCRIPTION
        The Set-cMsmqPermissions function sets permissions on the specified private MSMQ queue for the specified user.
    .PARAMETER Name
        Specifies the name of the queue.
    .PARAMETER Permission
        Specifies the Permission to assign or revoke
    .PARAMETER Principal
        Specifies the identity of the principal.
    .PARAMETER Cluster
        Specifies the name of the cluster
    #>
    [CmdletBinding()]
    param
    (
        [Parameter( Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,
        
        [Parameter(Mandatory = $true)]
        [System.Messaging.MessageQueueAccessRights]
        $Permission,
        
        [Parameter(Mandatory = $false)]
        [String]
        $Principal = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name),
        
        [Parameter(Mandatory = $false)]
        [String]
        $Cluster,
        
        [switch]
        $Revoke
    )
    
    $session = New-cMsmqSession $Cluster

    try
    {
        Invoke-Command -ScriptBlock {
            param
            (
                $Name,
                $Permission,
                $Principal,
                $Revoke
            )

            $QueuePath = '.\private$\{0}' -f $Name

            if (-not [System.Messaging.MessageQueue]::Exists($QueuePath))
            {
                Write-Error -Message "Queue '$Name' could not be found at the specified path: '$QueuePath'."
                return
            }

            $Queue = New-Object -TypeName System.Messaging.MessageQueue -ArgumentList $QueuePath

            if ($Revoke)
            {
                Write-Verbose -Message "Revoking permissions '$Permission' for principal '$Principal' on queue '$Name'."
                $Queue.SetPermissions($Principal, $Permission, [System.Messaging.AccessControlEntryType]::Revoke)
            }
            else
            {
                Write-Verbose -Message "Setting permission '$Permission' for principal '$Principal' on queue '$Name'."
                $Queue.SetPermissions($Principal, $Permission, [System.Messaging.AccessControlEntryType]::Set)
            }
        } -ArgumentList $Name,$Permission,$Principal,$Revoke -Session $session
    }
    finally
    {
        Remove-cMsmqSession $session | Out-Null
    }
}
