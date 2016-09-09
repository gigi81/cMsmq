#requires -Version 4.0 -Modules xDSCResourceDesigner

$DscModuleName   = 'cMsmq'
$DscResourceName = 'cMsmqQueuePermissionEntry'
$DscFriendlyName = 'cMsmqQueuePermissionEntry'
$DscClassVersion = '1.0.0'

$scriptRoot = Split-Path $MyInvocation.MyCommand.Path
$moduleRoot = Resolve-Path (join-Path $scriptroot "..\..")
$schemaPath = (join-path $scriptRoot "$resourceName.schema.mof")

$DscResourceProperties =  @(
    (New-xDscResourceProperty -Type String -Attribute Write -Name Ensure -ValidateSet 'Absent', 'Present' -Description 'Indicates whether the permission entry exists.')
    (New-xDscResourceProperty -Type String -Attribute Key -Name Name -Description 'Indicates the name of the queue.'),
    (New-xDscResourceProperty -Type String -Attribute Key -Name Principal -Description 'Indicates the identity of the principal.'),
    (New-xDscResourceProperty -Type String[] -Attribute Write -Name AccessRights -Description 'Indicates the access rights to be granted to the principal.'),
    (New-xDscResourceProperty -Type String -Attribute Write -Name Cluster -Description 'The name of the failover cluster where to create the queue is hosted')
)

Write-Host "updating '$moduleRoot' ..."

New-xDscResource -Name $DscResourceName `
                 -ModuleName $DscModuleName `
                 -Property $DscResourceProperties `
                 -Path $moduleRoot `
                 -FriendlyName $DscFriendlyName `
                 -ClassVersion $DscClassVersion `
                 -Verbose
