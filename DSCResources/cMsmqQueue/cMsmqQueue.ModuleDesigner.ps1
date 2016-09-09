#requires -Version 4.0 -Modules xDSCResourceDesigner

$DscResourceName = 'cMsmqQueue'
$DscFriendlyName = 'cMsmqQueue'
$DscClassVersion = '1.0.0'

$scriptRoot = Split-Path $MyInvocation.MyCommand.Path
$moduleRoot = Resolve-Path (join-Path $scriptroot "..\..")
$schemaPath = (join-path $scriptRoot "$resourceName.schema.mof")

$DscResourceProperties =  @(
    (New-xDscResourceProperty -Type String -Attribute Write -Name Ensure -ValidateSet 'Absent', 'Present' -Description 'Indicates whether the queue exists.')
    (New-xDscResourceProperty -Type String -Attribute Key -Name Name -Description 'Indicates the name of the queue.')
    (New-xDscResourceProperty -Type Boolean -Attribute Write -Name Transactional -Description 'Indicates whether the queue is transactional.')
    (New-xDscResourceProperty -Type Boolean -Attribute Write -Name Authenticate -Description 'Indicates whether the queue accepts only authenticated messages.')
    (New-xDscResourceProperty -Type Boolean -Attribute Write -Name Journaling -Description 'Indicates whether received messages are copied to the journal queue.')
    (New-xDscResourceProperty -Type UInt32 -Attribute Write -Name JournalQuota -Description 'Indicates the maximum size of the journal queue in KB.')
    (New-xDscResourceProperty -Type String -Attribute Write -Name Label -Description 'Indicates the description of the queue.')
    (New-xDscResourceProperty -Type String -Attribute Write -Name PrivacyLevel -ValidateSet 'None', 'Optional', 'Body' -Description 'Indicates the privacy level associated with the queue.')
    (New-xDscResourceProperty -Type UInt32 -Attribute Write -Name QueueQuota -Description 'Indicates the maximum size of the queue in KB.'),
    (New-xDscResourceProperty -Type String -Attribute Write -Name Cluster -Description 'The name of the failover cluster where to create the queue.')
)

Write-Host "updating '$moduleRoot' ..."

New-xDscResource -Name $DscResourceName `
                 -Property $DscResourceProperties `
                 -Path $moduleRoot `
                 -FriendlyName $DscFriendlyName `
                 -ClassVersion $DscClassVersion `
                 -Verbose
