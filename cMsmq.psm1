$dllFilePath = Join-Path -Path $PSScriptRoot -ChildPath 'cMsmq.dll'

if ([AppDomain]::CurrentDomain.GetAssemblies().Location -notcontains $dllFilePath)
{
	Add-Type -Path $dllFilePath -ErrorAction Stop
}

if ([AppDomain]::CurrentDomain.GetAssemblies().ManifestModule.Name -notcontains 'System.Messaging.dll')
{
	Add-Type -AssemblyName System.Messaging -ErrorAction Stop
}

$functionRoot = Join-Path -Path $PSScriptRoot -ChildPath 'Functions'

Get-ChildItem -Path $functionRoot -Filter '*.ps1' | 
	ForEach-Object {
		Write-Verbose ("Importing function {0}." -f $_.FullName)
		. $_.FullName | Out-Null
	}

