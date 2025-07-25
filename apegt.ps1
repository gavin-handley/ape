[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
$groupTag = Read-Host "Enter Group Tag for Autopilot enrollment"
Install-Script -Name Get-WindowsAutopilotInfo -Force -SkipPublisherCheck
Get-WindowsAutopilotInfo -Online -GroupTag $groupTag
