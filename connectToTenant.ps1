try{
    Get-InstalledModule -Name ExchangeOnlineManagement -ea stop
    Get-InstalledModule -Name AzureAD -ea stop
    Get-InstalledModule -Name MSOnline -ea stop
}
catch{
    Start-Process powershell -Verb runAs -ArgumentList "Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.1.0"
    Start-Process powershell -Verb runAs -ArgumentList "Install-Module -Name AzureAD"
    Start-Process powershell -Verb runAs -ArgumentList "Install-Module -Name MSOnline"
}



Write-Output "Dependencies installed"

Import-Module ExchangeOnlineManagement

Import-Module  AzureAD

Import-Module MSOnline

Write-Output "Specify domain of tenant you wish to connect to:"
$tenantDomain = Read-Host

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select an Admin Center'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Select which admin center to connect to'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

[void] $listBox.Items.Add('Exchange Online')
[void] $listBox.Items.Add('Teams')
[void] $listBox.Items.Add('Sharepoint')
[void] $listBox.Items.Add('Azure AD')
[void] $listBox.Items.Add('Security & Compliance')

$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $x = $listBox.SelectedItem
    
}

if (($x -clike 'Teams') -or ($x -clike 'Azure AD')){
    Connect-MsolService

    $tenantID = Get-MsolPartnerContract -DomainName $tenantDomain | Select-Object TenantID
    $tenantID.GetType()
    $finalTenantID = [System.Management.Automation.LanguagePrimitives]::ConvertTo($tenantID, [string]).substring(11) -replace ".$"
    $finalTenantID
}

switch ($x)
{
    'Exchange Online'{Start-Process powershell -ArgumentList "-noexit Connect-ExchangeOnline -DelegatedOrganization $tenantDomain"}
    'Teams' {Start-Process powershell -ArgumentList "-noexit Connect-MicrosoftTeams -TenantId $finalTenantID"}
    'Sharepoint'{Write-Output "Not supported yet, you need to connect with a global admin."}
    'Azure AD'{Connect-AzureAD -TenantId $finalTenantID}
    'Security & Compliance'{Connect-IPPSSession -DelegatedOrganization $tenantDomain}
}
