$defaultValues = @{
ErrorAction = 'SilentlyContinue'
WarningAction = 'SilentlyContinue'
Force = $true
}

$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\Common\ClientTelemetry' @defaultValues
$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Common' @defaultValues
$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Common\Feedback' @defaultValues
$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail' @defaultValues
$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Options' @defaultValues
$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Common\General' @defaultValues
$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Common\ptwatson' @defaultValues
$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Firstrun' @defaultValues
$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\OSM' @defaultValues
$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\OSM\preventedapplications' @defaultValues
$null = New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\Common' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\Common\General' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\Common\ptwatson' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\Firstrun' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\OSM' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\Excel\Security' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\Outlook\Security' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\PowerPoint\Security' @defaultValues
$null = New-Item -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\Word\Security' @defaultValues

$defaultValues = @{
PropertyType = 'DWord'
ErrorAction = 'SilentlyContinue'
WarningAction = 'SilentlyContinue'
Force = $true
}

$WorkPath = 'HKCU:\Software\Microsoft\Office\Common\ClientTelemetry'
$null = New-ItemProperty -Path $WorkPath -Name DisableTelemetry -Value 1 @defaultValues

$WorkPath = 'HKCU:\Software\Microsoft\Office\16.0\Common'
$null = New-ItemProperty -Path $WorkPath -Name sendcustomerdata -Value 0 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name qmenable -Value 0 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name updatereliabilitydata -Value 0 @defaultValues

$WorkPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\Feedback'
$null = New-ItemProperty -Path $WorkPath -Name enabled -Value 0 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name includescreenshot -Value 0 @defaultValues

$WorkPath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail'
$null = New-ItemProperty -Path $WorkPath -Name EnableLogging -Value 0 @defaultValues

$WorkPath = 'HKCU:\Software\Microsoft\Office\16.0\Word\Options'
$null = New-ItemProperty -Path $WorkPath -Name EnableLogging -Value 0 @defaultValues

$WorkPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\General'
$null = New-ItemProperty -Path $WorkPath -Name shownfirstrunoptin -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name skydrivesigninoption -Value 0 @defaultValues

$WorkPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\ptwatson'
$null = New-ItemProperty -Path $WorkPath -Name ptwoptin -Value 0 @defaultValues

$WorkPath = 'HKCU:\Software\Microsoft\Office\16.0\Firstrun'
$null = New-ItemProperty -Path $WorkPath -Name disablemovie -Value 1 @defaultValues

$WorkPath = 'HKCU:\Software\Microsoft\Office\16.0\OSM'
$null = New-ItemProperty -Path $WorkPath -Name EnableLogging -Value 0 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name EnableUpload -Value 0 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name EnableFileObfuscation -Value 1 @defaultValues

$WorkPath = 'HKCU:\Software\Microsoft\Office\16.0\OSM\preventedapplications'
$null = New-ItemProperty -Path $WorkPath -Name accesssolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name olksolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name onenotesolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name pptsolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name projectsolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name publishersolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name visiosolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name wdsolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name xlsolution -Value 1 @defaultValues

$WorkPath = 'HKCU:\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes'
$null = New-ItemProperty -Path $WorkPath -Name agave -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name appaddins -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name comaddins -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name documentfiles -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name templatefiles -Value 1 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\Common'
$null = New-ItemProperty -Path $WorkPath -Name qmenable -Value 0 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name updatereliabilitydata -Value 0 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\Common\General'
$null = New-ItemProperty -Path $WorkPath -Name shownfirstrunoptin -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name skydrivesigninoption -Value 0 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\Common\ptwatson'
$null = New-ItemProperty -Path $WorkPath -Name ptwoptin -Value 0 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\Firstrun'
$null = New-ItemProperty -Path $WorkPath -Name disablemovie -Value 1 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\OSM'
$null = New-ItemProperty -Path $WorkPath -Name EnableLogging -Value 0 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name EnableUpload -Value 0 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name EnableFileObfuscation -Value 1 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications'
$null = New-ItemProperty -Path $WorkPath -Name accesssolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name olksolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name onenotesolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name pptsolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name projectsolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name publishersolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name visiosolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name wdsolution -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name xlsolution -Value 1 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes'
$null = New-ItemProperty -Path $WorkPath -Name agave -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name appaddins -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name comaddins -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name documentfiles -Value 1 @defaultValues
$null = New-ItemProperty -Path $WorkPath -Name templatefiles -Value 1 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\Excel\Security'
$null = New-ItemProperty -Path $WorkPath -Name blockcontentexecutionfrominternet -Value 1 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\Outlook\Security'
$null = New-ItemProperty -Path $WorkPath -Name level -Value 2 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\PowerPoint\Security'
$null = New-ItemProperty -Path $WorkPath -Name blockcontentexecutionfrominternet -Value 1 @defaultValues

$WorkPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\Word\Security'
$null = New-ItemProperty -Path $WorkPath -Name blockcontentexecutionfrominternet -Value 1 @defaultValues