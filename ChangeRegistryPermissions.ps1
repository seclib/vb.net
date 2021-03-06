$aUsers = @("builtin\Power Users","builtin\Users","builtin\Administrators","Creator Owner")
#$aKeys = @("HKLM:\Software\Wow6432Node\Cad Zone","HKLM:\Software\Wow6432Node\Cad Zone\Crash Zone"`
#,"HKLM:\Software\Wow6432Node\Cad Zone\Forms")
Function fRegPerm($sUser){
	$pc = $env:COMPUTERNAME
	$acl = Get-Acl "HKLM:\Software\Wow6432Node\Cad Zone"
	$rule = New-Object System.Security.AccessControl.RegistryAccessRule ("$users","FullControl","ContainerInherit,ObjectInherit","None","Allow")
	$acl.SetAccessRule($rule)
	$acl | Set-Acl "HKLM:\Software\Wow6432Node\Cad Zone"
}
$aUsers | foreach{fRegPerm $_}
