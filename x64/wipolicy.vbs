' Windows Installer utility to manage installer policy settings
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of the installer policy keys
' Policy can be configured by an administrator using the NT Group Policy Editor
'
Option Explicit

Dim policies(21, 4)
policies(1, 0)="LM" : policies(1, 1)="HKLM" : policies(1, 2)="Logging"              : policies(1, 3)="REG_SZ"    : policies(1, 4) = "Logging modes if not supplied by install, set of iwearucmpv"
policies(2, 0)="DO" : policies(2, 1)="HKLM" : policies(2, 2)="Debug"                : policies(2, 3)="REG_DWORD" : policies(2, 4) = "OutputDebugString: 1=debug output, 2=verbose debug output, 7=include command line"
policies(3, 0)="DI" : policies(3, 1)="HKLM" : policies(3, 2)="DisableMsi"           : policies(3, 3)="REG_DWORD" : policies(3, 4) = "1=Disable non-managed installs, 2=disable all installs"
policies(4, 0)="WT" : policies(4, 1)="HKLM" : policies(4, 2)="Timeout"              : policies(4, 3)="REG_DWORD" : policies(4, 4) = "Wait timeout in seconds in case of no activity"
policies(5, 0)="DB" : policies(5, 1)="HKLM" : policies(5, 2)="DisableBrowse"        : policies(5, 3)="REG_DWORD" : policies(5, 4) = "Disable user browsing of source locations if 1"
policies(6, 0)="AB" : policies(6, 1)="HKLM" : policies(6, 2)="AllowLockdownBrowse"  : policies(6, 3)="REG_DWORD" : policies(6, 4) = "Allow non-admin users to browse to new sources for managed applications if 1 - use is not recommended"
policies(7, 0)="AM" : policies(7, 1)="HKLM" : policies(7, 2)="AllowLockdownMedia"   : policies(7, 3)="REG_DWORD" : policies(7, 4) = "Allow non-admin users to browse to new media sources for managed applications if 1 - use is not recommended"
policies(8, 0)="AP" : policies(8, 1)="HKLM" : policies(8, 2)="AllowLockdownPatch"   : policies(8, 3)="REG_DWORD" : policies(8, 4) = "Allow non-admin users to apply small and minor update patches to managed applications if 1 - use is not recommended"
policies(9, 0)="DU" : policies(9, 1)="HKLM" : policies(9, 2)="DisableUserInstalls"  : policies(9, 3)="REG_DWORD" : policies(9, 4) = "Disable per-user installs if 1 - available on Windows Installer version 2.0 and later"
policies(10, 0)="DP" : policies(10, 1)="HKLM" : policies(10, 2)="DisablePatch"         : policies(10, 3)="REG_DWORD" : policies(10, 4) = "Disable patch application to all products if 1"
policies(11, 0)="UC" : policies(11, 1)="HKLM" : policies(11, 2)="EnableUserControl"    : policies(11, 3)="REG_DWORD" : policies(11, 4) = "All public properties sent to install service if 1"
policies(12, 0)="ER" : policies(12, 1)="HKLM" : policies(12, 2)="EnableAdminTSRemote"  : policies(12, 3)="REG_DWORD" : policies(12, 4) = "Allow admins to perform installs from terminal server client sessions if 1"
policies(13, 0)="LS" : policies(13, 1)="HKLM" : policies(13, 2)="LimitSystemRestoreCheckpointing" : policies(13, 3)="REG_DWORD" : policies(13, 4) = "Turn off creation of system restore check points on Windows XP if 1 - available on Windows Installer version 2.0 and later"
policies(14, 0)="SS" : policies(14, 1)="HKLM" : policies(14, 2)="SafeForScripting"     : policies(14, 3)="REG_DWORD" : policies(14, 4) = "Do not prompt when scripts within a webpage access Installer automation interface if 1 - use is not recommended"
policies(15, 0)="TP" : policies(15,1)="HKLM" : policies(15, 2)="TransformsSecure"     : policies(15, 3)="REG_DWORD" : policies(15, 4) = "Pin tranforms in secure location if 1 (only admin and system have write privileges to cache location)"
policies(16, 0)="EM" : policies(16, 1)="HKLM" : policies(16, 2)="AlwaysInstallElevated": policies(16, 3)="REG_DWORD" : policies(16, 4) = "System privileges if 1 and HKCU value also set - dangerous policy as non-admin users can install with elevated privileges if enabled"
policies(17, 0)="EU" : policies(17, 1)="HKCU" : policies(17, 2)="AlwaysInstallElevated": policies(17, 3)="REG_DWORD" : policies(17, 4) = "System privileges if 1 and HKLM value also set - dangerous policy as non-admin users can install with elevated privileges if enabled"
policies(18,0)="DR" : policies(18,1)="HKCU" : policies(18,2)="DisableRollback"      : policies(18,3)="REG_DWORD" : policies(18,4) = "Disable rollback if 1 - use is not recommended"
policies(19,0)="TS" : policies(19,1)="HKCU" : policies(19,2)="TransformsAtSource"   : policies(19,3)="REG_DWORD" : policies(19,4) = "Locate transforms at root of source image if 1"
policies(20,0)="SO" : policies(20,1)="HKCU" : policies(20,2)="SearchOrder"          : policies(20,3)="REG_SZ"    : policies(20,4) = "Search order of source types, set of n,m,u (default=nmu)"
policies(21,0)="DM" : policies(21,1)="HKCU" : policies(21,2)="DisableMedia"          : policies(21,3)="REG_DWORD"    : policies(21,4) = "Browsing to media sources is disabled"

Dim argCount:argCount = Wscript.Arguments.Count
Dim message, iPolicy, policyKey, policyValue, WshShell, policyCode
On Error Resume Next

' If no arguments supplied, then list all current policy settings
If argCount = 0 Then
	Set WshShell = WScript.CreateObject("WScript.Shell") : CheckError
	For iPolicy = 0 To UBound(policies)
		policyValue = ReadPolicyValue(iPolicy)
		If Not IsEmpty(policyValue) Then 'policy key present, else skip display
			If Not IsEmpty(message) Then message = message & vbLf
			message = message & policies(iPolicy,0) & ": " & policies(iPolicy,2) & "(" & policies(iPolicy,1) & ") = " & policyValue
		End If
	Next
	If IsEmpty(message) Then message = "No installer policies set"
	Wscript.Echo message
	Wscript.Quit 0
End If

' Check for ?, and show help message if found
policyCode = UCase(Wscript.Arguments(0))
If InStr(1, policyCode, "?", vbTextCompare) <> 0 Then
	message = "Windows Installer utility to manage installer policy settings" &_
		vbLf & " If no arguments are supplied, current policy settings in list will be reported" &_
		vbLf & " The 1st argument specifies the policy to set, using a code from the list below" &_
		vbLf & " The 2nd argument specifies the new policy setting, use """" to remove the policy" &_
		vbLf & " If the 2nd argument is not supplied, the current policy value will be reported"
	For iPolicy = 0 To UBound(policies)
		message = message & vbLf & policies(iPolicy,0) & ": " & policies(iPolicy,2) & "(" & policies(iPolicy,1) & ")  " & policies(iPolicy,4) & vbLf
	Next
	message = message & vblf & vblf & "Copyright (C) Microsoft Corporation.  All rights reserved."

	Wscript.Echo message
	Wscript.Quit 1
End If

' Policy code supplied, look up in array
For iPolicy = 0 To UBound(policies)
	If policies(iPolicy, 0) = policyCode Then Exit For
Next
If iPolicy > UBound(policies) Then Wscript.Echo "Unknown policy code: " & policyCode : Wscript.Quit 2
Set WshShell = WScript.CreateObject("WScript.Shell") : CheckError

' If no value supplied, then simply report current value
policyValue = ReadPolicyValue(iPolicy)
If IsEmpty(policyValue) Then policyValue = "Not present"
message = policies(iPolicy,0) & ": " & policies(iPolicy,2) & "(" & policies(iPolicy,1) & ") = " & policyValue
If argCount > 1 Then ' Value supplied, set policy
	policyValue = WritePolicyValue(iPolicy, Wscript.Arguments(1))
	If IsEmpty(policyValue) Then policyValue = "Not present"
	message = message & " --> " & policyValue
End If
Wscript.Echo message

Function ReadPolicyValue(iPolicy)
	On Error Resume Next
	Dim policyKey:policyKey = policies(iPolicy,1) & "\Software\Policies\Microsoft\Windows\Installer\" & policies(iPolicy,2)
	ReadPolicyValue = WshShell.RegRead(policyKey)
End Function

Function WritePolicyValue(iPolicy, policyValue)
	On Error Resume Next
	Dim policyKey:policyKey = policies(iPolicy,1) & "\Software\Policies\Microsoft\Windows\Installer\" & policies(iPolicy,2)
	If Len(policyValue) Then
		WshShell.RegWrite policyKey, policyValue, policies(iPolicy,3) : CheckError
		WritePolicyValue = policyValue
	ElseIf Not IsEmpty(ReadPolicyValue(iPolicy)) Then
		WshShell.RegDelete policyKey : CheckError
	End If
End Function

Sub CheckError
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbLf & errRec.FormatText
	End If
	Wscript.Echo message
	Wscript.Quit 2
End Sub

'' SIG '' Begin signature block
'' SIG '' MIImBQYJKoZIhvcNAQcCoIIl9jCCJfICAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' /P2vh2Oob6dK4LJgJX9q0NCUL4Tr2otbVtqd2lYQZG6g
'' SIG '' ggtnMIIE7zCCA9egAwIBAgITMwAABVfPkN3H0cCIjAAA
'' SIG '' AAAFVzANBgkqhkiG9w0BAQsFADB+MQswCQYDVQQGEwJV
'' SIG '' UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
'' SIG '' UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
'' SIG '' aWduaW5nIFBDQSAyMDEwMB4XDTIzMTAxOTE5NTExMloX
'' SIG '' DTI0MTAxNjE5NTExMlowdDELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEeMBwGA1UEAxMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
'' SIG '' MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
'' SIG '' rNP5BRqxQTyYzc7lY4sbAK2Huz47DGso8p9wEvDxx+0J
'' SIG '' gngiIdoh+jhkos8Hcvx0lOW32XMWZ9uWBMn3+pgUKZad
'' SIG '' OuLXO3LnuVop+5akCowquXhMS3uzPTLONhyePNp74iWb
'' SIG '' 1StajQ3uGOx+fEw00mrTpNGoDeRj/cUHOqKb/TTx2TCt
'' SIG '' 7z32yj/OcNp5pk+8A5Gg1S6DMZhJjZ39s2LVGrsq8fs8
'' SIG '' y1RP3ZBb2irsMamIOUFSTar8asexaAgoNauVnQMqeAdE
'' SIG '' tNScUxT6m/cNfOZjrCItHZO7ieiaDk9ljrCS9QVLldjI
'' SIG '' JhadWdjiAa8JHXgeecBvJhe2s9XVho5OTQIDAQABo4IB
'' SIG '' bjCCAWowHwYDVR0lBBgwFgYKKwYBBAGCNz0GAQYIKwYB
'' SIG '' BQUHAwMwHQYDVR0OBBYEFGVIsKghPtVDZfZAsyDVZjTC
'' SIG '' rXm3MEUGA1UdEQQ+MDykOjA4MR4wHAYDVQQLExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xFjAUBgNVBAUTDTIzMDg2
'' SIG '' NSs1MDE1OTcwHwYDVR0jBBgwFoAU5vxfe7siAFjkck61
'' SIG '' 9CF0IzLm76wwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
'' SIG '' L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVj
'' SIG '' dHMvTWljQ29kU2lnUENBXzIwMTAtMDctMDYuY3JsMFoG
'' SIG '' CCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
'' SIG '' L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWND
'' SIG '' b2RTaWdQQ0FfMjAxMC0wNy0wNi5jcnQwDAYDVR0TAQH/
'' SIG '' BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAyi7DQuZQIWdC
'' SIG '' y9r24eaW4WAzNYbRIN/nYv+fHw77U3E/qC8KvnkT7iJX
'' SIG '' lGit+3mhHspwiQO1r3SSvRY72QQuBW5KoS7upUqqZVFH
'' SIG '' ic8Z+ttKnH7pfqYXFLM0GA8gLIeH43U8ybcdoxnoiXA9
'' SIG '' fd8iKCM4za5ZVwrRlTEo68sto4lOKXM6dVjo1qwi/X89
'' SIG '' Gb0fNdWGQJ4cj+s7tVfKXWKngOuvISr3X2c1aetBfGZK
'' SIG '' p7nDqWtViokBGBMJBubzkHcaDsWVnPjCenJnDYAPu0ny
'' SIG '' W29F1/obCiMyu02/xPXRCxfPOe97LWPgLrgKb2SwLBu+
'' SIG '' mlP476pcq3lFl+TN7ltkoTCCBnAwggRYoAMCAQICCmEM
'' SIG '' UkwAAAAAAAMwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBS
'' SIG '' b290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEwMB4X
'' SIG '' DTEwMDcwNjIwNDAxN1oXDTI1MDcwNjIwNTAxN1owfjEL
'' SIG '' MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
'' SIG '' EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
'' SIG '' c29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
'' SIG '' b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMDCCASIwDQYJ
'' SIG '' KoZIhvcNAQEBBQADggEPADCCAQoCggEBAOkOZFB5Z7XE
'' SIG '' 4/0JAEyelKz3VmjqRNjPxVhPqaV2fG1FutM5krSkHvn5
'' SIG '' ZYLkF9KP/UScCOhlk84sVYS/fQjjLiuoQSsYt6JLbklM
'' SIG '' axUH3tHSwokecZTNtX9LtK8I2MyI1msXlDqTziY/7Ob+
'' SIG '' NJhX1R1dSfayKi7VhbtZP/iQtCuDdMorsztG4/BGScEX
'' SIG '' ZlTJHL0dxFViV3L4Z7klIDTeXaallV6rKIDN1bKe5QO1
'' SIG '' Y9OyFMjByIomCll/B+z/Du2AEjVMEqa+Ulv1ptrgiwtI
'' SIG '' d9aFR9UQucboqu6Lai0FXGDGtCpbnCMcX0XjGhQebzfL
'' SIG '' GTOAaolNo2pmY3iT1TDPlR8CAwEAAaOCAeMwggHfMBAG
'' SIG '' CSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBTm/F97uyIA
'' SIG '' WORyTrX0IXQjMubvrDAZBgkrBgEEAYI3FAIEDB4KAFMA
'' SIG '' dQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUw
'' SIG '' AwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvX
'' SIG '' zpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3Js
'' SIG '' Lm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9N
'' SIG '' aWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYB
'' SIG '' BQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3
'' SIG '' Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0Nl
'' SIG '' ckF1dF8yMDEwLTA2LTIzLmNydDCBnQYDVR0gBIGVMIGS
'' SIG '' MIGPBgkrBgEEAYI3LgMwgYEwPQYIKwYBBQUHAgEWMWh0
'' SIG '' dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9QS0kvZG9jcy9D
'' SIG '' UFMvZGVmYXVsdC5odG0wQAYIKwYBBQUHAgIwNB4yIB0A
'' SIG '' TABlAGcAYQBsAF8AUABvAGwAaQBjAHkAXwBTAHQAYQB0
'' SIG '' AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIB
'' SIG '' ABp071dPKXvEFoV4uFDTIvwJnayCl/g0/yosl5US5eS/
'' SIG '' z7+TyOM0qduBuNweAL7SNW+v5X95lXflAtTx69jNTh4b
'' SIG '' YaLCWiMa8IyoYlFFZwjjPzwek/gwhRfIOUCm1w6zISnl
'' SIG '' paFpjCKTzHSY56FHQ/JTrMAPMGl//tIlIG1vYdPfB9XZ
'' SIG '' cgAsaYZ2PVHbpjlIyTdhbQfdUxnLp9Zhwr/ig6sP4Gub
'' SIG '' ldZ9KFGwiUpRpJpsyLcfShoOaanX3MF+0Ulwqratu3JH
'' SIG '' Yxf6ptaipobsqBBEm2O2smmJBsdGhnoYP+jFHSHVe/kC
'' SIG '' Iy3FQcu/HUzIFu+xnH/8IktJim4V46Z/dlvRU3mRhZ3V
'' SIG '' 0ts9czXzPK5UslJHasCqE5XSjhHamWdeMoz7N4XR3HWF
'' SIG '' nIfGWleFwr/dDY+Mmy3rtO7PJ9O1Xmn6pBYEAackZ3PP
'' SIG '' TU+23gVWl3r36VJN9HcFT4XG2Avxju1CCdENduMjVngi
'' SIG '' Jja+yrGMbqod5IXaRzNij6TJkTNfcR5Ar5hlySLoQiEl
'' SIG '' ihwtYNk3iUGJKhYP12E8lGhgUu/WR5mggEDuFYF3Ppzg
'' SIG '' UxgaUB04lZseZjMTJzkXeIc2zk7DX7L1PUdTtuDl2wth
'' SIG '' PSrXkizON1o+QEIxpB8QCMJWnL8kXVECnWp50hfT2sGU
'' SIG '' jgd7JXFEqwZq5tTG3yOalnXFMYIZ9jCCGfICAQEwgZUw
'' SIG '' fjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWlj
'' SIG '' cm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMAITMwAA
'' SIG '' BVfPkN3H0cCIjAAAAAAFVzANBglghkgBZQMEAgEFAKCC
'' SIG '' AQQwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
'' SIG '' KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
'' SIG '' hvcNAQkEMSIEIBu+HHdVsB4YvaPdNhQbqqKDiMS+Soon
'' SIG '' Mp2YYFv3MDmrMDwGCisGAQQBgjcKAxwxLgwsc1BZN3hQ
'' SIG '' QjdoVDVnNUhIcll0OHJETFNNOVZ1WlJ1V1phZWYyZTIy
'' SIG '' UnM1ND0wWgYKKwYBBAGCNwIBDDFMMEqgJIAiAE0AaQBj
'' SIG '' AHIAbwBzAG8AZgB0ACAAVwBpAG4AZABvAHcAc6EigCBo
'' SIG '' dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vd2luZG93czAN
'' SIG '' BgkqhkiG9w0BAQEFAASCAQAaO3RAOrtaNCRrz5IlVr9+
'' SIG '' YBCjPCzDzUJjc+MUDccFj8vpJwi+kaZMQoe2ttvWVA7S
'' SIG '' uqklAoAjDe/LQRZq2gNNtV0jpAlaaRM+jJxtBVJDYApm
'' SIG '' tp7d+gD/OxUaJuhakRPpOzSLABESnkve6zDJ815PB+is
'' SIG '' hXooffaRpRbsyhwVmZ0MjPtw+uTmYx6uaUYYagMi/mNX
'' SIG '' WdPKa9yF+39U+GOS3Jc1LjTXPWH7hXfkRllhdZc1drEP
'' SIG '' Pm0Qu2NHKIx0K+fI1z62U6+EZfaXfN6uzD5MgR1lXfVC
'' SIG '' c+ZAshN8mzbqjwXwNWuPp3b9n+ivlnZx3nOHblJ+JpHm
'' SIG '' ALdmqm/y2/daoYIXKTCCFyUGCisGAQQBgjcDAwExghcV
'' SIG '' MIIXEQYJKoZIhvcNAQcCoIIXAjCCFv4CAQMxDzANBglg
'' SIG '' hkgBZQMEAgEFADCCAVkGCyqGSIb3DQEJEAEEoIIBSASC
'' SIG '' AUQwggFAAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUD
'' SIG '' BAIBBQAEICCtOWu0gwP6PLFQdfXIYowJjagu5aSh2Thi
'' SIG '' Qw0Id6f1AgZl1eP4+ZEYEzIwMjQwMjIyMTA0NDIxLjU3
'' SIG '' MVowBIACAfSggdikgdUwgdIxCzAJBgNVBAYTAlVTMRMw
'' SIG '' EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
'' SIG '' b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
'' SIG '' b24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9w
'' SIG '' ZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVz
'' SIG '' IFRTUyBFU046MDg0Mi00QkU2LUMyOUExJTAjBgNVBAMT
'' SIG '' HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WgghF4
'' SIG '' MIIHJzCCBQ+gAwIBAgITMwAAAdqO1claANERsQABAAAB
'' SIG '' 2jANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzET
'' SIG '' MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
'' SIG '' bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
'' SIG '' aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
'' SIG '' cCBQQ0EgMjAxMDAeFw0yMzEwMTIxOTA2NTlaFw0yNTAx
'' SIG '' MTAxOTA2NTlaMIHSMQswCQYDVQQGEwJVUzETMBEGA1UE
'' SIG '' CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEe
'' SIG '' MBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0w
'' SIG '' KwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRp
'' SIG '' b25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
'' SIG '' RVNOOjA4NDItNEJFNi1DMjlBMSUwIwYDVQQDExxNaWNy
'' SIG '' b3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkq
'' SIG '' hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAk5AGCHa1UVHW
'' SIG '' PyNADg0N/xtxWtdI3TzQI0o9JCjtLnuwKc9TQUoXjvDY
'' SIG '' vqoe3CbgScKUXZyu5cWn+Xs+kxCDbkTtfzEOa/GvwEET
'' SIG '' qIBIA8J+tN5u68CxlZwliHLumuAK4F/s6J1emCxbXLyn
'' SIG '' pWzuwPZq6n/S695jF5eUq2w+MwKmUeSTRtr4eAuGjQnr
'' SIG '' wp2OLcMzYrn3AfL3Gu2xgr5f16tsMZnaaZffvrlpLlDv
'' SIG '' +6APExWDPKPzTImfpQueScP2LiRRDFWGpXV1z8MXpQF6
'' SIG '' 7N+6SQx53u2vNQRkxHKVruqG/BR5CWDMJCGlmPP7OxCC
'' SIG '' leU9zO8Z3SKqvuUALB9UaiDmmUjN0TG+3VMDwmZ5/zX1
'' SIG '' pMrAfUhUQjBgsDq69LyRF0DpHG8xxv/+6U2Mi4Zx7LKQ
'' SIG '' wBcTKdWssb1W8rit+sKwYvePfQuaJ26D6jCtwKNBqBia
'' SIG '' saTWEHKReKWj1gHxDLLlDUqEa4frlXfMXLxrSTBsoFGz
'' SIG '' xVHge2g9jD3PUN1wl9kE7Z2HNffIAyKkIabpKa+a9q9G
'' SIG '' xeHLzTmOICkPI36zT9vuizbPyJFYYmToz265Pbj3eAVX
'' SIG '' /0ksaDlgkkIlcj7LGQ785edkmy4a3T7NYt0dLhchcEbX
'' SIG '' ug+7kqwV9FMdESWhHZ0jobBprEjIPJIdg628jJ2Vru7i
'' SIG '' V+d8KNj+opMCAwEAAaOCAUkwggFFMB0GA1UdDgQWBBSh
'' SIG '' fI3JUT1mE5WLMRRXCE2Avw9fRTAfBgNVHSMEGDAWgBSf
'' SIG '' pxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSg
'' SIG '' UqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
'' SIG '' b3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIw
'' SIG '' UENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBe
'' SIG '' MFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29m
'' SIG '' dC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRp
'' SIG '' bWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNV
'' SIG '' HRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
'' SIG '' MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOC
'' SIG '' AgEAuYNV1O24jSMAS3jU7Y4zwJTbftMYzKGsavsXMoIQ
'' SIG '' VpfG2iqT8g5tCuKrVxodWHa/K5DbifPdN04G/utyz+qc
'' SIG '' +M7GdcUvJk95pYuw24BFWZRWLJVheNdgHkPDNpZmBJxj
'' SIG '' wYovvIaPJauHvxYlSCHusTX7lUPmHT/quz10FGoDMj1+
'' SIG '' FnPuymyO3y+fHnRYTFsFJIfut9psd6d2l6ptOZb9F9xp
'' SIG '' P4YUixP6DZ6PvBEoir9CGeygXyakU08dXWr9Yr+sX8KG
'' SIG '' i+SEkwO+Wq0RNaL3saiU5IpqZkL1tiBw8p/Pbx53blYn
'' SIG '' LXRW1D0/n4L/Z058NrPVGZ45vbspt6CFrRJ89yuJN85F
'' SIG '' W+o8NJref03t2FNjv7j0jx6+hp32F1nwJ8g49+3C3fFN
'' SIG '' fZGExkkJWgWVpsdy99vzitoUzpzPkRiT7HVpUSJe2Arp
'' SIG '' HTGfXCMxcd/QBaVKOpGTO9KdErMWxnASXvhVqGUpWEj4
'' SIG '' KL1FP37oZzTFbMnvNAhQUTcmKLHn7sovwCsd8Fj1QUvP
'' SIG '' iydugntCKncgANuRThkvSJDyPwjGtrtpJh9OhR5+Zy3d
'' SIG '' 0zr19/gR6HYqH02wqKKmHnz0Cn/FLWMRKWt+Mv+D9luh
'' SIG '' pLl31rZ8Dn3ya5sO8sPnHk8/fvvTS+b9j48iGanZ9O+5
'' SIG '' Layd15kGbJOpxQ0dE2YKT6eNXecwggdxMIIFWaADAgEC
'' SIG '' AhMzAAAAFcXna54Cm0mZAAAAAAAVMA0GCSqGSIb3DQEB
'' SIG '' CwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
'' SIG '' aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
'' SIG '' ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQD
'' SIG '' EylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRo
'' SIG '' b3JpdHkgMjAxMDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5
'' SIG '' MzAxODMyMjVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
'' SIG '' EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
'' SIG '' HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
'' SIG '' BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAy
'' SIG '' MDEwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
'' SIG '' AgEA5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51
'' SIG '' yMo1V/YBf2xK4OK9uT4XYDP/XE/HZveVU3Fa4n5KWv64
'' SIG '' NmeFRiMMtY0Tz3cywBAY6GB9alKDRLemjkZrBxTzxXb1
'' SIG '' hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9cmmv
'' SIG '' Haus9ja+NSZk2pg7uhp7M62AW36MEBydUv626GIl3GoP
'' SIG '' z130/o5Tz9bshVZN7928jaTjkY+yOSxRnOlwaQ3KNi1w
'' SIG '' jjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDuaRr3tpK56
'' SIG '' KTesy+uDRedGbsoy1cCGMFxPLOJiss254o2I5JasAUq7
'' SIG '' vnGpF1tnYN74kpEeHT39IM9zfUGaRnXNxF803RKJ1v2l
'' SIG '' IH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2K26oElHovwUD
'' SIG '' o9Fzpk03dJQcNIIP8BDyt0cY7afomXw/TNuvXsLz1dhz
'' SIG '' PUNOwTM5TI4CvEJoLhDqhFFG4tG9ahhaYQFzymeiXtco
'' SIG '' dgLiMxhy16cg8ML6EgrXY28MyTZki1ugpoMhXV8wdJGU
'' SIG '' lNi5UPkLiWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsl
'' SIG '' uq9QBXpsxREdcu+N+VLEhReTwDwV2xo3xwgVGD94q0W2
'' SIG '' 9R6HXtqPnhZyacaue7e3PmriLq0CAwEAAaOCAd0wggHZ
'' SIG '' MBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUC
'' SIG '' BBYEFCqnUv5kxJq+gpE8RjUpzxD/LwTuMB0GA1UdDgQW
'' SIG '' BBSfpxVdAF5iXYP05dJlpxtTNRnpcjBcBgNVHSAEVTBT
'' SIG '' MFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNo
'' SIG '' dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0Rv
'' SIG '' Y3MvUmVwb3NpdG9yeS5odG0wEwYDVR0lBAwwCgYIKwYB
'' SIG '' BQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEw
'' SIG '' CwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYD
'' SIG '' VR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYD
'' SIG '' VR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3Nv
'' SIG '' ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2Vy
'' SIG '' QXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4w
'' SIG '' TDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3Nv
'' SIG '' ZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAx
'' SIG '' MC0wNi0yMy5jcnQwDQYJKoZIhvcNAQELBQADggIBAJ1V
'' SIG '' ffwqreEsH2cBMSRb4Z5yS/ypb+pcFLY+TkdkeLEGk5c9
'' SIG '' MTO1OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulm
'' SIG '' ZzpTTd2YurYeeNg2LpypglYAA7AFvonoaeC6Ce5732pv
'' SIG '' vinLbtg/SHUB2RjebYIM9W0jVOR4U3UkV7ndn/OOPcbz
'' SIG '' aN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECWOKz3
'' SIG '' +SmJw7wXsFSFQrP8DJ6LGYnn8AtqgcKBGUIZUnWKNsId
'' SIG '' w2FzLixre24/LAl4FOmRsqlb30mjdAy87JGA0j3mSj5m
'' SIG '' O0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3UwxTSwethQ
'' SIG '' /gpY3UA8x1RtnWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9
'' SIG '' swFXSVRk2XPXfx5bRAGOWhmRaw2fpCjcZxkoJLo4S5pu
'' SIG '' +yFUa2pFEUep8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyh
'' SIG '' YWxz/gq77EFmPWn9y8FBSX5+k77L+DvktxW/tM4+pTFR
'' SIG '' hLy/AsGConsXHRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0
'' SIG '' +CQ1ZyvgDbjmjJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAx
'' SIG '' M328y+l7vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQ
'' SIG '' wXEGahC0HVUzWLOhcGbyoYIC1DCCAj0CAQEwggEAoYHY
'' SIG '' pIHVMIHSMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
'' SIG '' aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
'' SIG '' ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQL
'' SIG '' EyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExp
'' SIG '' bWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjA4
'' SIG '' NDItNEJFNi1DMjlBMSUwIwYDVQQDExxNaWNyb3NvZnQg
'' SIG '' VGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoD
'' SIG '' FQBCoh8hiWMdRs2hjT/COFdGf+xIDaCBgzCBgKR+MHwx
'' SIG '' CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
'' SIG '' MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
'' SIG '' c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3
'' SIG '' DQEBBQUAAgUA6YELGzAiGA8yMDI0MDIyMjA3NTE1NVoY
'' SIG '' DzIwMjQwMjIzMDc1MTU1WjB0MDoGCisGAQQBhFkKBAEx
'' SIG '' LDAqMAoCBQDpgQsbAgEAMAcCAQACAgmDMAcCAQACAhKp
'' SIG '' MAoCBQDpglybAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwG
'' SIG '' CisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMB
'' SIG '' hqAwDQYJKoZIhvcNAQEFBQADgYEAKtGq9WhQL7wR2DFD
'' SIG '' OxXh2SPrabTYCUphee6aAOWPA+12OCg444nuxVh0F62j
'' SIG '' Sl8ReVFD/bzwKFPYbs2ZoZyXvqcVOKnyO7VXXUx+HHCm
'' SIG '' U+0Z7S4X30dgLQZIiQiO1YBMjAMUAgjWKJU+b8Nkpvfn
'' SIG '' gq51ZOWBT8Ad2DLB9F3boEgxggQNMIIECQIBATCBkzB8
'' SIG '' MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
'' SIG '' bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
'' SIG '' cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
'' SIG '' b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdqO
'' SIG '' 1claANERsQABAAAB2jANBglghkgBZQMEAgEFAKCCAUow
'' SIG '' GgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqG
'' SIG '' SIb3DQEJBDEiBCANvGOEcnB9cx7krfNWLxD8JOkRxKgH
'' SIG '' G6KQzD29/eG2rTCB+gYLKoZIhvcNAQkQAi8xgeowgecw
'' SIG '' geQwgb0EICKlo2liwO+epN73kOPULT3TbQjmWOJutb+d
'' SIG '' 0gI7GD3GMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAg
'' SIG '' UENBIDIwMTACEzMAAAHajtXJWgDREbEAAQAAAdowIgQg
'' SIG '' piHx0CFBPF5brSgqI65SjQTTifdIdCak5/R+mEoGeMUw
'' SIG '' DQYJKoZIhvcNAQELBQAEggIAEEtvAj9hAmT7VDK6F98e
'' SIG '' vx/d8zq35Z1uMFdKr5xyslCkloOYteHqpsSVdQKSnTGY
'' SIG '' G5HnHHW+xUEK1F8GHhBSSTzXYxidBk3XfhsA7wcFTAUU
'' SIG '' wtqv+nrSHsXM+nGXl4BnYoaI8A84x68VGxCnJHQzbtvJ
'' SIG '' f/OVAx0Aw1+N29vyhDN80EzThJDiDqO098Wh/ru4YIyJ
'' SIG '' GGZZeSxmjSmNkclXhv1RRFyxatwhBjezxln3QfRQdq7e
'' SIG '' tBhxBC6ndt5PikNrev5grIfErcUBfCejY+0ivPYwEYCI
'' SIG '' X8GA8AB3lT1g07oHB4B6afa2mk7RS0bLiqDDjnZ2Rebf
'' SIG '' zVM38tpc70tqlEVqyffLxogzKWJxDXsffmhQ9b6lLmLr
'' SIG '' eOiOEHYSrWEf4cqPIvsd+dB3fBUhgVqYQvmi0Iayholz
'' SIG '' aqij50WSmq3OVuRbDpbap0hfORRVPrjjggqYSAHJQMX0
'' SIG '' 8uwvAyfetZWC2pIz+INRpvl3tmb9JtyQmHmsi2d3fXzT
'' SIG '' LcFuNbbIS++YOX53U3aUAIHCKvbt8vSn5PKmrJRfkin2
'' SIG '' SoXhvc8ggGFI224XNASwhMdgbXPHqDPbI1ekwInEAdii
'' SIG '' BGX+oDMU6B+IS2y7Srgyhd3+QUWhyieTTuPLDlWAazTf
'' SIG '' WArQHDRg9K9Aghqv2aEU6YarFxuuneV7j6BDGosuzLRqGyQ=
'' SIG '' End signature block
