'******************************************************************************
'Microsoft Confidential. © 2002-2003 Microsoft Corporation. All rights reserved.
'
' This file may contain preliminary information or inaccuracies, 
' and may not correctly represent any associated Microsoft 
' Product as commercially released. All Materials are provided entirely 
' “AS IS.” To the extent permitted by law, MICROSOFT MAKES NO 
' WARRANTY OF ANY KIND, DISCLAIMS ALL EXPRESS, IMPLIED AND STATUTORY 
' WARRANTIES, AND ASSUMES NO LIABILITY TO YOU FOR ANY DAMAGES OF 
' ANY TYPE IN CONNECTION WITH THESE MATERIALS OR ANY INTELLECTUAL PROPERTY IN THEM. 
'******************************************************************************

Option Explicit

Wscript.Echo "" 
Wscript.Echo "REGISTER_APP.VBS version 1.6 for Windows Server 2008"
Wscript.Echo "Copyright (C) Microsoft Corporation 2002-2003. All rights reserved."
Wscript.Echo "" 


'******************************************************************************
' Parse command line arguments
'******************************************************************************
Dim Args
Set Args = Wscript.Arguments
If Args.Count < 1 Then 
	PrintsUsage
End If

Dim ProviderName, ProviderDLL, ProviderDescription
If Args.Item(0) = "-register" Then 
	If Args.Count <> 4 Then PrintsUsage

	ProviderName = Args.Item(1)
	ProviderDLL = Args.Item(2)
	ProviderDescription = Args.Item(3)

	UninstallProvider
	InstallProvider
	Wscript.Quit 0
End If 

If Args.Item(0) = "-unregister" Then 
	If Not Args.Count = 2 Then PrintsUsage
	ProviderName = Args.Item(1)
	UninstallProvider
	Wscript.Quit 0
End If

' Wrong options?
PrintsUsage

Wscript.Quit 0

'******************************************************************************
' Prints the usage
'******************************************************************************
Sub PrintsUsage

	Wscript.Echo "Usage:" 
	Wscript.Echo "" 
	Wscript.Echo " 1) Registering a VSS/VDS Provider as a COM+ application:" 
	Wscript.Echo "      CScript.exe " & Wscript.ScriptName & " -register <Provider_Name> <Provider.DLL>  <Provider_Description>" 
	Wscript.Echo "" 
	Wscript.Echo " 2) Unregistering a COM+ application associated with a VSS/VDS provider:" 
	Wscript.Echo "      CScript.exe " & Wscript.ScriptName & " -unregister <Provider_Name>" 
	Wscript.Echo "" 
	Wscript.Quit 1

End Sub


'******************************************************************************
' Installs the Provider
'******************************************************************************
Sub InstallProvider
	On Error Resume Next

	Wscript.Echo "Creating a new COM+ application:" 

	Wscript.Echo "- Creating the catalog object "
	Dim cat
	Set cat = CreateObject("COMAdmin.COMAdminCatalog") 	
	CheckError 101

	wscript.echo "- Get the Applications collection"
	Dim collApps
	Set collApps = cat.GetCollection("Applications")
	CheckCollectionError 102, cat

	Wscript.Echo "- Populate..." 
	collApps.Populate 
	CheckCollectionError 103, collApps

	Wscript.Echo "- Add new application object" 
	Dim app
	Set app = collApps.Add 
	CheckCollectionError 104, collApps

	Wscript.Echo "- Set app name = " & ProviderName & " "
	app.Value("Name") = ProviderName
	CheckObjectError 105, collApps, app

	Wscript.Echo "- Set app description = " & ProviderDescription & " "
	app.Value("Description") = ProviderDescription 
	CheckObjectError 106, collApps, app

	' Only roles added below are allowed to call in.
	Wscript.Echo "- Set app access check = true "
	app.Value("ApplicationAccessChecksEnabled") = 1   
	CheckObjectError 107, collApps, app

	' Encrypting communication
	Wscript.Echo "- Set encrypted COM communication = true "
	app.Value("Authentication") = 6	                  
	CheckObjectError 108, collApps, app

	' Secure references
	Wscript.Echo "- Set secure references = true "
	app.Value("AuthenticationCapability") = 2         
	CheckObjectError 109, collApps, app

	' Do not allow impersonation
	Wscript.Echo "- Set impersonation = false "
	app.Value("ImpersonationLevel") = 2               
	CheckObjectError 110, collApps, app

	Wscript.Echo "- Save changes..."
	collApps.SaveChanges
	CheckCollectionError 111, collApps

	wscript.echo "- Create Windows service running as Local System"
	cat.CreateServiceForApplication ProviderName, ProviderName , "SERVICE_AUTO_START", "SERVICE_ERROR_NORMAL", "", ".\localsystem", "", 0
	CheckCollectionError 112, cat

	wscript.echo "- Add the DLL component"
	cat.InstallComponent ProviderName, ProviderDLL , "", ""
        CheckCollectionError 113, cat

	'
	' Add the new role for the Local SYSTEM account
	'

	wscript.echo "Secure the COM+ application:"
	wscript.echo "- Get roles collection"
	Dim collRoles
	Set collRoles = collApps.GetCollection("Roles", app.Key)
	CheckCollectionError 120, cat

	wscript.echo "- Populate..."
	collRoles.Populate
	CheckCollectionError 121, collRoles

	wscript.echo "- Add new role"
	Dim role
	Set role = collRoles.Add
	CheckCollectionError 122, collRoles

	wscript.echo "- Set name = Administrators "
	role.Value("Name") = "Administrators"
	CheckObjectError 123, collRoles, role

	wscript.echo "- Set description = Administrators group "
	role.Value("Description") = "Administrators group"
	CheckObjectError 124, collRoles, role

	wscript.echo "- Save changes ..."
	collRoles.SaveChanges
	CheckCollectionError 125, collRoles
	
	'
	' Add users into role
	'

	wscript.echo "Granting user permissions:"
	Dim collUsersInRole
	Set collUsersInRole = collRoles.GetCollection("UsersInRole", role.Key)
	CheckCollectionError 130, collRoles

	wscript.echo "- Populate..."
	collUsersInRole.Populate
	CheckCollectionError 131, collUsersInRole

	wscript.echo "- Add new user"
	Dim user
	Set user = collUsersInRole.Add
	CheckCollectionError 132, collUsersInRole

	wscript.echo "- Searching for the Administrators account using WMI..."

	' Get the Administrators account domain and name
	Dim strQuery
	strQuery = "select * from Win32_Account where SID='S-1-5-32-544' and localAccount=TRUE"
	Dim objSet
	set objSet = GetObject("winmgmts:").ExecQuery(strQuery)
	CheckError 133

	Dim obj, Account
	for each obj in objSet
	    set Account = obj
		exit for
	next

	wscript.echo "- Set user name = .\" & Account.Name & " "
	user.Value("User") = ".\" & Account.Name
	CheckObjectError 140, collUsersInRole, user

	wscript.echo "- Add new user"
	Set user = collUsersInRole.Add
	CheckCollectionError 141, collUsersInRole

	wscript.echo "- Set user name = Local SYSTEM "
	user.Value("User") = "NT AUTHORITY\SYSTEM"
	CheckObjectError 142, collUsersInRole, user

	wscript.echo "- Save changes..."
	collUsersInRole.SaveChanges
	CheckCollectionError 143, collUsersInRole
	
	Set app      = Nothing
	Set cat      = Nothing
	Set role     = Nothing
	Set user     = Nothing

	Set collApps = Nothing
	Set collRoles = Nothing
	Set collUsersInRole	= Nothing

	set objSet   = Nothing
	set obj      = Nothing

	Wscript.Echo "Done." 

	On Error GoTo 0
End Sub


'******************************************************************************
' Uninstalls the Provider
'******************************************************************************
Sub UninstallProvider
	On Error Resume Next

	Wscript.Echo "Unregistering the existing application..." 

	wscript.echo "- Create the catalog object"
	Dim cat
	Set cat = CreateObject("COMAdmin.COMAdminCatalog")
	CheckError 201
	
	wscript.echo "- Get the Applications collection"
	Dim collApps
	Set collApps = cat.GetCollection("Applications")
	CheckCollectionError 202, cat

	wscript.echo "- Populate..."
	collApps.Populate
	CheckCollectionError 203, collApps
	
	wscript.echo "- Search for " & ProviderName & " application..."
	Dim numApps
	numApps = collApps.Count
	Dim i
	For i = numApps - 1 To 0 Step -1
	    If collApps.Item(i).Value("Name") = ProviderName Then
	        collApps.Remove(i)
		CheckCollectionError 204, collApps
                WScript.echo "- Application " & ProviderName & " removed!"
	    End If
	Next
	
	wscript.echo "- Saving changes..."
	collApps.SaveChanges
	CheckCollectionError 205, collApps

	Set collApps = Nothing
	Set cat      = Nothing

	Wscript.Echo "Done." 

	On Error GoTo 0
End Sub



'******************************************************************************
' Sub CheckError
'******************************************************************************
Sub CheckError(exitCode)
    If Err = 0 Then Exit Sub
    DumpVBScriptError exitCode

    Wscript.Quit exitCode
End Sub


'******************************************************************************
' Sub CheckCollectionError
'******************************************************************************
Sub CheckCollectionError(exitCode, coll)
    If Err = 0 Then Exit Sub
    DumpVBScriptError exitCode

    DumpComPlusError(coll.GetCollection("ErrorInfo"))

    Wscript.Quit exitCode
End Sub


'******************************************************************************
' Sub CheckObjectError
'******************************************************************************
Sub CheckObjectError(exitCode, coll, object)
    If Err = 0 Then Exit Sub
    DumpVBScriptError exitCode

    ' DumpComPlusError(coll.GetCollection("ErrorInfo", object.Key))
    DumpComPlusError(coll.GetCollection("ErrorInfo"))

    Wscript.Quit exitCode
End Sub



'******************************************************************************
' Sub DumpVBScriptError
'******************************************************************************
Sub DumpVBScriptError(exitCode)
    WScript.Echo vbNewLine & "ERROR:"
    WScript.Echo "- Error code: " & Err & " [0x" & Hex(Err) & "]"
    WScript.Echo "- Exit code: " & exitCode
    WScript.Echo "- Description: " & Err.Description
    WScript.Echo "- Source: " & Err.Source
    WScript.Echo "- Help file: " & Err.Helpfile
    WScript.Echo "- Help context: " & Err.HelpContext
End Sub


'******************************************************************************
' Sub DumpComPlusError
'******************************************************************************
Sub DumpComPlusError(errors)
    errors.Populate
    WScript.Echo "- COM+ Errors detected: (" & errors.Count & ")"

    Dim error
    Dim I
    For I = 0 to errors.Count - 1
	Set error = errors.Item(I)
        WScript.Echo "   * (COM+ ERROR " & I & ") on " & error.Value("Name")
        WScript.Echo "       ErrorCode: " & error.Value("ErrorCode") & " [0x" & Hex(error.Value("ErrorCode")) & "]"
        WScript.Echo "       MajorRef: " & error.Value("MajorRef")
        WScript.Echo "       MinorRef: " & error.Value("MinorRef")
    Next
End Sub


'' SIG '' Begin signature block
'' SIG '' MIImBQYJKoZIhvcNAQcCoIIl9jCCJfICAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' t2OGjVuwrDi7m9eD1oGHZt1e8mT97G6PYHdAzoXpmRWg
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
'' SIG '' hvcNAQkEMSIEIKK6iEgLHrRNtkju0ARNJFYXKWBPVpnu
'' SIG '' yofrkEcDFCOSMDwGCisGAQQBgjcKAxwxLgwsc1BZN3hQ
'' SIG '' QjdoVDVnNUhIcll0OHJETFNNOVZ1WlJ1V1phZWYyZTIy
'' SIG '' UnM1ND0wWgYKKwYBBAGCNwIBDDFMMEqgJIAiAE0AaQBj
'' SIG '' AHIAbwBzAG8AZgB0ACAAVwBpAG4AZABvAHcAc6EigCBo
'' SIG '' dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vd2luZG93czAN
'' SIG '' BgkqhkiG9w0BAQEFAASCAQBlUHfLcEGiXGi3td95FZjP
'' SIG '' +6SvVIc0fMEkaVDhgLCJQT9Z2sv2ioBQ6yPNKI6pS5C+
'' SIG '' ZsIB1Dj2IeGofhWVNTiwFudYOCpA4IjGKrCzBk+B+idp
'' SIG '' OFsZlBuKHdEGFBRPPg07kWHnLLIZNSzTTm2RDgFiKp3Q
'' SIG '' 1JOdC4oNknYflibZc2c1kow6UsMS5J1SEkhooo+7m0i8
'' SIG '' Gee3JG29v4amsf3JNLBRocTEjFR6lp2Dm1Umj0wmW2+R
'' SIG '' 6ZvdzpK1Ubx5DpqC+YIz6JKUzt7SGBrxmuEAonsJtYiS
'' SIG '' dYT9pIluQnr6leZM5Ou+NOKAPxMP5KNJqxCYJRoIAlJY
'' SIG '' AYeN0wtVA5quoYIXKTCCFyUGCisGAQQBgjcDAwExghcV
'' SIG '' MIIXEQYJKoZIhvcNAQcCoIIXAjCCFv4CAQMxDzANBglg
'' SIG '' hkgBZQMEAgEFADCCAVkGCyqGSIb3DQEJEAEEoIIBSASC
'' SIG '' AUQwggFAAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUD
'' SIG '' BAIBBQAEIHFJtPAM1GirtqP7reh9X/Tvh2xNiw7SFb9W
'' SIG '' KEsBGVIyAgZl1eP46NMYEzIwMjQwMjIyMTA0MjI2LjU4
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
'' SIG '' SIb3DQEJBDEiBCB1oSpn8z+6FuwmSRNSZsI7UBoCwCRZ
'' SIG '' bYdhnmBEA3+/1jCB+gYLKoZIhvcNAQkQAi8xgeowgecw
'' SIG '' geQwgb0EICKlo2liwO+epN73kOPULT3TbQjmWOJutb+d
'' SIG '' 0gI7GD3GMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAg
'' SIG '' UENBIDIwMTACEzMAAAHajtXJWgDREbEAAQAAAdowIgQg
'' SIG '' piHx0CFBPF5brSgqI65SjQTTifdIdCak5/R+mEoGeMUw
'' SIG '' DQYJKoZIhvcNAQELBQAEggIAXJnajRFxNQD+oBbBe/Kt
'' SIG '' 0kEXvVGa1b98ve1hoPuvN9oo+e3ifpw6ae61p8Ck+f9C
'' SIG '' 4mx+vO2mLNynwwaMfdOLEx09plnRH6VHOQrKyDUvPYq7
'' SIG '' 1BD2YZsdngEBCnWOXxoda09US2NwlyLr6wXYTza3wW8K
'' SIG '' dMMfPH8EBLZCpXBDk3X6kt3E3f4Go3V8K1X9M1bKaRev
'' SIG '' IJfZG9Vf9Mh0pXflSCGGcczsi6QlKLAq0UqkJzi0KszI
'' SIG '' s+Kb4WsBQVghNwzhHNTKdUHqvS3Qv8L6Lzu1xNTGx5iC
'' SIG '' eOljQdezDbP1m7pYjb+I5jhu8Rr8/4FpTvftNiUTocTL
'' SIG '' KXyeX/1g51JPTMdvEoiIuqhsYPd/guBIrBmzPAjOVfgr
'' SIG '' 0hkZ769NOooDkP7Xdq1k1SxhcaXHOYZdXBwBpgI2aT4g
'' SIG '' rw5se0vk6/n1Cb+QDdQAuU6WY/EiewIZaF2N4Dfl6lsc
'' SIG '' mCsolGj6TFsIXV+0Ee4I0v3Urazx1L7PQtGC0njd6pol
'' SIG '' 4Kxdfw1buOMrpQxX7bH+w60ItrU0nOSmFDC9MMdY5xHf
'' SIG '' mdmyP8PJ9FYY9L2Gh2TCEhQxhGR+KqU7kYUfnqpVHwbI
'' SIG '' W6FTvhOAZ4ICOAYFS3zYUmoWNJLkhzMyoF2b8BQxqwtR
'' SIG '' wvuuSsDSIzDb04Y+Hk+OxEYBJ9VGwrdLxLA9xVHim/Ugsh0=
'' SIG '' End signature block
