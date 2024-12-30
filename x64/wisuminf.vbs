' Windows Installer utility to manage the summary information stream
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of the database summary information methods

Option Explicit

Const msiOpenDatabaseModeReadOnly     = 0
Const msiOpenDatabaseModeTransact     = 1
Const msiOpenDatabaseModeCreate       = 3

Dim propList(19, 1)
propList( 1,0) = "Codepage"    : propList( 1,1) = "ANSI codepage of text strings in summary information only"
propList( 2,0) = "Title"       : propList( 2,1) = "Package type, e.g. Installation Database"
propList( 3,0) = "Subject"     : propList( 3,1) = "Product full name or description"
propList( 4,0) = "Author"      : propList( 4,1) = "Creator, typically vendor name"
propList( 5,0) = "Keywords"    : propList( 5,1) = "List of keywords for use by file browsers"
propList( 6,0) = "Comments"    : propList( 6,1) = "Description of purpose or use of package"
propList( 7,0) = "Template"    : propList( 7,1) = "Target system: Platform(s);Language(s)"
propList( 8,0) = "LastAuthor"  : propList( 8,1) = "Used for transforms only: New target: Platform(s);Language(s)"
propList( 9,0) = "Revision"    : propList( 9,1) = "Package code GUID, for transforms contains old and new info"
propList(11,0) = "Printed"     : propList(11,1) = "Date and time of installation image, same as Created if CD"
propList(12,0) = "Created"     : propList(12,1) = "Date and time of package creation"
propList(13,0) = "Saved"       : propList(13,1) = "Date and time of last package modification"
propList(14,0) = "Pages"       : propList(14,1) = "Minimum Windows Installer version required: Major * 100 + Minor"
propList(15,0) = "Words"       : propList(15,1) = "Source and Elevation flags: 1=short names, 2=compressed, 4=network image, 8=LUA package"
propList(16,0) = "Characters"  : propList(16,1) = "Used for transforms only: validation and error flags"
propList(18,0) = "Application" : propList(18,1) = "Application associated with file, ""Windows Installer"" for MSI"
propList(19,0) = "Security"    : propList(19,1) = "0=Read/write 2=Readonly recommended 4=Readonly enforced"

Dim iArg, iProp, property, value, message
Dim argCount:argCount = Wscript.Arguments.Count
If argCount > 0 Then If InStr(1, Wscript.Arguments(0), "?", vbTextCompare) > 0 Then argCount = 0
If (argCount = 0) Then
	message = "Windows Installer utility to manage summary information stream" &_
		vbNewLine & " 1st argument is the path to the storage file (installer package)" &_
		vbNewLine & " If no other arguments are supplied, summary properties will be listed" &_
		vbNewLine & " Subsequent arguments are property=value pairs to be updated" &_
		vbNewLine & " Either the numeric or the names below may be used for the property" &_
		vbNewLine & " Date and time fields use current locale format, or ""Now"" or ""Date""" &_
		vbNewLine & " Some properties have specific meaning for installer packages"
	For iProp = 1 To UBound(propList)
		property = propList(iProp, 0)
		If Not IsEmpty(property) Then
			message = message & vbNewLine & Right(" " & iProp, 2) & "  " & property & " - " & propLIst(iProp, 1)
		End If
	Next
	message = message & vbNewLine & vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."

	Wscript.Echo message
	Wscript.Quit 1
End If

' Connect to Windows Installer object
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : If CheckError("MSI.DLL not registered") Then Wscript.Quit 2

' Evaluate command-line arguments and open summary information
Dim cUpdate:cUpdate = 0 : If argCount > 1 Then cUpdate = 20
Dim sumInfo  : Set sumInfo = installer.SummaryInformation(Wscript.Arguments(0), cUpdate) : If CheckError(Empty) Then Wscript.Quit 2

' If only package name supplied, then list all properties in summary information stream
If argCount = 1 Then
	For iProp = 1 to UBound(propList)
		value = sumInfo.Property(iProp) : CheckError(Empty)
		If Not IsEmpty(value) Then message = message & vbNewLine & Right(" " & iProp, 2) & "  " &  propList(iProp, 0) & " = " & value
	Next
	Wscript.Echo message
	Wscript.Quit 0
End If

' Process property settings, combining arguments if equal sign has spaces before or after it
For iArg = 1 To argCount - 1
	property = property & Wscript.Arguments(iArg)
	Dim iEquals:iEquals = InStr(1, property, "=", vbTextCompare) 'Must contain an equals sign followed by a value
	If iEquals > 0 And iEquals <> Len(property) Then
		value = Right(property, Len(property) - iEquals)
		property = Left(property, iEquals - 1)
		If IsNumeric(property) Then
			iProp = CLng(property)
		Else  ' Lookup property name if numeric property ID not supplied
			For iProp = 1 To UBound(propList)
				If propList(iProp, 0) = property Then Exit For
			Next
		End If
		If iProp > UBound(propList) Then
			Wscript.Echo "Unknown summary property name: " & property
			sumInfo.Persist ' Note! must write even if error, else entire stream will be deleted
			Wscript.Quit 2
		End If
		If iProp = 11 Or iProp = 12 Or iProp = 13 Then
			If UCase(value) = "NOW"  Then value = Now
			If UCase(value) = "DATE" Then value = Date
			value = CDate(value)
		End If
		If iProp = 1 Or iProp = 14 Or iProp = 15 Or iProp = 16 Or iProp = 19 Then value = CLng(value)
		sumInfo.Property(iProp) = value : CheckError("Bad format for property value " & iProp)
		property = Empty
	End If
Next
If Not IsEmpty(property) Then
	Wscript.Echo "Arguments must be in the form: property=value  " & property
	sumInfo.Persist ' Note! must write even if error, else entire stream will be deleted
	Wscript.Quit 2
End If

' Write new property set. Note! must write even if error, else entire stream will be deleted
sumInfo.Persist : If CheckError("Error persisting summary property stream") Then Wscript.Quit 2
Wscript.Quit 0


Function CheckError(message)
	If Err = 0 Then Exit Function
	If IsEmpty(message) Then message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Dim errRec : Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText
	End If
	Wscript.Echo message
	CheckError = True
	Err.Clear
End Function

'' SIG '' Begin signature block
'' SIG '' MIImBQYJKoZIhvcNAQcCoIIl9jCCJfICAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' bn8llKyfjYiHNwaF/UnnU74Wl84HND+puok0mU7lHYug
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
'' SIG '' hvcNAQkEMSIEIJc5C67bYA975XEYqU5/mVAzg90fi/h3
'' SIG '' 7wwxxriTMM9sMDwGCisGAQQBgjcKAxwxLgwsc1BZN3hQ
'' SIG '' QjdoVDVnNUhIcll0OHJETFNNOVZ1WlJ1V1phZWYyZTIy
'' SIG '' UnM1ND0wWgYKKwYBBAGCNwIBDDFMMEqgJIAiAE0AaQBj
'' SIG '' AHIAbwBzAG8AZgB0ACAAVwBpAG4AZABvAHcAc6EigCBo
'' SIG '' dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vd2luZG93czAN
'' SIG '' BgkqhkiG9w0BAQEFAASCAQA3VJ4LJ5Bg0xRFwdalG5jZ
'' SIG '' r03k4ZTrrExz4MavxoxYAV9lISKHzJOL22jsyXUsNNF9
'' SIG '' hZ95+kJY/hBwcmGaeEvG+8mt70sevpG1Yi9gq5UJTQZf
'' SIG '' Pz972/nQBPuWJ/u9css34raJpdZIvrOSDue78h2i0Svm
'' SIG '' 2gNv3PMlVjlMgucuBb6UK51GhjMO5uzIzzVYCha3EVxZ
'' SIG '' 6g/ptxjO6qEECUPuUWkAo1+t58c//ufP4YQJlxkcTz5c
'' SIG '' jkm1v6DyFyd4bLD42rx0Z2KBGrEpu4YuhgGAdP75MiNY
'' SIG '' 9hUgtAoxqrUf6+geDdrBILeAnu7zTsX4M9FvKu8sB4Zu
'' SIG '' UD+CmlB9cBPFoYIXKTCCFyUGCisGAQQBgjcDAwExghcV
'' SIG '' MIIXEQYJKoZIhvcNAQcCoIIXAjCCFv4CAQMxDzANBglg
'' SIG '' hkgBZQMEAgEFADCCAVkGCyqGSIb3DQEJEAEEoIIBSASC
'' SIG '' AUQwggFAAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUD
'' SIG '' BAIBBQAEILoIg5ds+BWSGi5mIl9f5Q+FB/+oYosg0y4v
'' SIG '' U4LjdcYVAgZl1fuullMYEzIwMjQwMjIyMTA0NDMxLjQw
'' SIG '' NlowBIACAfSggdikgdUwgdIxCzAJBgNVBAYTAlVTMRMw
'' SIG '' EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
'' SIG '' b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
'' SIG '' b24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9w
'' SIG '' ZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVz
'' SIG '' IFRTUyBFU046M0JENC00QjgwLTY5QzMxJTAjBgNVBAMT
'' SIG '' HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WgghF4
'' SIG '' MIIHJzCCBQ+gAwIBAgITMwAAAeWPasDzPbQLowABAAAB
'' SIG '' 5TANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzET
'' SIG '' MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
'' SIG '' bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
'' SIG '' aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
'' SIG '' cCBQQ0EgMjAxMDAeFw0yMzEwMTIxOTA3MzVaFw0yNTAx
'' SIG '' MTAxOTA3MzVaMIHSMQswCQYDVQQGEwJVUzETMBEGA1UE
'' SIG '' CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEe
'' SIG '' MBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0w
'' SIG '' KwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRp
'' SIG '' b25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
'' SIG '' RVNOOjNCRDQtNEI4MC02OUMzMSUwIwYDVQQDExxNaWNy
'' SIG '' b3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkq
'' SIG '' hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAqXvgOtq7Y7os
'' SIG '' usk7cfJO871pdqL/943I/kwtZmuZQY04kw/AwjTxX3MF
'' SIG '' 9E81y5yt4hhLIkeOQwhQaa6HSs9Xn/b5QIsas3U/vuf1
'' SIG '' +r+Z3Ncw3UXOpo8d0oSUqd4lDxHpw/h2u7YbKaa3WusZ
'' SIG '' w17zTQJwPp3812iiaaR3X3pWo62NkUDVda74awUF5YeJ
'' SIG '' 7P8+WWpwz95ae2RAyzSUrTOYJ8f4G7uLWH4UNFHwXtbN
'' SIG '' Sv/szeOFV0+kB+rbNgIDxlUs2ASLNj68WaDH7MO65T8Y
'' SIG '' KEMruSUNwLD7+BWgS5I6XlyVCzJ1ZCMklftnbJX7UoLo
'' SIG '' bUlWqk/d2ko8A//i502qlHkch5vxNrUl+NFTNK/epKN7
'' SIG '' nL1FhP8CNY1hDuCx7O4NYz/xxnXWRyjUm9TI5DzH8kOQ
'' SIG '' wWpJHCPW/6ZzosoqWP/91YIb8fD2ml2VYlfqmwN6xC5B
'' SIG '' HsVXt4KpX+V9qOguk83H/3MXV2/zJcF3OZYk94KJ7Zdp
'' SIG '' CesAoOqUbfNe7H201CbPYv3pN3Gcg7Y4aZjEEABkBagp
'' SIG '' ua1gj4KLPgJjI7QWblndPjRrl3som5+0XoJOhxxz9Sn+
'' SIG '' OkV9CK0t+N3vVxL5wsJ6hD6rSfQgAu9X5pxsQ2i5I6uO
'' SIG '' /9C1jgUiMeUjnN0nMmcayRUnmjOGOLRhGxy/VbAkUC7L
'' SIG '' IIxC8t2Y910CAwEAAaOCAUkwggFFMB0GA1UdDgQWBBTf
'' SIG '' /5+Hu01zMSJ8ReUJCAU5eAyHqjAfBgNVHSMEGDAWgBSf
'' SIG '' pxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSg
'' SIG '' UqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
'' SIG '' b3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIw
'' SIG '' UENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBe
'' SIG '' MFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29m
'' SIG '' dC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRp
'' SIG '' bWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNV
'' SIG '' HRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
'' SIG '' MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOC
'' SIG '' AgEAM/rCE4WMPVp3waQQn2gsG69+Od0zIZD1HgeAEpKU
'' SIG '' +3elrRdUtyKasmUOcoaAUGJbAjpc6DDzaF2iUOIwMEst
'' SIG '' ZExMkdZV5RjWBThcHw44jEFz39DzfNvVzRFYS6mALjwj
'' SIG '' 5v7bHZU2AYlSxAjI9HY+JdCFPk/J6syBqD05Kh1CMXCk
'' SIG '' 10aKudraulXbcRTAV47n7ehJfgl4I1m+DJQ7MqnIy+pV
'' SIG '' q5uj4aV/+mx9bm0hwyNlW3R6WzB+rSok1CChiKltpO+/
'' SIG '' vGaLFQkZNuLFiJ9PACK89wo116Kxma22zs4dsAzv3lm8
'' SIG '' otISpeJFSMNhnJ4fIDKwwQAtsiF1eAcSHrQqhnLOUFfP
'' SIG '' dXESKsTueG5w3Aza1WI6XAjsSR5TmG51y2dcIbnkm4zD
'' SIG '' /BvtzvVEqKZkD8peVamYG+QmQHQFkRLw4IYN37Nj9P0G
'' SIG '' dOnyyLfpOqXzhV+lh72IebLs+qrGowXYKfirZrSYQyek
'' SIG '' Gu4MYT+BH1zxJUnae2QBHLlJ+W64n8wHrXJG9PWZTHeX
'' SIG '' Kmk7bZ4+MGOfCgS9XFsONPWOF0w116864N4kbNEsr0c2
'' SIG '' ZMML5N1lCWP5UyAibxl4QhE0XShq+IX5BlxRktbNZtir
'' SIG '' rIOiTwRkoWJFHmi0GgYu9pgWnEFlQTyacsq4OVihuOvG
'' SIG '' HuWfCvFX98zLQX19KjYnEWa0uC0wggdxMIIFWaADAgEC
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
'' SIG '' bWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjNC
'' SIG '' RDQtNEI4MC02OUMzMSUwIwYDVQQDExxNaWNyb3NvZnQg
'' SIG '' VGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoD
'' SIG '' FQD3jaIa5gWuwTjDNYN3zkSkzpGLCqCBgzCBgKR+MHwx
'' SIG '' CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
'' SIG '' MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
'' SIG '' c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3
'' SIG '' DQEBBQUAAgUA6YEivjAiGA8yMDI0MDIyMjA5MzI0NloY
'' SIG '' DzIwMjQwMjIzMDkzMjQ2WjB0MDoGCisGAQQBhFkKBAEx
'' SIG '' LDAqMAoCBQDpgSK+AgEAMAcCAQACAgJdMAcCAQACAhJ4
'' SIG '' MAoCBQDpgnQ+AgEAMDYGCisGAQQBhFkKBAIxKDAmMAwG
'' SIG '' CisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMB
'' SIG '' hqAwDQYJKoZIhvcNAQEFBQADgYEAPiL5H2aaxnvZgzld
'' SIG '' ovmjgJ3Vz2yWpdjZxj2dLPIhLiWmCJ6MshDzEJJdBE/o
'' SIG '' mXyOglQMaJsekvMM6MDRsZYdxx9I3tmibmmIiKFvt1C5
'' SIG '' azPMD7NAVpD4HIwmbZPsZAeLGTHCCiiTfKgQbZyqJd13
'' SIG '' YYyH/x4N2tlqTU89PnqRD9IxggQNMIIECQIBATCBkzB8
'' SIG '' MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
'' SIG '' bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
'' SIG '' cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
'' SIG '' b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAeWP
'' SIG '' asDzPbQLowABAAAB5TANBglghkgBZQMEAgEFAKCCAUow
'' SIG '' GgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqG
'' SIG '' SIb3DQEJBDEiBCB9xdx1xVCWhWcRFlte8hW1APp7P9+s
'' SIG '' 1avxmMgaXmm/VzCB+gYLKoZIhvcNAQkQAi8xgeowgecw
'' SIG '' geQwgb0EIBWp0//+qPEYWF7ZhugRd5vwj+kCh/TULCFv
'' SIG '' FQf1Tr3tMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAg
'' SIG '' UENBIDIwMTACEzMAAAHlj2rA8z20C6MAAQAAAeUwIgQg
'' SIG '' dOPx9eZl7xazJUHS7mjI3NkEavT4dl3cK1osxAyj76sw
'' SIG '' DQYJKoZIhvcNAQELBQAEggIANb908yy8dfCWo9WCREYx
'' SIG '' s7ycbqe7hHA8pa1O0z/kXtdQKPFR8qtU4sebukULDk5/
'' SIG '' uOFFREU2oQCyYM/jgYmYcRTJs+aify7ndyA/2RHuBdRS
'' SIG '' Sj9HNxidwcmYwS0ypH96GlUxgd8rR3sGaxLvibpwOB80
'' SIG '' TDZmZQLaBfTOP5N4/vBbxCCTC5cWoX6I1YGTAnFAVSun
'' SIG '' 7QeamDqVV1pnMa4EPbAnsR6HuXnVHXnaTzzAvb3bxddv
'' SIG '' r+N9b9npEtgwpk8evhQeGUeGEuzjcpNcUJFmS0bwZ7VO
'' SIG '' /NaO/48l+IMd4eAlamor7XJ4MmeZp3Wyj3PGV2j7gf5A
'' SIG '' baf6d/Lj24h6W8gHz9NiT28zTgW2g7zVgR/f235o4rSy
'' SIG '' FI4JFQ/QJIUsUcve2rjXFBTZ1pn7DNwlLQ7KzLWPO41Q
'' SIG '' WhH1beFq0r3/et53gR21svbVtfPJXFDlZ15iVuvXYhKl
'' SIG '' sZ8XAniQZzy7oNxbgo02GsCzsL5eOTGsidbTfz11sy22
'' SIG '' MoPLS4II2WKb0asbO00oOVWsBfEIK4VjfW+TgDmFSqp6
'' SIG '' ++vdC2V+5Zwu3R9ayOFnDmtE09EDbXxY4JB63Gd+zKs5
'' SIG '' sqtaHDw07Q3ZWjfZkrIG2tXxCIFQxZ/0C4o1eHIHIK0L
'' SIG '' r1gayG0B5X5PI51sXioIRl0HyaW4fokvZulY4cSmqgxoOVM=
'' SIG '' End signature block
