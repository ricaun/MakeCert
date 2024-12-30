' Windows Installer transform viewer for use with Windows Scripting Host
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of the database APIs for viewing transform files
'
Option Explicit

Const iteAddExistingRow      = 1
Const iteDelNonExistingRow   = 2
Const iteAddExistingTable    = 4
Const iteDelNonExistingTable = 8
Const iteUpdNonExistingRow   = 16
Const iteChangeCodePage      = 32
Const iteViewTransform       = 256

Const icdLong       = 0
Const icdShort      = &h400
Const icdObject     = &h800
Const icdString     = &hC00
Const icdNullable   = &h1000
Const icdPrimaryKey = &h2000
Const icdNoNulls    = &h0000
Const icdPersistent = &h0100
Const icdTemporary  = &h0000

Const idoReadOnly = 0

Dim gErrors, installer, base, database, argCount, arg, argValue
gErrors = iteAddExistingRow + iteDelNonExistingRow + iteAddExistingTable + iteDelNonExistingTable + iteUpdNonExistingRow + iteChangeCodePage
Set database = Nothing

' Check arg count, and display help if no all arguments present
argCount = WScript.Arguments.Count
If (argCount < 2) Then
	WScript.Echo "Windows Installer Transform Viewer for Windows Scripting Host (CScript.exe)" &_
		vbNewLine & " 1st non-numeric argument is path to base database which transforms reference" &_
		vbNewLine & " Subsequent non-numeric arguments are paths to the transforms to be viewed" &_
		vbNewLine & " Numeric argument is optional error suppression flags (default is ignore all)" &_
		vbNewLine & " Arguments are executed left-to-right, as encountered" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Cannot run with GUI script host, as listing is performed to standard out
If UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "W" Then
	WScript.Echo "Cannot use WScript.exe - must use CScript.exe with this program"
	Wscript.Quit 2
End If

' Create installer object
On Error Resume Next
Set installer = CreateObject("WindowsInstaller.Installer") : CheckError

' Process arguments, opening database and applying transforms
For arg = 0 To argCount - 1
	argValue = WScript.Arguments(arg)
	If IsNumeric(argValue) Then
		gErrors = argValue
	ElseIf database Is Nothing Then
		Set database = installer.OpenDatabase(argValue, idoReadOnly)
	Else
		database.ApplyTransform argValue, iteViewTransform + gErrors
	End If
	CheckError
Next
ListTransform(database)

Function DecodeColDef(colDef)
	Dim def
	Select Case colDef AND (icdShort OR icdObject)
	Case icdLong
		def = "LONG"
	Case icdShort
		def = "SHORT"
	Case icdObject
		def = "OBJECT"
	Case icdString
		def = "CHAR(" & (colDef AND 255) & ")"
	End Select
	If (colDef AND icdNullable)   =  0 Then def = def & " NOT NULL"
	If (colDef AND icdPrimaryKey) <> 0 Then def = def & " PRIMARY KEY"
	DecodeColDef = def
End Function

Sub ListTransform(database)
	Dim view, record, row, column, change
	On Error Resume Next
	Set view = database.OpenView("SELECT * FROM `_TransformView` ORDER BY `Table`, `Row`") : CheckError
	view.Execute : CheckError
	Do
		Set record = view.Fetch : CheckError
		If record Is Nothing Then Exit Do
		change = Empty
		If record.IsNull(3) Then
			row = "<DDL>"
			If NOT record.IsNull(4) Then change = "[" & record.StringData(5) & "]: " & DecodeColDef(record.StringData(4))
		Else
			row = "[" & Join(Split(record.StringData(3), vbTab, -1), ",") & "]"
			If record.StringData(2) <> "INSERT" AND record.StringData(2) <> "DELETE" Then change = "{" & record.StringData(5) & "}->{" & record.StringData(4) & "}"
		End If
		column = record.StringData(1) & " " & record.StringData(2)
		if Len(column) < 24 Then column = column & Space(24 - Len(column))
		WScript.Echo column, row, change
	Loop
End Sub

Sub CheckError
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText
	End If
	Wscript.Echo message
	Wscript.Quit 2
End Sub

'' SIG '' Begin signature block
'' SIG '' MIImAwYJKoZIhvcNAQcCoIIl9DCCJfACAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' bSE4J2+6sdFrt3HJs8WbQ0/DQ6WkcvTESWq4fOF4kUmg
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
'' SIG '' jgd7JXFEqwZq5tTG3yOalnXFMYIZ9DCCGfACAQEwgZUw
'' SIG '' fjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWlj
'' SIG '' cm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMAITMwAA
'' SIG '' BVfPkN3H0cCIjAAAAAAFVzANBglghkgBZQMEAgEFAKCC
'' SIG '' AQQwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
'' SIG '' KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
'' SIG '' hvcNAQkEMSIEIPzOCw65ATLvFc2MyvnVG32kO/C83VNT
'' SIG '' MF2K5qWy6t0vMDwGCisGAQQBgjcKAxwxLgwsc1BZN3hQ
'' SIG '' QjdoVDVnNUhIcll0OHJETFNNOVZ1WlJ1V1phZWYyZTIy
'' SIG '' UnM1ND0wWgYKKwYBBAGCNwIBDDFMMEqgJIAiAE0AaQBj
'' SIG '' AHIAbwBzAG8AZgB0ACAAVwBpAG4AZABvAHcAc6EigCBo
'' SIG '' dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vd2luZG93czAN
'' SIG '' BgkqhkiG9w0BAQEFAASCAQBQUKwCY0IbNKsB59xJLxWU
'' SIG '' Zho7YNr6T3dlgQ0RVWUsR+M7wDzjhpcqW1tGOhyNqesw
'' SIG '' RUIeboNPQ0YTpL7OG8BOocKdcS3piIxgAziRGrPZ7JZd
'' SIG '' jsIPJatq0bZD/PlxziSq2D7HvSGHld7FIK1fPisrump3
'' SIG '' 4mXoYL9LJZMdYpACzqzLKroC00YMW3Ns3kGbSHUAyKJo
'' SIG '' h+8Nyxz3zWU0W2AAkJIX53EvoKN4miLtWn8GdXdprBWH
'' SIG '' vs9M5nbeFHmGnkDxK+H7i9szpHa2dD5e7b8HUdvulYz3
'' SIG '' z3r/Olt3XKZN1FUTx8q2Jkvl++IvgaS5x8MQxy5s1+Ks
'' SIG '' lkE4UaZEU4P0oYIXJzCCFyMGCisGAQQBgjcDAwExghcT
'' SIG '' MIIXDwYJKoZIhvcNAQcCoIIXADCCFvwCAQMxDzANBglg
'' SIG '' hkgBZQMEAgEFADCCAVcGCyqGSIb3DQEJEAEEoIIBRgSC
'' SIG '' AUIwggE+AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUD
'' SIG '' BAIBBQAEIHdfWlw6WHc/nRVLtSmQ3ZA+d1J+UjXq7PzU
'' SIG '' xliqrN0kAgZl1fuulPQYETIwMjQwMjIyMTA0NDIzLjFa
'' SIG '' MASAAgH0oIHYpIHVMIHSMQswCQYDVQQGEwJVUzETMBEG
'' SIG '' A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
'' SIG '' ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
'' SIG '' MS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVy
'' SIG '' YXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBU
'' SIG '' U1MgRVNOOjNCRDQtNEI4MC02OUMzMSUwIwYDVQQDExxN
'' SIG '' aWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloIIReDCC
'' SIG '' BycwggUPoAMCAQICEzMAAAHlj2rA8z20C6MAAQAAAeUw
'' SIG '' DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAg
'' SIG '' UENBIDIwMTAwHhcNMjMxMDEyMTkwNzM1WhcNMjUwMTEw
'' SIG '' MTkwNzM1WjCB0jELMAkGA1UEBhMCVVMxEzARBgNVBAgT
'' SIG '' Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
'' SIG '' BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsG
'' SIG '' A1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9u
'' SIG '' cyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVT
'' SIG '' TjozQkQ0LTRCODAtNjlDMzElMCMGA1UEAxMcTWljcm9z
'' SIG '' b2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZI
'' SIG '' hvcNAQEBBQADggIPADCCAgoCggIBAKl74Drau2O6LLrJ
'' SIG '' O3HyTvO9aXai//eNyP5MLWZrmUGNOJMPwMI08V9zBfRP
'' SIG '' NcucreIYSyJHjkMIUGmuh0rPV5/2+UCLGrN1P77n9fq/
'' SIG '' mdzXMN1FzqaPHdKElKneJQ8R6cP4dru2Gymmt1rrGcNe
'' SIG '' 800CcD6d/Ndoommkd196VqOtjZFA1XWu+GsFBeWHiez/
'' SIG '' PllqcM/eWntkQMs0lK0zmCfH+Bu7i1h+FDRR8F7WzUr/
'' SIG '' 7M3jhVdPpAfq2zYCA8ZVLNgEizY+vFmgx+zDuuU/GChD
'' SIG '' K7klDcCw+/gVoEuSOl5clQsydWQjJJX7Z2yV+1KC6G1J
'' SIG '' VqpP3dpKPAP/4udNqpR5HIeb8Ta1JfjRUzSv3qSje5y9
'' SIG '' RYT/AjWNYQ7gsezuDWM/8cZ11kco1JvUyOQ8x/JDkMFq
'' SIG '' SRwj1v+mc6LKKlj//dWCG/Hw9ppdlWJX6psDesQuQR7F
'' SIG '' V7eCqV/lfajoLpPNx/9zF1dv8yXBdzmWJPeCie2XaQnr
'' SIG '' AKDqlG3zXux9tNQmz2L96TdxnIO2OGmYxBAAZAWoKbmt
'' SIG '' YI+Ciz4CYyO0Fm5Z3T40a5d7KJuftF6CToccc/Up/jpF
'' SIG '' fQitLfjd71cS+cLCeoQ+q0n0IALvV+acbENouSOrjv/Q
'' SIG '' tY4FIjHlI5zdJzJnGskVJ5ozhji0YRscv1WwJFAuyyCM
'' SIG '' QvLdmPddAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQU3/+f
'' SIG '' h7tNczEifEXlCQgFOXgMh6owHwYDVR0jBBgwFoAUn6cV
'' SIG '' XQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
'' SIG '' UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9w
'' SIG '' cy9jcmwvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBD
'' SIG '' QSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAwXjBc
'' SIG '' BggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQu
'' SIG '' Y29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBUaW1l
'' SIG '' LVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQwDAYDVR0T
'' SIG '' AQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
'' SIG '' BgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIB
'' SIG '' ADP6whOFjD1ad8GkEJ9oLBuvfjndMyGQ9R4HgBKSlPt3
'' SIG '' pa0XVLcimrJlDnKGgFBiWwI6XOgw82hdolDiMDBLLWRM
'' SIG '' TJHWVeUY1gU4XB8OOIxBc9/Q83zb1c0RWEupgC48I+b+
'' SIG '' 2x2VNgGJUsQIyPR2PiXQhT5PyerMgag9OSodQjFwpNdG
'' SIG '' irna2rpV23EUwFeO5+3oSX4JeCNZvgyUOzKpyMvqVaub
'' SIG '' o+Glf/psfW5tIcMjZVt0elswfq0qJNQgoYipbaTvv7xm
'' SIG '' ixUJGTbixYifTwAivPcKNdeisZmtts7OHbAM795ZvKLS
'' SIG '' EqXiRUjDYZyeHyAysMEALbIhdXgHEh60KoZyzlBXz3Vx
'' SIG '' EirE7nhucNwM2tViOlwI7EkeU5hudctnXCG55JuMw/wb
'' SIG '' 7c71RKimZA/KXlWpmBvkJkB0BZES8OCGDd+zY/T9BnTp
'' SIG '' 8si36Tql84VfpYe9iHmy7PqqxqMF2Cn4q2a0mEMnpBru
'' SIG '' DGE/gR9c8SVJ2ntkARy5SfluuJ/MB61yRvT1mUx3lypp
'' SIG '' O22ePjBjnwoEvVxbDjT1jhdMNdevOuDeJGzRLK9HNmTD
'' SIG '' C+TdZQlj+VMgIm8ZeEIRNF0oaviF+QZcUZLWzWbYq6yD
'' SIG '' ok8EZKFiRR5otBoGLvaYFpxBZUE8mnLKuDlYobjrxh7l
'' SIG '' nwrxV/fMy0F9fSo2JxFmtLgtMIIHcTCCBVmgAwIBAgIT
'' SIG '' MwAAABXF52ueAptJmQAAAAAAFTANBgkqhkiG9w0BAQsF
'' SIG '' ADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
'' SIG '' bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
'' SIG '' FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMp
'' SIG '' TWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
'' SIG '' aXR5IDIwMTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMw
'' SIG '' MTgzMjI1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
'' SIG '' V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
'' SIG '' A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
'' SIG '' VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
'' SIG '' MDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
'' SIG '' AOThpkzntHIhC3miy9ckeb0O1YLT/e6cBwfSqWxOdcjK
'' SIG '' NVf2AX9sSuDivbk+F2Az/1xPx2b3lVNxWuJ+Slr+uDZn
'' SIG '' hUYjDLWNE893MsAQGOhgfWpSg0S3po5GawcU88V29YZQ
'' SIG '' 3MFEyHFcUTE3oAo4bo3t1w/YJlN8OWECesSq/XJprx2r
'' SIG '' rPY2vjUmZNqYO7oaezOtgFt+jBAcnVL+tuhiJdxqD89d
'' SIG '' 9P6OU8/W7IVWTe/dvI2k45GPsjksUZzpcGkNyjYtcI4x
'' SIG '' yDUoveO0hyTD4MmPfrVUj9z6BVWYbWg7mka97aSueik3
'' SIG '' rMvrg0XnRm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75x
'' SIG '' qRdbZ2De+JKRHh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9
'' SIG '' fvzZnkXftnIv231fgLrbqn427DZM9ituqBJR6L8FA6PR
'' SIG '' c6ZNN3SUHDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1D
'' SIG '' TsEzOUyOArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7XKHYC
'' SIG '' 4jMYctenIPDC+hIK12NvDMk2ZItboKaDIV1fMHSRlJTY
'' SIG '' uVD5C4lh8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqv
'' SIG '' UAV6bMURHXLvjflSxIUXk8A8FdsaN8cIFRg/eKtFtvUe
'' SIG '' h17aj54WcmnGrnu3tz5q4i6tAgMBAAGjggHdMIIB2TAS
'' SIG '' BgkrBgEEAYI3FQEEBQIDAQABMCMGCSsGAQQBgjcVAgQW
'' SIG '' BBQqp1L+ZMSavoKRPEY1Kc8Q/y8E7jAdBgNVHQ4EFgQU
'' SIG '' n6cVXQBeYl2D9OXSZacbUzUZ6XIwXAYDVR0gBFUwUzBR
'' SIG '' BgwrBgEEAYI3TIN9AQEwQTA/BggrBgEFBQcCARYzaHR0
'' SIG '' cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9Eb2Nz
'' SIG '' L1JlcG9zaXRvcnkuaHRtMBMGA1UdJQQMMAoGCCsGAQUF
'' SIG '' BwMIMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsG
'' SIG '' A1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1Ud
'' SIG '' IwQYMBaAFNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1Ud
'' SIG '' HwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0
'' SIG '' LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1
'' SIG '' dF8yMDEwLTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEww
'' SIG '' SgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0
'' SIG '' LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAt
'' SIG '' MDYtMjMuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCdVX38
'' SIG '' Kq3hLB9nATEkW+Geckv8qW/qXBS2Pk5HZHixBpOXPTEz
'' SIG '' tTnXwnE2P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7pZmc6
'' SIG '' U03dmLq2HnjYNi6cqYJWAAOwBb6J6Gngugnue99qb74p
'' SIG '' y27YP0h1AdkY3m2CDPVtI1TkeFN1JFe53Z/zjj3G82jf
'' SIG '' ZfakVqr3lbYoVSfQJL1AoL8ZthISEV09J+BAljis9/kp
'' SIG '' icO8F7BUhUKz/AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNh
'' SIG '' cy4sa3tuPywJeBTpkbKpW99Jo3QMvOyRgNI95ko+ZjtP
'' SIG '' u4b6MhrZlvSP9pEB9s7GdP32THJvEKt1MMU0sHrYUP4K
'' SIG '' WN1APMdUbZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMB
'' SIG '' V0lUZNlz138eW0QBjloZkWsNn6Qo3GcZKCS6OEuabvsh
'' SIG '' VGtqRRFHqfG3rsjoiV5PndLQTHa1V1QJsWkBRH58oWFs
'' SIG '' c/4Ku+xBZj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8
'' SIG '' vwLBgqJ7Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrpNPgk
'' SIG '' NWcr4A245oyZ1uEi6vAnQj0llOZ0dFtq0Z4+7X6gMTN9
'' SIG '' vMvpe784cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFx
'' SIG '' BmoQtB1VM1izoXBm8qGCAtQwggI9AgEBMIIBAKGB2KSB
'' SIG '' 1TCB0jELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
'' SIG '' bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
'' SIG '' FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMk
'' SIG '' TWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1p
'' SIG '' dGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjozQkQ0
'' SIG '' LTRCODAtNjlDMzElMCMGA1UEAxMcTWljcm9zb2Z0IFRp
'' SIG '' bWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUA
'' SIG '' 942iGuYFrsE4wzWDd85EpM6RiwqggYMwgYCkfjB8MQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
'' SIG '' MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3Nv
'' SIG '' ZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0B
'' SIG '' AQUFAAIFAOmBIr4wIhgPMjAyNDAyMjIwOTMyNDZaGA8y
'' SIG '' MDI0MDIyMzA5MzI0NlowdDA6BgorBgEEAYRZCgQBMSww
'' SIG '' KjAKAgUA6YEivgIBADAHAgEAAgICXTAHAgEAAgISeDAK
'' SIG '' AgUA6YJ0PgIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgor
'' SIG '' BgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYag
'' SIG '' MA0GCSqGSIb3DQEBBQUAA4GBAD4i+R9mmsZ72YM5XaL5
'' SIG '' o4Cd1c9slqXY2cY9nSzyIS4lpgiejLIQ8xCSXQRP6Jl8
'' SIG '' joJUDGibHpLzDOjA0bGWHccfSN7Zom5piIihb7dQuWsz
'' SIG '' zA+zQFaQ+ByMJm2T7GQHixkxwgook3yoEG2cqiXdd2GM
'' SIG '' h/8eDdrZak1PPT56kQ/SMYIEDTCCBAkCAQEwgZMwfDEL
'' SIG '' MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
'' SIG '' EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
'' SIG '' c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
'' SIG '' b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAHlj2rA
'' SIG '' 8z20C6MAAQAAAeUwDQYJYIZIAWUDBAIBBQCgggFKMBoG
'' SIG '' CSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG
'' SIG '' 9w0BCQQxIgQg+FYU7w9NKKQxKMa0XRanUR1TXiapplT1
'' SIG '' bm/Nbi3cwlwwgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHk
'' SIG '' MIG9BCAVqdP//qjxGFhe2YboEXeb8I/pAof01CwhbxUH
'' SIG '' 9U697TCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
'' SIG '' VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
'' SIG '' MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
'' SIG '' JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
'' SIG '' QSAyMDEwAhMzAAAB5Y9qwPM9tAujAAEAAAHlMCIEIHTj
'' SIG '' 8fXmZe8WsyVB0u5oyNzZBGr0+HZd3CtaLMQMo++rMA0G
'' SIG '' CSqGSIb3DQEBCwUABIICABlO3Y7A42IO6Cmn7ywPQ/yl
'' SIG '' xq0Z/IBFGMFxRwEnhxLF9NR9smeKeUNxFMRBHGsj0zBb
'' SIG '' 1cBpAqhO7D1g+oIeu8EQ19s7KaxE4EOiSkbM9BoGxgMe
'' SIG '' iBpaWC77Z6rUkuGnJ8uIlaNEjSnLUyOXOuQJrSclGBaQ
'' SIG '' /Tybt+ZgddXu05SmOP5rhcO8p85D1XOXXJhWsJ0cLaeX
'' SIG '' omesposU0qkfRD/jw2aDZFF1rnjfKWIZYCfmX+BdlRgo
'' SIG '' 7K3jSzslyRpBcSH7BEKU5KiEo+SuYpxDQd7sbJUE9i1A
'' SIG '' dXNCDXHyMrTI/qrQN6C7ZoanazW7+AylfW+0Uir0IKud
'' SIG '' DawhFjzG41N5wvwvoAPZlOwN8b490MxwgAwPVAr1ea1A
'' SIG '' nsyNVAxb15mW9vD0gehgrXWYmtsvhvF7qtw/b4YHYtgm
'' SIG '' YgO+L1nEn06P3Jo1o1dwyrDyW2OKJL9dQBSyl6OUxA5u
'' SIG '' rELForntyH9ANf3OohjRT4RYmmPzOKoQA2Ncf0tfsS5U
'' SIG '' G+jeY7QsBbEdqruwBxnPjCxB/sZNVu0YwJvYqdQU8You
'' SIG '' IcS1qWVzVa7JrOJluU682+wYwT9ThjWYtYFVW56Jzvpo
'' SIG '' lJFBEEZfDUenC9KQPRPjB+tJkqN970EMV8mE0hk2WwWU
'' SIG '' Z3LUQQi6Ammx93pFIjiKKtzKMj4DoOZxbtgWLXoBggoZ
'' SIG '' End signature block
