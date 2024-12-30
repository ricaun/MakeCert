' Windows Installer utility to report the language and codepage for a package
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the access of language and codepage values                 
'
Option Explicit

Const msiOpenDatabaseModeReadOnly     = 0
Const msiOpenDatabaseModeTransact     = 1
Const ForReading = 1
Const ForWriting = 2
Const TristateFalse = 0

Const msiViewModifyInsert         = 1
Const msiViewModifyUpdate         = 2
Const msiViewModifyAssign         = 3
Const msiViewModifyReplace        = 4
Const msiViewModifyDelete         = 6

Dim argCount:argCount = Wscript.Arguments.Count
If argCount > 0 Then If InStr(1, Wscript.Arguments(0), "?", vbTextCompare) > 0 Then argCount = 0
If (argCount = 0) Then
	message = "Windows Installer utility to manage language and codepage values for a package." &_
		vbNewLine & "The package language is a summary information property that designates the" &_
		vbNewLine & " primary language and any language transforms that are available, comma delim." &_
		vbNewLine & "The ProductLanguage in the database Property table is the language that is" &_
		vbNewLine & " registered for the product and determines the language used to load resources." &_
		vbNewLine & "The codepage is the ANSI codepage of the database strings, 0 if all ASCII data," &_
		vbNewLine & " and must represent the text data to avoid loss when persisting the database." &_
		vbNewLine & "The 1st argument is the path to MSI database (installer package)" &_
		vbNewLine & "To update a value, the 2nd argument contains the keyword and the 3rd the value:" &_
		vbNewLine & "   Package  {base LangId optionally followed by list of language transforms}" &_
		vbNewLine & "   Product  {LangId of the product (could be updated by language transforms)}" &_
		vbNewLine & "   Codepage {ANSI codepage of text data (use with caution when text exists!)}" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Echo message
	Wscript.Quit 1
End If

' Connect to Windows Installer object
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError


' Open database
Dim databasePath:databasePath = Wscript.Arguments(0)
Dim openMode : If argCount >= 3 Then openMode = msiOpenDatabaseModeTransact Else openMode = msiOpenDatabaseModeReadOnly
Dim database : Set database = installer.OpenDatabase(databasePath, openMode) : CheckError

' Update value if supplied
If argCount >= 3 Then
	Dim value:value = Wscript.Arguments(2)
	Select Case UCase(Wscript.Arguments(1))
		Case "PACKAGE"  : SetPackageLanguage database, value
		Case "PRODUCT"  : SetProductLanguage database, value
		Case "CODEPAGE" : SetDatabaseCodepage database, value
		Case Else       : Fail "Invalid value keyword"
	End Select
	CheckError
End If

' Extract language info and compose report message
Dim message:message = "Package language = "         & PackageLanguage(database) &_
					", ProductLanguage = " & ProductLanguage(database) &_
					", Database codepage = "        & DatabaseCodepage(database)
database.Commit : CheckError  ' no effect if opened ReadOnly
Set database = nothing
Wscript.Echo message
Wscript.Quit 0

' Get language list from summary information
Function PackageLanguage(database)
	On Error Resume Next
	Dim sumInfo  : Set sumInfo = database.SummaryInformation(0) : CheckError
	Dim template : template = sumInfo.Property(7) : CheckError
	Dim iDelim:iDelim = InStr(1, template, ";", vbTextCompare)
	If iDelim = 0 Then template = "Not specified!"
	PackageLanguage = Right(template, Len(template) - iDelim)
	If Len(PackageLanguage) = 0 Then PackageLanguage = "0"
End Function

' Get ProductLanguge property from Property table
Function ProductLanguage(database)
	On Error Resume Next
	Dim view : Set view = database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'ProductLanguage'")
	view.Execute : CheckError
	Dim record : Set record = view.Fetch : CheckError
	If record Is Nothing Then ProductLanguage = "Not specified!" Else ProductLanguage = record.IntegerData(1)
End Function

' Get ANSI codepage of database text data
Function DatabaseCodepage(database)
	On Error Resume Next
	Dim WshShell : Set WshShell = Wscript.CreateObject("Wscript.Shell") : CheckError
	Dim tempPath:tempPath = WshShell.ExpandEnvironmentStrings("%TEMP%") : CheckError
	database.Export "_ForceCodepage", tempPath, "codepage.idt" : CheckError
	Dim fileSys : Set fileSys = CreateObject("Scripting.FileSystemObject") : CheckError
	Dim file : Set file = fileSys.OpenTextFile(tempPath & "\codepage.idt", ForReading, False, TristateFalse) : CheckError
	file.ReadLine ' skip column name record
	file.ReadLine ' skip column defn record
	DatabaseCodepage = file.ReadLine
	file.Close
	Dim iDelim:iDelim = InStr(1, DatabaseCodepage, vbTab, vbTextCompare)
	If iDelim = 0 Then Fail "Failure in codepage export file"
	DatabaseCodepage = Left(DatabaseCodepage, iDelim - 1)
	fileSys.DeleteFile(tempPath & "\codepage.idt")
End Function

' Set ProductLanguge property in Property table
Sub SetProductLanguage(database, language)
	On Error Resume Next
	If Not IsNumeric(language) Then Fail "ProductLanguage must be numeric"
	Dim view : Set view = database.OpenView("SELECT `Property`,`Value` FROM `Property`")
	view.Execute : CheckError
	Dim record : Set record = installer.CreateRecord(2)
	record.StringData(1) = "ProductLanguage"
	record.StringData(2) = CStr(language)
	view.Modify msiViewModifyAssign, record : CheckError
End Sub

' Set ANSI codepage of database text data
Sub SetDatabaseCodepage(database, codepage)
	On Error Resume Next
	If Not IsNumeric(codepage) Then Fail "Codepage must be numeric"
	Dim WshShell : Set WshShell = Wscript.CreateObject("Wscript.Shell") : CheckError
	Dim tempPath:tempPath = WshShell.ExpandEnvironmentStrings("%TEMP%") : CheckError
	Dim fileSys : Set fileSys = CreateObject("Scripting.FileSystemObject") : CheckError
	Dim file : Set file = fileSys.OpenTextFile(tempPath & "\codepage.idt", ForWriting, True, TristateFalse) : CheckError
	file.WriteLine ' dummy column name record
	file.WriteLine ' dummy column defn record
	file.WriteLine codepage & vbTab & "_ForceCodepage"
	file.Close : CheckError
	database.Import tempPath, "codepage.idt" : CheckError
	fileSys.DeleteFile(tempPath & "\codepage.idt")
End Sub     

' Set language list in summary information
Sub SetPackageLanguage(database, language)
	On Error Resume Next
	Dim sumInfo  : Set sumInfo = database.SummaryInformation(1) : CheckError
	Dim template : template = sumInfo.Property(7) : CheckError
	Dim iDelim:iDelim = InStr(1, template, ";", vbTextCompare)
	Dim platform : If iDelim = 0 Then platform = ";" Else platform = Left(template, iDelim)
	sumInfo.Property(7) = platform & language
	sumInfo.Persist : CheckError
End Sub

Sub CheckError
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText
	End If
	Fail message
End Sub

Sub Fail(message)
	Wscript.Echo message
	Wscript.Quit 2
End Sub

'' SIG '' Begin signature block
'' SIG '' MIImEwYJKoZIhvcNAQcCoIImBDCCJgACAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' P5ZR+tRLXw+tvFB7cXDc0jFoO6HhZPDQciZh+dfNY5qg
'' SIG '' ggt2MIIE/jCCA+agAwIBAgITMwAABVbJICsfdDJdLQAA
'' SIG '' AAAFVjANBgkqhkiG9w0BAQsFADB+MQswCQYDVQQGEwJV
'' SIG '' UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
'' SIG '' UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
'' SIG '' aWduaW5nIFBDQSAyMDEwMB4XDTIzMTAxOTE5NTExMVoX
'' SIG '' DTI0MTAxNjE5NTExMVowdDELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEeMBwGA1UEAxMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
'' SIG '' MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
'' SIG '' ltpIPPc1p7LIgxvQBav7MapD0+N1eDGer8LuwuPrJcuO
'' SIG '' kCQOFDcUkZxg8/bvH9fDkdfwK/YLkA6kbYazjpLS2qJe
'' SIG '' PR2X7/JdQxHgf7oLlktKhSCXvnCum+4K1X5dEme1PMjl
'' SIG '' 7uu5+ds/kCTfolMXCJNClnLv7CWfCn3sCsZzQzAyBx4V
'' SIG '' B7yI0FobysTiwv08C9IuME8pF7kMG8CGbrhou02APNkN
'' SIG '' i5GDi5cDkzzm9HqMIXFCOwml5VN9CIKBuH62PprWTGZ0
'' SIG '' 8dIGv2t+hlTXaujXgSs5RmywdNv1iD/nOQAwwl7IXlqZ
'' SIG '' IsybfWj4c2LqJ7fjcdDoSB9OJSRbwqo5YwIDAQABo4IB
'' SIG '' fTCCAXkwHwYDVR0lBBgwFgYKKwYBBAGCNz0GAQYIKwYB
'' SIG '' BQUHAwMwHQYDVR0OBBYEFCbfBYUBcF+4OQP9HpQ8ZI8M
'' SIG '' PNnaMFQGA1UdEQRNMEukSTBHMS0wKwYDVQQLEyRNaWNy
'' SIG '' b3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQx
'' SIG '' FjAUBgNVBAUTDTIzMDg2NSs1MDE2NTUwHwYDVR0jBBgw
'' SIG '' FoAU5vxfe7siAFjkck619CF0IzLm76wwVgYDVR0fBE8w
'' SIG '' TTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29t
'' SIG '' L3BraS9jcmwvcHJvZHVjdHMvTWljQ29kU2lnUENBXzIw
'' SIG '' MTAtMDctMDYuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggr
'' SIG '' BgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29t
'' SIG '' L3BraS9jZXJ0cy9NaWNDb2RTaWdQQ0FfMjAxMC0wNy0w
'' SIG '' Ni5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsF
'' SIG '' AAOCAQEAQp2ZaDMYxwVRyRD+nftLexAyXzQdIe4/Yjl+
'' SIG '' i0IjzHUAFdcagOiYG/1RD0hFbNO+ggCZ9yj+Saa+Azrq
'' SIG '' NdgRNgqArrGQx5/u2j9ssZ4DBhkHCSs+FHzswzEvWK9r
'' SIG '' Jd0enzD9fE+AnubeyGBSt+jyPx37xzvAMwd09CoVSIn6
'' SIG '' rEsGfJhLpMP8IuHbiWLpWMVdpWNpDB8L/zirygLK03d9
'' SIG '' /B5Z7kfs/TWb0rTVItWvLE8HBDKxD/JYLaMWmXtGKbvz
'' SIG '' oZ+D6k3nxFVikCS1Nihciw5KGpg3XtMnQM8x2BKnQUDF
'' SIG '' tIMVsryfX44BfwtjykFbv9EjAYXMKNOHhc3/8O6WfzCC
'' SIG '' BnAwggRYoAMCAQICCmEMUkwAAAAAAAMwDQYJKoZIhvcN
'' SIG '' AQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
'' SIG '' YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
'' SIG '' VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNV
'' SIG '' BAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1
'' SIG '' dGhvcml0eSAyMDEwMB4XDTEwMDcwNjIwNDAxN1oXDTI1
'' SIG '' MDcwNjIwNTAxN1owfjELMAkGA1UEBhMCVVMxEzARBgNV
'' SIG '' BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
'' SIG '' HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEo
'' SIG '' MCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQ
'' SIG '' Q0EgMjAxMDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCC
'' SIG '' AQoCggEBAOkOZFB5Z7XE4/0JAEyelKz3VmjqRNjPxVhP
'' SIG '' qaV2fG1FutM5krSkHvn5ZYLkF9KP/UScCOhlk84sVYS/
'' SIG '' fQjjLiuoQSsYt6JLbklMaxUH3tHSwokecZTNtX9LtK8I
'' SIG '' 2MyI1msXlDqTziY/7Ob+NJhX1R1dSfayKi7VhbtZP/iQ
'' SIG '' tCuDdMorsztG4/BGScEXZlTJHL0dxFViV3L4Z7klIDTe
'' SIG '' XaallV6rKIDN1bKe5QO1Y9OyFMjByIomCll/B+z/Du2A
'' SIG '' EjVMEqa+Ulv1ptrgiwtId9aFR9UQucboqu6Lai0FXGDG
'' SIG '' tCpbnCMcX0XjGhQebzfLGTOAaolNo2pmY3iT1TDPlR8C
'' SIG '' AwEAAaOCAeMwggHfMBAGCSsGAQQBgjcVAQQDAgEAMB0G
'' SIG '' A1UdDgQWBBTm/F97uyIAWORyTrX0IXQjMubvrDAZBgkr
'' SIG '' BgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
'' SIG '' AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV
'' SIG '' 9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEug
'' SIG '' SaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
'' SIG '' L2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0w
'' SIG '' Ni0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUF
'' SIG '' BzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
'' SIG '' L2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNy
'' SIG '' dDCBnQYDVR0gBIGVMIGSMIGPBgkrBgEEAYI3LgMwgYEw
'' SIG '' PQYIKwYBBQUHAgEWMWh0dHA6Ly93d3cubWljcm9zb2Z0
'' SIG '' LmNvbS9QS0kvZG9jcy9DUFMvZGVmYXVsdC5odG0wQAYI
'' SIG '' KwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AUABvAGwA
'' SIG '' aQBjAHkAXwBTAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJ
'' SIG '' KoZIhvcNAQELBQADggIBABp071dPKXvEFoV4uFDTIvwJ
'' SIG '' nayCl/g0/yosl5US5eS/z7+TyOM0qduBuNweAL7SNW+v
'' SIG '' 5X95lXflAtTx69jNTh4bYaLCWiMa8IyoYlFFZwjjPzwe
'' SIG '' k/gwhRfIOUCm1w6zISnlpaFpjCKTzHSY56FHQ/JTrMAP
'' SIG '' MGl//tIlIG1vYdPfB9XZcgAsaYZ2PVHbpjlIyTdhbQfd
'' SIG '' UxnLp9Zhwr/ig6sP4GubldZ9KFGwiUpRpJpsyLcfShoO
'' SIG '' aanX3MF+0Ulwqratu3JHYxf6ptaipobsqBBEm2O2smmJ
'' SIG '' BsdGhnoYP+jFHSHVe/kCIy3FQcu/HUzIFu+xnH/8IktJ
'' SIG '' im4V46Z/dlvRU3mRhZ3V0ts9czXzPK5UslJHasCqE5XS
'' SIG '' jhHamWdeMoz7N4XR3HWFnIfGWleFwr/dDY+Mmy3rtO7P
'' SIG '' J9O1Xmn6pBYEAackZ3PPTU+23gVWl3r36VJN9HcFT4XG
'' SIG '' 2Avxju1CCdENduMjVngiJja+yrGMbqod5IXaRzNij6TJ
'' SIG '' kTNfcR5Ar5hlySLoQiElihwtYNk3iUGJKhYP12E8lGhg
'' SIG '' Uu/WR5mggEDuFYF3PpzgUxgaUB04lZseZjMTJzkXeIc2
'' SIG '' zk7DX7L1PUdTtuDl2wthPSrXkizON1o+QEIxpB8QCMJW
'' SIG '' nL8kXVECnWp50hfT2sGUjgd7JXFEqwZq5tTG3yOalnXF
'' SIG '' MYIZ9TCCGfECAQEwgZUwfjELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmlu
'' SIG '' ZyBQQ0EgMjAxMAITMwAABVbJICsfdDJdLQAAAAAFVjAN
'' SIG '' BglghkgBZQMEAgEFAKCCAQQwGQYJKoZIhvcNAQkDMQwG
'' SIG '' CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisG
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIACfHEUjpbb6
'' SIG '' OMsQM89w72uZFdkYNMUaoednrszUaXmzMDwGCisGAQQB
'' SIG '' gjcKAxwxLgwsc1BZN3hQQjdoVDVnNUhIcll0OHJETFNN
'' SIG '' OVZ1WlJ1V1phZWYyZTIyUnM1ND0wWgYKKwYBBAGCNwIB
'' SIG '' DDFMMEqgJIAiAE0AaQBjAHIAbwBzAG8AZgB0ACAAVwBp
'' SIG '' AG4AZABvAHcAc6EigCBodHRwOi8vd3d3Lm1pY3Jvc29m
'' SIG '' dC5jb20vd2luZG93czANBgkqhkiG9w0BAQEFAASCAQAk
'' SIG '' XwK2y+ZTcr/TcWyW7FX/eBX0wwgwa7M2eswW47IxgnJg
'' SIG '' szgq4dUfWUZSVQHt37nbkr3U0zi4NeDJdAxXRbQAaOiw
'' SIG '' ROmE6NG6JsFje1MM+ExgfWhs2y8eRFuaUm+mxI4HfBU2
'' SIG '' 6oFVu4rrftkstZw9Usw/oHa1jVK2LWTAhalY1+f7Ix8Z
'' SIG '' sMadbZvukPk4508bFis5vtJWoR6enSEyjaHqFSFWYV/w
'' SIG '' E9+NmxwIYZ2RhU7CHxC9fRSFAefHz8cLKGJwTbS9Gpn2
'' SIG '' ymZlL6qZ0pxchqXUocfB3N/ztuagt1SyfjC8jyccjNoU
'' SIG '' 2ciS9PCqxNg36HHkGaL2g+0EoIg9d7jCoYIXKDCCFyQG
'' SIG '' CisGAQQBgjcDAwExghcUMIIXEAYJKoZIhvcNAQcCoIIX
'' SIG '' ATCCFv0CAQMxDzANBglghkgBZQMEAgEFADCCAVgGCyqG
'' SIG '' SIb3DQEJEAEEoIIBRwSCAUMwggE/AgEBBgorBgEEAYRZ
'' SIG '' CgMBMDEwDQYJYIZIAWUDBAIBBQAEIFop5+dg7tS3hU1i
'' SIG '' mU/ZkhsqVu/Zxoovqj5YZYbEtxqNAgZl1eEvhBcYEjIw
'' SIG '' MjQwMjIyMTA0NDI3LjUxWjAEgAIB9KCB2KSB1TCB0jEL
'' SIG '' MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
'' SIG '' EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
'' SIG '' c29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9z
'' SIG '' b2Z0IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMSYw
'' SIG '' JAYDVQQLEx1UaGFsZXMgVFNTIEVTTjoyQUQ0LTRCOTIt
'' SIG '' RkEwMTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3Rh
'' SIG '' bXAgU2VydmljZaCCEXgwggcnMIIFD6ADAgECAhMzAAAB
'' SIG '' 3p5InpafKEQ9AAEAAAHeMA0GCSqGSIb3DQEBCwUAMHwx
'' SIG '' CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
'' SIG '' MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
'' SIG '' c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTIzMTAx
'' SIG '' MjE5MDcxMloXDTI1MDExMDE5MDcxMlowgdIxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
'' SIG '' cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UE
'' SIG '' CxMdVGhhbGVzIFRTUyBFU046MkFENC00QjkyLUZBMDEx
'' SIG '' JTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
'' SIG '' cnZpY2UwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
'' SIG '' AoICAQC0gfQchfVCA4QOsRazp4sP8bA5fLEovazgjl0k
'' SIG '' juFTEI5zRgKOVR8dIoozBDB/S2NklCAZFUEtDJepEfk2
'' SIG '' oJFD22hKcI4UNZqa4UYCU/45Up4nONlQwKNHp+CSOsZ1
'' SIG '' 6AKFqCskmPP0TiCnaaYYCOziW+Fx5NT97F9qTWd9iw2N
'' SIG '' ZLXIStf4Vsj5W5WlwB0btBN8p78K0vP23KKwDTug47sr
'' SIG '' Mkvc1Jq/sNx9wBL0oLNkXri49qZAXH1tVDwhbnS3eyD2
'' SIG '' dkQuKHUHBD52Ndo8qWD50usmQLNKS6atCkRVMgdcesej
'' SIG '' lO97LnYhzjdephNJeiy0/TphqNEveAcYNzf92hOn1G51
'' SIG '' aHplXOxZBS7pvCpGXG0O3Dh0gFhicXQr6OTrVLUXUqn/
'' SIG '' ORZJQlyCJIOLJu5zPU5LVFXztJKepMe5srIA9EK8cev+
'' SIG '' aGqp8Dk1izcyvgQotRu51A9abXrl70KfHxNSqU45xv9T
'' SIG '' iXnocCjTT4xrffFdAZqIGU3t0sQZDnjkMiwPvuR8oPy+
'' SIG '' vKXvg62aGT1yWhlP4gYhZi/rpfzot3fN8ywB5R0Jh/1R
'' SIG '' jQX0cD/osb6ocpPxHm8Ll1SWPq08n20X7ofZ9AGjIYTc
'' SIG '' cYOrRismUuBABIg8axfZgGRMvHvK3+nZSiF+Xd2kC6PX
'' SIG '' w3WtWUzsPlwHAL49vzdwy1RmZR5x5QIDAQABo4IBSTCC
'' SIG '' AUUwHQYDVR0OBBYEFGswJm8bHmmqYHccyvDrPp2j0BLI
'' SIG '' MB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1Gely
'' SIG '' MF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWlj
'' SIG '' cm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
'' SIG '' MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBs
'' SIG '' BggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6
'' SIG '' Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMv
'' SIG '' TWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIw
'' SIG '' MTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/
'' SIG '' BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQDAgeAMA0G
'' SIG '' CSqGSIb3DQEBCwUAA4ICAQDilMB7Fw2nBjr1CILORw4D
'' SIG '' 7NC2dash0ugusHypS2g9+rWX21rdcfhjIms0rsvhrMYl
'' SIG '' R85ITFvhaivIK7i0Fjf7Dgl/nxlIE/S09tXESKXGY+P2
'' SIG '' RSL8LZAXLAs9VxFLF2DkiVD4rWOxPG25XZpoWGdvafl0
'' SIG '' KSHLBv6vmI5KgVvZsNK7tTH8TE0LPTEw4g9vIAFRqzwN
'' SIG '' zcpIkgob3aku1V/vy3BM/VG87aP8NvFgPBzgh6gU2w0R
'' SIG '' 5oj+zCI/kkJiPVSGsmLCBkY73pZjWtDr21PQiUs/zXzB
'' SIG '' IH9jRzGVGFvCqlhIyIz3xyCsVpTTGIbln1kUh2QisiAD
'' SIG '' QNGiS+LKB0Lc82djJzX42GPOdcB2IxoMFI/4ZS0YEDuU
'' SIG '' t9Gce/BqgSn8paduWjlif6j4Qvg1zNoF2oyF25fo6RnF
'' SIG '' QDcLRRbowiUXWW3h9UfkONRY4AYOJtzkxQxqLeQ0rlZE
'' SIG '' II5Lu6TlT7ZXROOkJQ4P9loT6U0MVx+uLD9Rn5AMFLbe
'' SIG '' q62TPzwsERuoIq2Jp00Sy7InAYaGC4fhBBY1b4lwBk5O
'' SIG '' qZ7vI8f+Fj1rtI7M+8hc4PNvxTKgpPcCty78iwMgxzfh
'' SIG '' cWxwMbYMGne6C0DzNFhhEXQdbpjwiImLEn/4+/RKh3aD
'' SIG '' cEGETlZvmV9dEV95+m0ZgJ7JHjYYtMJ1WnlaICzHRg/p
'' SIG '' 6jCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAA
'' SIG '' ABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVT
'' SIG '' MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
'' SIG '' ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
'' SIG '' YXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENl
'' SIG '' cnRpZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkz
'' SIG '' MDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
'' SIG '' bWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEB
'' SIG '' AQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9
'' SIG '' DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25PhdgM/9c
'' SIG '' T8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1q
'' SIG '' UoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP
'' SIG '' 2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6GnszrYBb
'' SIG '' fowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOOR
'' SIG '' j7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c
'' SIG '' +gVVmG1oO5pGve2krnopN6zL64NF50ZuyjLVwIYwXE8s
'' SIG '' 4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9
'' SIG '' QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+
'' SIG '' Nuw2TPYrbqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjt
'' SIG '' p+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqEUUbi
'' SIG '' 0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJ
'' SIG '' NmSLW6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR
'' SIG '' /bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435UsSFF5PA
'' SIG '' PBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIu
'' SIG '' rQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEA
'' SIG '' ATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxGNSnP
'' SIG '' EP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1
'' SIG '' GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEw
'' SIG '' PwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0
'' SIG '' LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTAT
'' SIG '' BgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIE
'' SIG '' DB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0T
'' SIG '' AQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo
'' SIG '' 0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRw
'' SIG '' Oi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9k
'' SIG '' dWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmww
'' SIG '' WgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRw
'' SIG '' Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01p
'' SIG '' Y1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG
'' SIG '' 9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv
'' SIG '' 6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsP
'' SIG '' MeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYunKmCVgAD
'' SIG '' sAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU
'' SIG '' 5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/
'' SIG '' GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVCs/wMnosZiefw
'' SIG '' C2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvf
'' SIG '' SaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT9
'' SIG '' 9kxybxCrdTDFNLB62FD+CljdQDzHVG2dY3RILLFORy3B
'' SIG '' FARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFr
'' SIG '' DZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S
'' SIG '' 0Ex2tVdUCbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6T
'' SIG '' vsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7ntdAo
'' SIG '' GokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9
'' SIG '' JZTmdHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoI
'' SIG '' Yn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKhggLU
'' SIG '' MIICPQIBATCCAQChgdikgdUwgdIxCzAJBgNVBAYTAlVT
'' SIG '' MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
'' SIG '' ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
'' SIG '' YXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5k
'' SIG '' IE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhh
'' SIG '' bGVzIFRTUyBFU046MkFENC00QjkyLUZBMDExJTAjBgNV
'' SIG '' BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wi
'' SIG '' IwoBATAHBgUrDgMCGgMVAGigUorMuMvOqZfF8ttgiWRM
'' SIG '' RNrzoIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNV
'' SIG '' BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
'' SIG '' HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEm
'' SIG '' MCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
'' SIG '' IDIwMTAwDQYJKoZIhvcNAQEFBQACBQDpgQhoMCIYDzIw
'' SIG '' MjQwMjIyMDc0MDI0WhgPMjAyNDAyMjMwNzQwMjRaMHQw
'' SIG '' OgYKKwYBBAGEWQoEATEsMCowCgIFAOmBCGgCAQAwBwIB
'' SIG '' AAICFN8wBwIBAAICEUQwCgIFAOmCWegCAQAwNgYKKwYB
'' SIG '' BAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQAC
'' SIG '' AwehIKEKMAgCAQACAwGGoDANBgkqhkiG9w0BAQUFAAOB
'' SIG '' gQAjz16H48Dhbws9LukCU/xXm5lt7RD6kaqzAWOEheLQ
'' SIG '' J5BADgjJQxHaQNGwt27FAa/QeTevWwpvgqyQj2cjmVe/
'' SIG '' sdOipvE8iPXYvl1lZ476P/b9//mjUSSl1p4Lrd67pyCe
'' SIG '' RgBtyS20/k9BjrgktVfWTpiWTlTGdHNPkSb3KDFsNjGC
'' SIG '' BA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYD
'' SIG '' VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
'' SIG '' MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
'' SIG '' JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
'' SIG '' QSAyMDEwAhMzAAAB3p5InpafKEQ9AAEAAAHeMA0GCWCG
'' SIG '' SAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZI
'' SIG '' hvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIJ0mKHFqRSsT
'' SIG '' nK6YgMJhckOOuFPNMcRxMUi0G59UvqtHMIH6BgsqhkiG
'' SIG '' 9w0BCRACLzGB6jCB5zCB5DCBvQQgjj4jnw3BXhAQSQJ/
'' SIG '' 5gtzIK0+cP1Ns/NS2A+OB3N+HXswgZgwgYCkfjB8MQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
'' SIG '' MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3Nv
'' SIG '' ZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAd6eSJ6W
'' SIG '' nyhEPQABAAAB3jAiBCDK+U+C7VDu3sCeZj3Z1XlPP2wD
'' SIG '' uuO/HK6jFNJasRWJAjANBgkqhkiG9w0BAQsFAASCAgBE
'' SIG '' /LcDGIMQ/1MPNF+ZQynHJiP1QldszF9dDGt9mZePpjBW
'' SIG '' 4VZBEuQcEY9bxA2X/peGUYrq4BSVffUJyn/myYsvpfrq
'' SIG '' IyJDFY4+Lv4absf4qlse26eebodCr3f6bJ9zsWCEajzY
'' SIG '' 8sG0KMduThk3+3RVmwCD9R+rKSgrf8DXnvYlXyLt8cJp
'' SIG '' qkv8vLPJ7YC+8jdd6a8tUpZfYKF+FI5+QmIv9zAXnXvD
'' SIG '' 6/LQ7BQNWbL7K4RLOvfc8mZJKpyEHsmMHqz+AROedQ6m
'' SIG '' LmvvzD7abyNtkGPuKUWBk+arPTm0QisD575R/Q5CP/fV
'' SIG '' CYYN97UGpsY60KxXvCkRn5+60FT3ALJn52/82egCWyA0
'' SIG '' DOYcBt+4TJiAcHdld4qdFf3nuE5SGPAujZPNA2AfPTyC
'' SIG '' 3YEitqfFfgzhDIDaSMZdvI/i9I4eL/asNYbobCoRT8OZ
'' SIG '' SvbBZFJt+BlfEtaQPkzLeGcCTUBLQaAV78N3vN3IMoPc
'' SIG '' ifqQp8jMUp6wcMLTi9Y0XtFn2c/5/KX+DCZw0rCcgQg+
'' SIG '' 2nU0+hC+y9ssQSAfpeOCtNiRdoHxolQpH616R39iKo9K
'' SIG '' 27zUutYc8FFtiRvAeyeJIIMgCdnnWN5fRUdD59nZxdBQ
'' SIG '' YV6Q2iYbH/lrdBx+IYtoNWRsNNJXuL6JjPgLAmGEHSQ7
'' SIG '' yX6NYPewl8qKPQP6i6GgKw==
'' SIG '' End signature block
