' Windows Installer database table import for use with Windows Scripting Host
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of the Database.Import method and MsiDatabaseImport API
'
Option Explicit

Const msiOpenDatabaseModeReadOnly     = 0
Const msiOpenDatabaseModeTransact     = 1
Const msiOpenDatabaseModeCreate       = 3
Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
Const TristateTrue = -1

Dim argCount:argCount = Wscript.Arguments.Count
Dim iArg:iArg = 0
If (argCount < 3) Then
	Wscript.Echo "Windows Installer database table import utility" &_
		vbNewLine & " 1st argument is the path to MSI database (installer package)" &_
		vbNewLine & " 2nd argument is the path to folder containing the imported files" &_
		vbNewLine & " Subseqent arguments are names of archive files to import" &_
		vbNewLine & " Wildcards, such as *.idt, can be used to import multiple files" &_
		vbNewLine & " Specify /c or -c anywhere before file list to create new database" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Connect to Windows Installer object
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError

Dim openMode:openMode = msiOpenDatabaseModeTransact
Dim databasePath:databasePath = NextArgument
Dim folder:folder = NextArgument

Dim WshShell, fileSys
Set WshShell = Wscript.CreateObject("Wscript.Shell") : CheckError
Set fileSys = CreateObject("Scripting.FileSystemObject") : CheckError

' Open database and process list of files
Dim database, table
Set database = installer.OpenDatabase(databasePath, openMode) : CheckError
While iArg < argCount
	table = NextArgument
	' Check file name for wildcard specification
	If (InStr(1,table,"*",vbTextCompare) <> 0) Or (InStr(1,table,"?",vbTextCompare) <> 0) Then
		' Obtain list of files matching wildcard specification
		Dim file, tempFilePath
		tempFilePath = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\dir.tmp"
		WshShell.Run "cmd.exe /U /c dir /b " & folder & "\" & table & ">" & tempFilePath, 0, True : CheckError
		Set file = fileSys.OpenTextFile(tempFilePath, ForReading, False, TristateTrue) : CheckError
		' Import each file in directory list
		Do While file.AtEndOfStream <> True
			table = file.ReadLine
			database.Import folder, table : CheckError
		Loop
		file.Close
		fileSys.DeleteFile(tempFilePath)
	Else
		database.Import folder, table : CheckError
	End If
Wend
database.Commit 'commit changes if no import errors
Wscript.Quit 0

Function NextArgument
	Dim arg, chFlag
	Do
		arg = Wscript.Arguments(iArg)
		iArg = iArg + 1
		chFlag = AscW(arg)
		If (chFlag = AscW("/")) Or (chFlag = AscW("-")) Then
			chFlag = UCase(Right(arg, Len(arg)-1))
			If chFlag = "C" Then 
				openMode = msiOpenDatabaseModeCreate
			Else
				Wscript.Echo "Invalid option flag:", arg : Wscript.Quit 1
			End If
		Else
			Exit Do
		End If
	Loop
	NextArgument = arg
End Function

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
'' SIG '' MIImFAYJKoZIhvcNAQcCoIImBTCCJgECAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' URtattTTdos4rg+Jt96T7zPQ8Pzvx4qAdD0rt5bwgTmg
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
'' SIG '' MYIZ9jCCGfICAQEwgZUwfjELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmlu
'' SIG '' ZyBQQ0EgMjAxMAITMwAABVbJICsfdDJdLQAAAAAFVjAN
'' SIG '' BglghkgBZQMEAgEFAKCCAQQwGQYJKoZIhvcNAQkDMQwG
'' SIG '' CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisG
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIO3llHxswHP+
'' SIG '' QO84qZFS5dkzT2yJc3rEKv73TS6ueMCwMDwGCisGAQQB
'' SIG '' gjcKAxwxLgwsc1BZN3hQQjdoVDVnNUhIcll0OHJETFNN
'' SIG '' OVZ1WlJ1V1phZWYyZTIyUnM1ND0wWgYKKwYBBAGCNwIB
'' SIG '' DDFMMEqgJIAiAE0AaQBjAHIAbwBzAG8AZgB0ACAAVwBp
'' SIG '' AG4AZABvAHcAc6EigCBodHRwOi8vd3d3Lm1pY3Jvc29m
'' SIG '' dC5jb20vd2luZG93czANBgkqhkiG9w0BAQEFAASCAQAH
'' SIG '' fsEYAtqFAPvINruJDMC3eXTXEM7PeQLjmrYNvix4d+07
'' SIG '' ZV9QpLGJtcmod2MyKwOoOGOiWhvsqvPTnVTkXfZ/lhqR
'' SIG '' tluJ5u5ClNb92A+EM7r1mySTe6Y7Npo1Bq7Fa/RXmgQT
'' SIG '' xS3GtezD7+SwbjBDKuKcH8BXLDivOGhFa6B9s90eZBct
'' SIG '' 49yq3O3oIYjXU0vrAGsBLRcuuNtiUBNxEd8jqA1gkORL
'' SIG '' m/gtY10/cHvQBqR9vPPJ2SeV2X1v2oDbXwd6MBlpNbtg
'' SIG '' 7MUHckSE+zya3oTU72R7wXYzFCXN3AZp/9KqGREnrlOo
'' SIG '' /lq3S0sKO49Y+QnnUKtTEwkBMJPUqBeioYIXKTCCFyUG
'' SIG '' CisGAQQBgjcDAwExghcVMIIXEQYJKoZIhvcNAQcCoIIX
'' SIG '' AjCCFv4CAQMxDzANBglghkgBZQMEAgEFADCCAVkGCyqG
'' SIG '' SIb3DQEJEAEEoIIBSASCAUQwggFAAgEBBgorBgEEAYRZ
'' SIG '' CgMBMDEwDQYJYIZIAWUDBAIBBQAEIIQHB/P+oVjwj1M7
'' SIG '' 5D7O/JYBVvZb5jFurETmu5pACDXmAgZl1eEvg+gYEzIw
'' SIG '' MjQwMjIyMTA0NDIzLjQzNlowBIACAfSggdikgdUwgdIx
'' SIG '' CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
'' SIG '' MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jv
'' SIG '' c29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEm
'' SIG '' MCQGA1UECxMdVGhhbGVzIFRTUyBFU046MkFENC00Qjky
'' SIG '' LUZBMDExJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0
'' SIG '' YW1wIFNlcnZpY2WgghF4MIIHJzCCBQ+gAwIBAgITMwAA
'' SIG '' Ad6eSJ6WnyhEPQABAAAB3jANBgkqhkiG9w0BAQsFADB8
'' SIG '' MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
'' SIG '' bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
'' SIG '' cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
'' SIG '' b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzEw
'' SIG '' MTIxOTA3MTJaFw0yNTAxMTAxOTA3MTJaMIHSMQswCQYD
'' SIG '' VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
'' SIG '' A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
'' SIG '' IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQg
'' SIG '' SXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNV
'' SIG '' BAsTHVRoYWxlcyBUU1MgRVNOOjJBRDQtNEI5Mi1GQTAx
'' SIG '' MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
'' SIG '' ZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
'' SIG '' CgKCAgEAtIH0HIX1QgOEDrEWs6eLD/GwOXyxKL2s4I5d
'' SIG '' JI7hUxCOc0YCjlUfHSKKMwQwf0tjZJQgGRVBLQyXqRH5
'' SIG '' NqCRQ9toSnCOFDWamuFGAlP+OVKeJzjZUMCjR6fgkjrG
'' SIG '' degChagrJJjz9E4gp2mmGAjs4lvhceTU/exfak1nfYsN
'' SIG '' jWS1yErX+FbI+VuVpcAdG7QTfKe/CtLz9tyisA07oOO7
'' SIG '' KzJL3NSav7DcfcAS9KCzZF64uPamQFx9bVQ8IW50t3sg
'' SIG '' 9nZELih1BwQ+djXaPKlg+dLrJkCzSkumrQpEVTIHXHrH
'' SIG '' o5Tvey52Ic43XqYTSXostP06YajRL3gHGDc3/doTp9Ru
'' SIG '' dWh6ZVzsWQUu6bwqRlxtDtw4dIBYYnF0K+jk61S1F1Kp
'' SIG '' /zkWSUJcgiSDiybucz1OS1RV87SSnqTHubKyAPRCvHHr
'' SIG '' /mhqqfA5NYs3Mr4EKLUbudQPWm165e9Cnx8TUqlOOcb/
'' SIG '' U4l56HAo00+Ma33xXQGaiBlN7dLEGQ545DIsD77kfKD8
'' SIG '' vryl74Otmhk9cloZT+IGIWYv66X86Ld3zfMsAeUdCYf9
'' SIG '' UY0F9HA/6LG+qHKT8R5vC5dUlj6tPJ9tF+6H2fQBoyGE
'' SIG '' 3HGDq0YrJlLgQASIPGsX2YBkTLx7yt/p2Uohfl3dpAuj
'' SIG '' 18N1rVlM7D5cBwC+Pb83cMtUZmUeceUCAwEAAaOCAUkw
'' SIG '' ggFFMB0GA1UdDgQWBBRrMCZvGx5pqmB3HMrw6z6do9AS
'' SIG '' yDAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnp
'' SIG '' cjBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1p
'' SIG '' Y3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQl
'' SIG '' MjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmww
'' SIG '' bAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRw
'' SIG '' Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRz
'' SIG '' L01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAy
'' SIG '' MDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB
'' SIG '' /wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDAN
'' SIG '' BgkqhkiG9w0BAQsFAAOCAgEA4pTAexcNpwY69QiCzkcO
'' SIG '' A+zQtnWrIdLoLrB8qUtoPfq1l9ta3XH4YyJrNK7L4azG
'' SIG '' JUfOSExb4WoryCu4tBY3+w4Jf58ZSBP0tPbVxEilxmPj
'' SIG '' 9kUi/C2QFywLPVcRSxdg5IlQ+K1jsTxtuV2aaFhnb2n5
'' SIG '' dCkhywb+r5iOSoFb2bDSu7Ux/ExNCz0xMOIPbyABUas8
'' SIG '' Dc3KSJIKG92pLtVf78twTP1RvO2j/DbxYDwc4IeoFNsN
'' SIG '' EeaI/swiP5JCYj1UhrJiwgZGO96WY1rQ69tT0IlLP818
'' SIG '' wSB/Y0cxlRhbwqpYSMiM98cgrFaU0xiG5Z9ZFIdkIrIg
'' SIG '' A0DRokviygdC3PNnYyc1+NhjznXAdiMaDBSP+GUtGBA7
'' SIG '' lLfRnHvwaoEp/KWnblo5Yn+o+EL4NczaBdqMhduX6OkZ
'' SIG '' xUA3C0UW6MIlF1lt4fVH5DjUWOAGDibc5MUMai3kNK5W
'' SIG '' RCCOS7uk5U+2V0TjpCUOD/ZaE+lNDFcfriw/UZ+QDBS2
'' SIG '' 3qutkz88LBEbqCKtiadNEsuyJwGGhguH4QQWNW+JcAZO
'' SIG '' Tqme7yPH/hY9a7SOzPvIXODzb8UyoKT3Arcu/IsDIMc3
'' SIG '' 4XFscDG2DBp3ugtA8zRYYRF0HW6Y8IiJixJ/+Pv0Sod2
'' SIG '' g3BBhE5Wb5lfXRFfefptGYCeyR42GLTCdVp5WiAsx0YP
'' SIG '' 6eowggdxMIIFWaADAgECAhMzAAAAFcXna54Cm0mZAAAA
'' SIG '' AAAVMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJV
'' SIG '' UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
'' SIG '' UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBD
'' SIG '' ZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0yMTA5
'' SIG '' MzAxODIyMjVaFw0zMDA5MzAxODMyMjVaMHwxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBU
'' SIG '' aW1lLVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0B
'' SIG '' AQEFAAOCAg8AMIICCgKCAgEA5OGmTOe0ciELeaLL1yR5
'' SIG '' vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK4OK9uT4XYDP/
'' SIG '' XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY6GB9
'' SIG '' alKDRLemjkZrBxTzxXb1hlDcwUTIcVxRMTegCjhuje3X
'' SIG '' D9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk2pg7uhp7M62A
'' SIG '' W36MEBydUv626GIl3GoPz130/o5Tz9bshVZN7928jaTj
'' SIG '' kY+yOSxRnOlwaQ3KNi1wjjHINSi947SHJMPgyY9+tVSP
'' SIG '' 3PoFVZhtaDuaRr3tpK56KTesy+uDRedGbsoy1cCGMFxP
'' SIG '' LOJiss254o2I5JasAUq7vnGpF1tnYN74kpEeHT39IM9z
'' SIG '' fUGaRnXNxF803RKJ1v2lIH1+/NmeRd+2ci/bfV+Autuq
'' SIG '' fjbsNkz2K26oElHovwUDo9Fzpk03dJQcNIIP8BDyt0cY
'' SIG '' 7afomXw/TNuvXsLz1dhzPUNOwTM5TI4CvEJoLhDqhFFG
'' SIG '' 4tG9ahhaYQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28M
'' SIG '' yTZki1ugpoMhXV8wdJGUlNi5UPkLiWHzNgY1GIRH29wb
'' SIG '' 0f2y1BzFa/ZcUlFdEtsluq9QBXpsxREdcu+N+VLEhReT
'' SIG '' wDwV2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3Pmri
'' SIG '' Lq0CAwEAAaOCAd0wggHZMBIGCSsGAQQBgjcVAQQFAgMB
'' SIG '' AAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5kxJq+gpE8RjUp
'' SIG '' zxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJlpxtT
'' SIG '' NRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQBgjdMg30BATBB
'' SIG '' MD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29m
'' SIG '' dC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0w
'' SIG '' EwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJKwYBBAGCNxQC
'' SIG '' BAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1Ud
'' SIG '' EwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxi
'' SIG '' aNE9lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0
'' SIG '' cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJv
'' SIG '' ZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3Js
'' SIG '' MFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0
'' SIG '' cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9N
'' SIG '' aWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwDQYJKoZI
'' SIG '' hvcNAQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/yp
'' SIG '' b+pcFLY+TkdkeLEGk5c9MTO1OdfCcTY/2mRsfNB1OW27
'' SIG '' DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYeeNg2LpypglYA
'' SIG '' A7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM9W0j
'' SIG '' VOR4U3UkV7ndn/OOPcbzaN9l9qRWqveVtihVJ9AkvUCg
'' SIG '' vxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSFQrP8DJ6LGYnn
'' SIG '' 8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4FOmRsqlb
'' SIG '' 30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW9I/2kQH2zsZ0
'' SIG '' /fZMcm8Qq3UwxTSwethQ/gpY3UA8x1RtnWN0SCyxTkct
'' SIG '' wRQEcb9k+SS+c23Kjgm9swFXSVRk2XPXfx5bRAGOWhmR
'' SIG '' aw2fpCjcZxkoJLo4S5pu+yFUa2pFEUep8beuyOiJXk+d
'' SIG '' 0tBMdrVXVAmxaQFEfnyhYWxz/gq77EFmPWn9y8FBSX5+
'' SIG '' k77L+DvktxW/tM4+pTFRhLy/AsGConsXHRWJjXD+57XQ
'' SIG '' KBqJC4822rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdC
'' SIG '' PSWU5nR0W2rRnj7tfqAxM328y+l7vzhwRNGQ8cirOoo6
'' SIG '' CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGbyoYIC
'' SIG '' 1DCCAj0CAQEwggEAoYHYpIHVMIHSMQswCQYDVQQGEwJV
'' SIG '' UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
'' SIG '' UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFu
'' SIG '' ZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRo
'' SIG '' YWxlcyBUU1MgRVNOOjJBRDQtNEI5Mi1GQTAxMSUwIwYD
'' SIG '' VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
'' SIG '' oiMKAQEwBwYFKw4DAhoDFQBooFKKzLjLzqmXxfLbYIlk
'' SIG '' TETa86CBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
'' SIG '' VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
'' SIG '' MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
'' SIG '' JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
'' SIG '' QSAyMDEwMA0GCSqGSIb3DQEBBQUAAgUA6YEIaDAiGA8y
'' SIG '' MDI0MDIyMjA3NDAyNFoYDzIwMjQwMjIzMDc0MDI0WjB0
'' SIG '' MDoGCisGAQQBhFkKBAExLDAqMAoCBQDpgQhoAgEAMAcC
'' SIG '' AQACAhTfMAcCAQACAhFEMAoCBQDpglnoAgEAMDYGCisG
'' SIG '' AQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEA
'' SIG '' AgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQEFBQAD
'' SIG '' gYEAI89eh+PA4W8LPS7pAlP8V5uZbe0Q+pGqswFjhIXi
'' SIG '' 0CeQQA4IyUMR2kDRsLduxQGv0Hk3r1sKb4KskI9nI5lX
'' SIG '' v7HToqbxPIj12L5dZWeO+j/2/f/5o1EkpdaeC63eu6cg
'' SIG '' nkYAbckttP5PQY64JLVX1k6Ylk5UxnRzT5Em9ygxbDYx
'' SIG '' ggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEG
'' SIG '' A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
'' SIG '' ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
'' SIG '' MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
'' SIG '' Q0EgMjAxMAITMwAAAd6eSJ6WnyhEPQABAAAB3jANBglg
'' SIG '' hkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqG
'' SIG '' SIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCCoOW/+Tx4w
'' SIG '' mEf7yrfkOgCYWmZ2aWxb4gr/fplMxYhTdjCB+gYLKoZI
'' SIG '' hvcNAQkQAi8xgeowgecwgeQwgb0EII4+I58NwV4QEEkC
'' SIG '' f+YLcyCtPnD9TbPzUtgPjgdzfh17MIGYMIGApH4wfDEL
'' SIG '' MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
'' SIG '' EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
'' SIG '' c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
'' SIG '' b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAHenkie
'' SIG '' lp8oRD0AAQAAAd4wIgQgyvlPgu1Q7t7AnmY92dV5Tz9s
'' SIG '' A7rjvxyuoxTSWrEViQIwDQYJKoZIhvcNAQELBQAEggIA
'' SIG '' K406dw6lf+vGGOrrWxL3SzR1MCEZ68JIiqSsan8NCR0j
'' SIG '' uedTH+L2rqKUGd0WqBW/GzeGwwZd8kIJwz/YNFGtOXiP
'' SIG '' 59geDaiE79hVt1QLvy3+M4wf+nWyOyUR1hJR6xi7U4Pp
'' SIG '' 7HtLRrEQRGB/hJjgKIXPe3P9lml51Q8h0rLspup0viDr
'' SIG '' LdYV1++vE7I3f05flVjH2GSGZPFBhP3Qdo0skC0jPovz
'' SIG '' mVI06xEop+ayI8CqRsSWjuvmUxxYvhHMuYlprKenGjkW
'' SIG '' etBi6CNRCoyrZtsslxS+OixZGPORx5GUxkqhTNPI5XkU
'' SIG '' BETZ5IK9/Ue//YVJ13hNelsqp7XA+2Gy1Fz0Phr7vKJW
'' SIG '' gHoOlljSXMLoOpohLNhsVcZC5nQvQOViNylYTc2N5KNF
'' SIG '' 00xbeemHPpXZXshbT8U9c1fHMibZTtthgatCJWSJVv0W
'' SIG '' YiJcJnXHOCWysKq6sIKifEEJGBs9+KiNHqRlj/yPBqCR
'' SIG '' xmuN3Q14ea8A5Jz0oHOEhazZ9n+AhlxCafx9SV75Q6xk
'' SIG '' eMw1j4smAG2BsICge7DlYvQfJaQqOxXjPDn9n1QpvLxf
'' SIG '' izmTUL0KU8yG6Zn9NR/gcgimhnwdXNonq/Fb0mZK3yCJ
'' SIG '' q41Wv7SQ0wfK4kZuKtIE/UFGjQvXgLs2/2BhCLiyhRHo
'' SIG '' fGepOxYI80jeAOniUr15peM=
'' SIG '' End signature block
