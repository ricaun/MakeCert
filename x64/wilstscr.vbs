' Windows Installer script viewer for use with Windows Scripting Host CScript.exe only
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of the special database processing mode for viewing script files
'
Option Explicit

Const msiOpenDatabaseModeListScript = 5

' Check arg count, and display help if argument not present or contains ?
Dim argCount:argCount = Wscript.Arguments.Count
If argCount > 0 Then If InStr(1, Wscript.Arguments(0), "?", vbTextCompare) > 0 Then argCount = 0
If argCount = 0 Then
	Wscript.Echo "Windows Installer Script Viewer for Windows Scripting Host (CScript.exe)" &_
		vbNewLine & " Argument is path to installer execution script" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Cannot run with GUI script host, as listing is performed to standard out
If UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "W" Then
	Wscript.Echo "Cannot use WScript.exe - must use CScript.exe with this program"
	Wscript.Quit 2
End If

Dim installer, view, database, record, fieldCount, template, index, field
On Error Resume Next
Set installer = CreateObject("WindowsInstaller.Installer") : CheckError
Set database = installer.Opendatabase(Wscript.Arguments(0), msiOpenDatabaseModeListScript) : CheckError
Set view = database.Openview("")
view.Execute : CheckError
Do
   Set record = view.Fetch
   If record Is Nothing Then Exit Do
   fieldCount = record.FieldCount
   template = record.StringData(0)
   index = InstrRev(template, "[") + 1
   If (index > 1) Then
      field = Int(Mid(template, index, InstrRev(template, "]") - index))
      If field < fieldCount Then
         template = Left(template, Len(template) - 1)
         While field < fieldCount
            field = field + 1
            template = template & ",[" & field & "]"
         Wend
         record.StringData(0) = template & ")"
      End If
   End If
   Wscript.Echo record.FormatText
Loop
Wscript.Quit 0

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
'' SIG '' MIImFgYJKoZIhvcNAQcCoIImBzCCJgMCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' yGbucx5C9ty0NdJlcwpY0JNDXmCZOM9FQgmr4/kXaQOg
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
'' SIG '' MYIZ+DCCGfQCAQEwgZUwfjELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmlu
'' SIG '' ZyBQQ0EgMjAxMAITMwAABVbJICsfdDJdLQAAAAAFVjAN
'' SIG '' BglghkgBZQMEAgEFAKCCAQQwGQYJKoZIhvcNAQkDMQwG
'' SIG '' CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisG
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEICM1o7gzYsS+
'' SIG '' pp2C6d0mLCxT/Z30fa3Bp+PDJHTWK6/cMDwGCisGAQQB
'' SIG '' gjcKAxwxLgwsc1BZN3hQQjdoVDVnNUhIcll0OHJETFNN
'' SIG '' OVZ1WlJ1V1phZWYyZTIyUnM1ND0wWgYKKwYBBAGCNwIB
'' SIG '' DDFMMEqgJIAiAE0AaQBjAHIAbwBzAG8AZgB0ACAAVwBp
'' SIG '' AG4AZABvAHcAc6EigCBodHRwOi8vd3d3Lm1pY3Jvc29m
'' SIG '' dC5jb20vd2luZG93czANBgkqhkiG9w0BAQEFAASCAQBW
'' SIG '' stLQK7tRBzEnIvo9k8b8VqfFz+7BJnZOgPY58zD1ixNo
'' SIG '' 0Lwr6tTL1CPEqw6JEck84tZwoTJtHaxH1yZ6xzas4n6E
'' SIG '' klG/HFHuHBMC8ctPM2TzvKNM8DTUen3UQY59d0lvc+ld
'' SIG '' i29C7Cn94VgP1XS+BYGF7JDOteybjLNpApvn4Iel2oKU
'' SIG '' w0gdfyH8ZpMaI5HsCRXYX3W/81ASr05C86ESQlYgQacv
'' SIG '' 1emnSZd9U+ddPkUraXmoPOpBT5bIIW5y+Vh0H6tGvzCv
'' SIG '' 0yThdkBgKwzc85uEO6QfAYVh/gileC08CcjOMpwLpYR3
'' SIG '' hldr0YvD32Q3UmWBJLKGytV1e9ovSoCUoYIXKzCCFycG
'' SIG '' CisGAQQBgjcDAwExghcXMIIXEwYJKoZIhvcNAQcCoIIX
'' SIG '' BDCCFwACAQMxDzANBglghkgBZQMEAgEFADCCAVkGCyqG
'' SIG '' SIb3DQEJEAEEoIIBSASCAUQwggFAAgEBBgorBgEEAYRZ
'' SIG '' CgMBMDEwDQYJYIZIAWUDBAIBBQAEIAoC4avYI+syEYQU
'' SIG '' Z60G2mOT3BWCP0/mhiftwRm1vxl/AgZl1gZE4OwYEzIw
'' SIG '' MjQwMjIyMTA0NDI2LjU1M1owBIACAfSggdikgdUwgdIx
'' SIG '' CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
'' SIG '' MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jv
'' SIG '' c29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEm
'' SIG '' MCQGA1UECxMdVGhhbGVzIFRTUyBFU046MTc5RS00QkIw
'' SIG '' LTgyNDYxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0
'' SIG '' YW1wIFNlcnZpY2WgghF6MIIHJzCCBQ+gAwIBAgITMwAA
'' SIG '' AeDU/B8TFR9+XQABAAAB4DANBgkqhkiG9w0BAQsFADB8
'' SIG '' MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
'' SIG '' bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
'' SIG '' cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
'' SIG '' b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzEw
'' SIG '' MTIxOTA3MTlaFw0yNTAxMTAxOTA3MTlaMIHSMQswCQYD
'' SIG '' VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
'' SIG '' A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
'' SIG '' IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQg
'' SIG '' SXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNV
'' SIG '' BAsTHVRoYWxlcyBUU1MgRVNOOjE3OUUtNEJCMC04MjQ2
'' SIG '' MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
'' SIG '' ZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
'' SIG '' CgKCAgEArIec86HFu9EBOcaNv/p+4GGHdkvOi0DECB0t
'' SIG '' pn/OREVR15IrPI23e2qiswrsYO9xd0qz6ogxRu96eUf7
'' SIG '' Dneyw9rqtg/vrRm4WsAGt+x6t/SQVrI1dXPBPuNqsk4S
'' SIG '' OcUwGn7KL67BDZOcm7FzNx4bkUMesgjqwXoXzv2U/rJ1
'' SIG '' jQEFmRn23f17+y81GJ4DmBSe/9hwz9sgxj9BiZ30XQH5
'' SIG '' 5sViL48fgCRdqE2QWArzk4hpGsMa+GfE5r/nMYvs6KKL
'' SIG '' v4n39AeR0kaV+dF9tDdBcz/n+6YE4obgmgVjWeJnlFUf
'' SIG '' k9PT64KPByqFNue9S18r437IHZv2sRm+nZO/hnBjMR30
'' SIG '' D1Wxgy5mIJJtoUyTvsvBVuSWmfDhodYlcmQRiYm/FFtx
'' SIG '' OETwVDI6hWRK4pzk5Znb5Yz+PnShuUDS0JTncBq69Q5l
'' SIG '' GhAGHz2ccr6bmk5cpd1gwn5x64tgXyHnL9xctAw6aosn
'' SIG '' PmXswuobBTTMdX4wQ7wvUWjbMQRDiIvgFfxiScpeiccZ
'' SIG '' BpxIJotmi3aTIlVGwVLGfQ+U+8dWnRh2wIzN16LD2MBn
'' SIG '' sr2zVbGxkYQGsr+huKlfq7GMSnJQD2ZtU+WOVvdHgxYj
'' SIG '' QTbEj80zoXgBzwJ5rHdhYtP5pYJl6qIgwvHLJZmD6LUp
'' SIG '' jxkTMx41MoIQjnAXXDGqvpPX8xCj7y0CAwEAAaOCAUkw
'' SIG '' ggFFMB0GA1UdDgQWBBRwXhc/bp1X7xK6ygDVddDZMNKZ
'' SIG '' 0jAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnp
'' SIG '' cjBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1p
'' SIG '' Y3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQl
'' SIG '' MjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmww
'' SIG '' bAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRw
'' SIG '' Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRz
'' SIG '' L01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAy
'' SIG '' MDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB
'' SIG '' /wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDAN
'' SIG '' BgkqhkiG9w0BAQsFAAOCAgEAwBPODpH8DSV07syobEPV
'' SIG '' UmOLnJUDWEdvQdzRiO2/taTFDyLB9+W6VflSzri0Pf7c
'' SIG '' 1PUmSmFbNoBZ/bAp0DDflHG1AbWI43ccRnRfbed17gqD
'' SIG '' 9Z9vHmsQeRn1vMqdH/Y3kDXr7D/WlvAnN19FyclPdwvJ
'' SIG '' rCv+RiMxZ3rc4/QaWrvS5rhZQT8+jmlTutBFtYShCjNj
'' SIG '' biECo5zC5FyboJvQkF5M4J5EGe0QqCMp6nilFpC3tv2+
'' SIG '' 6xP3tZ4lx9pWiyaY+2xmxrCCekiNsFrnm0d+6TS8ORm1
'' SIG '' sheNTiavl2ez12dqcF0FLY9jc3eEh8I8Q6zOq7AcuR+Q
'' SIG '' Vn/1vHDz95EmV22i6QejXpp8T8Co/+yaYYmHllHSmaBb
'' SIG '' pBxf7rWt2LmQMlPMIVqgzJjNRLRIRvKsNn+nYo64oBg2
'' SIG '' eCWOI6WWVy3S4lXPZqB9zMaOOwqLYBLVZpe86GBk2YbD
'' SIG '' jZIUHWpqWhrwpq7H1DYccsTyB57/muA6fH3NJt9VRzsh
'' SIG '' xE2h2rpHu/5HP4/pcq06DIKpb/6uE+an+fsWrYEZNGRz
'' SIG '' L/+GZLfanqrKCWvYrg6gkMlfEWzqXBzwPzqqVR4aNTKj
'' SIG '' uFXLlW/ID7LSYacQC4Dzm2w5xQ+XPBYXmy/4Hl/Pfk5b
'' SIG '' dfhKmTlKI26WcsVE8zlcKxIeq9xsLxHerCPbDV68+FnE
'' SIG '' O40wggdxMIIFWaADAgECAhMzAAAAFcXna54Cm0mZAAAA
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
'' SIG '' 1jCCAj8CAQEwggEAoYHYpIHVMIHSMQswCQYDVQQGEwJV
'' SIG '' UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
'' SIG '' UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFu
'' SIG '' ZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRo
'' SIG '' YWxlcyBUU1MgRVNOOjE3OUUtNEJCMC04MjQ2MSUwIwYD
'' SIG '' VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
'' SIG '' oiMKAQEwBwYFKw4DAhoDFQBt89HV8FfofFh/I/HzNjMl
'' SIG '' Tl8hDKCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
'' SIG '' VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
'' SIG '' MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
'' SIG '' JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
'' SIG '' QSAyMDEwMA0GCSqGSIb3DQEBBQUAAgUA6YCEjjAiGA8y
'' SIG '' MDI0MDIyMTIyMTc1MFoYDzIwMjQwMjIyMjIxNzUwWjB2
'' SIG '' MDwGCisGAQQBhFkKBAExLjAsMAoCBQDpgISOAgEAMAkC
'' SIG '' AQACARECAf8wBwIBAAICEWAwCgIFAOmB1g4CAQAwNgYK
'' SIG '' KwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgC
'' SIG '' AQACAwehIKEKMAgCAQACAwGGoDANBgkqhkiG9w0BAQUF
'' SIG '' AAOBgQB2/cIqLDPOHnt8QbgxJC+grtvmkjGfymZ9k9mM
'' SIG '' /raoM37cTAuQH7ZHc11P2AlTcPMw2+v0wJxnPKuXT45i
'' SIG '' n6STaNH82WjCnLxAxnmCSf9irJROKSNO6XKK8wBmQZjh
'' SIG '' OHjLdOreuJ4S7c/1tP/XAvHeYYjnHaYvVTh9QSELjXzn
'' SIG '' MTGCBA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMw
'' SIG '' EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
'' SIG '' b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
'' SIG '' b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
'' SIG '' IFBDQSAyMDEwAhMzAAAB4NT8HxMVH35dAAEAAAHgMA0G
'' SIG '' CWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYL
'' SIG '' KoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIE+zsl9j
'' SIG '' h6jF1ZcZB0TmONRcCG3ZIIOr7CqJs7RSPLr/MIH6Bgsq
'' SIG '' hkiG9w0BCRACLzGB6jCB5zCB5DCBvQQg4+5Sv/I55W04
'' SIG '' z73O+wwgkm+E2QKWPZyZucIbCv9pCsEwgZgwgYCkfjB8
'' SIG '' MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
'' SIG '' bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
'' SIG '' cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
'' SIG '' b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAeDU
'' SIG '' /B8TFR9+XQABAAAB4DAiBCC/ZsCBpGJaKDIHmmfGZb+9
'' SIG '' 0sLFjaszn4ArkzMFIXK69jANBgkqhkiG9w0BAQsFAASC
'' SIG '' AgAcmWGXRZqLBXY9vjvuO33c3bm+L99PJLYwljwZpsNy
'' SIG '' SgtbTZRfbcXXZ7IHUgAu/1sO0o80IYgi2ymHNonHvVB3
'' SIG '' aBGrDtDB0Ox0dPTSXK/YPQO7sNR20gRe+dxJ/81eZSey
'' SIG '' yroSmf8ZXAT53BG2yY9o/zs0xKt1tCY8M0CbjGLgJaOJ
'' SIG '' fzlG8I9k1+w6s3F78BYjEOD6iUY6QasSqb1BVVTJnQmC
'' SIG '' 1ZLbIugjQ3lDa7FYx4EJbfx5CpxhRLpLrPCQRUpMS5Wq
'' SIG '' /wqkSoZ1OZyIeju/KRhEQcN3IB6gZxHOs5NxQ+mN242Y
'' SIG '' 1VAfSzwB7vHhrtbsjnd0yYjJ4tz+OXozTyqHc7CBJ0ya
'' SIG '' Iu4GpOU2ZZXzxJ2TetZHmEGtBf/wdeV0nQ/Pcc5jM03n
'' SIG '' Ooub4mlSM9uRoZoxJxQFCwgNy30g+3bq3bOjLAQbOhb2
'' SIG '' 7UKRWQvlI6DXpYMHcsujUC4fJK1fFmxXTk0Pen93LBuH
'' SIG '' 9wU08Ai9p+0Xynev6zbhP9/w5VMYk5R5K4GQD+UWSZZ9
'' SIG '' yOtEHbWGBJulTSrYRiAqKC03UMko7SKQPCYc1jOXxhyY
'' SIG '' uxHJM7wbB4MxilMpQr1K8KHX+5WLKwrhc0NwzA9KThgy
'' SIG '' Kmu/wV08IeyyWtB3bbB7JTOzb2Xfw3UlpgaieHKip3ws
'' SIG '' ca9rCRyC+p6SF0p7NK6mex0/Ng==
'' SIG '' End signature block
