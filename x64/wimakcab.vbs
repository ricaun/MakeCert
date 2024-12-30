' Windows Installer utility to generate file cabinets from MSI database
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the access to install engine and actions
'
Option Explicit

' FileSystemObject.CreateTextFile and FileSystemObject.OpenTextFile
Const OpenAsASCII   = 0 
Const OpenAsUnicode = -1

' FileSystemObject.CreateTextFile
Const OverwriteIfExist = -1
Const FailIfExist      = 0

' FileSystemObject.OpenTextFile
Const OpenAsDefault    = -2
Const CreateIfNotExist = -1
Const FailIfNotExist   = 0
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Const msiOpenDatabaseModeReadOnly = 0
Const msiOpenDatabaseModeTransact = 1

Const msiViewModifyInsert         = 1
Const msiViewModifyUpdate         = 2
Const msiViewModifyAssign         = 3
Const msiViewModifyReplace        = 4
Const msiViewModifyDelete         = 6

Const msiUILevelNone = 2

Const msiRunModeSourceShortNames = 9

Const msidbFileAttributesNoncompressed = &h00002000

Dim argCount:argCount = Wscript.Arguments.Count
Dim iArg:iArg = 0
If argCount > 0 Then If InStr(1, Wscript.Arguments(0), "?", vbTextCompare) > 0 Then argCount = 0
If (argCount < 2) Then
	Wscript.Echo "Windows Installer utility to generate compressed file cabinets from MSI database" &_
		vbNewLine & " The 1st argument is the path to MSI database, at the source file root" &_
		vbNewLine & " The 2nd argument is the base name used for the generated files (DDF, INF, RPT)" &_
		vbNewLine & " The 3rd argument can optionally specify separate source location from the MSI" &_
		vbNewLine & " The following options may be specified at any point on the command line" &_
		vbNewLine & "  /L to use LZX compression instead of MSZIP" &_
		vbNewLine & "  /F to limit cabinet size to 1.44 MB floppy size rather than CD" &_
		vbNewLine & "  /C to run compression, else only generates the .DDF file" &_
		vbNewLine & "  /U to update the MSI database to reference the generated cabinet" &_
		vbNewLine & "  /E to embed the cabinet file in the installer package as a stream" &_
		vbNewLine & "  /S to sequence number file table, ordered by directories" &_
		vbNewLine & "  /R to revert to non-cabinet install, removes cabinet if /E specified" &_
		vbNewLine & " Notes:" &_
		vbNewLine & "  In order to generate a cabinet, MAKECAB.EXE must be on the PATH" &_
		vbNewLine & "  base name used for files and cabinet stream is case-sensitive" &_
		vbNewLine & "  If source type set to compressed, all files will be opened at the root" &_
		vbNewLine & "  (The /R option removes the compressed bit - SummaryInfo property 15 & 2)" &_
		vbNewLine & "  To replace an embedded cabinet, include the options: /R /C /U /E" &_
		vbNewLine & "  Does not handle updating of Media table to handle multiple cabinets" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Get argument values, processing any option flags
Dim compressType : compressType = "MSZIP"
Dim cabSize      : cabSize      = "CDROM"
Dim makeCab      : makeCab      = False
Dim embedCab     : embedCab     = False
Dim updateMsi    : updateMsi    = False
Dim sequenceFile : sequenceFile = False
Dim removeCab    : removeCab    = False
Dim databasePath : databasePath = NextArgument
Dim baseName     : baseName     = NextArgument
Dim sourceFolder : sourceFolder = NextArgument
If Not IsEmpty(NextArgument) Then Fail "More than 3 arguments supplied" ' process any trailing options
If Len(baseName) < 1 Or Len(baseName) > 8 Then Fail "Base file name must be from 1 to 8 characters"
If Not IsEmpty(sourceFolder) And Right(sourceFolder, 1) <> "\" Then sourceFolder = sourceFolder & "\"
Dim cabFile : cabFile = baseName & ".CAB"
Dim cabName : cabName = cabFile : If embedCab Then cabName = "#" & cabName

' Connect to Windows Installer object
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError

' Open database
Dim database, openMode, view, record, updateMode, sumInfo, sequence, lastSequence
If updateMsi Or sequenceFile Or removeCab Then openMode = msiOpenDatabaseModeTransact Else openMode = msiOpenDatabaseModeReadOnly
Set database = installer.OpenDatabase(databasePath, openMode) : CheckError

' Remove existing cabinet(s) and revert to source tree install if options specified
If removeCab Then
	Set view = database.OpenView("SELECT DiskId, LastSequence, Cabinet FROM Media ORDER BY DiskId") : CheckError
	view.Execute : CheckError
	updateMode = msiViewModifyUpdate
	Set record = view.Fetch : CheckError
	If Not record Is Nothing Then ' Media table not empty
		If Not record.IsNull(3) Then
			If record.StringData(3) <> cabName Then Wscript.Echo "Warning, cabinet name in media table, " & record.StringData(3) & " does not match " & cabName
			record.StringData(3) = Empty
		End If
		record.IntegerData(2) = 9999 ' in case of multiple cabinets, force all files from 1st media
		view.Modify msiViewModifyUpdate, record : CheckError
		Do
			Set record = view.Fetch : CheckError
			If record Is Nothing Then Exit Do
			view.Modify msiViewModifyDelete, record : CheckError 'remove other cabinet records
		Loop
	End If
	Set sumInfo = database.SummaryInformation(3) : CheckError
	sumInfo.Property(11) = Now
	sumInfo.Property(13) = Now
	sumInfo.Property(15) = sumInfo.Property(15) And Not 2
	sumInfo.Persist
	Set view = database.OpenView("SELECT `Name`,`Data` FROM _Streams WHERE `Name`= '" & cabFile & "'") : CheckError
	view.Execute : CheckError
	Set record = view.Fetch
	If record Is Nothing Then
		Wscript.Echo "Warning, cabinet stream not found in package: " & cabFile
	Else
		view.Modify msiViewModifyDelete, record : CheckError
	End If
	Set sumInfo = Nothing ' must release stream
	database.Commit : CheckError
	If Not updateMsi Then Wscript.Quit 0
End If

' Create an install session and execute actions in order to perform directory resolution
installer.UILevel = msiUILevelNone
Dim session : Set session = installer.OpenPackage(database,1) : If Err <> 0 Then Fail "Database: " & databasePath & ". Invalid installer package format"
Dim shortNames : shortNames = session.Mode(msiRunModeSourceShortNames) : CheckError
If Not IsEmpty(sourceFolder) Then session.Property("OriginalDatabase") = sourceFolder : CheckError
Dim stat : stat = session.DoAction("CostInitialize") : CheckError
If stat <> 1 Then Fail "CostInitialize failed, returned " & stat

' Check for non-cabinet files to avoid sequence number collisions
lastSequence = 0
If sequenceFile Then
	Set view = database.OpenView("SELECT Sequence,Attributes FROM File") : CheckError
	view.Execute : CheckError
	Do
		Set record = view.Fetch : CheckError
		If record Is Nothing Then Exit Do
		sequence = record.IntegerData(1)
		If (record.IntegerData(2) And msidbFileAttributesNoncompressed) <> 0 And sequence > lastSequence Then lastSequence = sequence
	Loop	
End If

' Join File table to Component table in order to find directories
Dim orderBy : If sequenceFile Then orderBy = "Directory_" Else orderBy = "Sequence"
Set view = database.OpenView("SELECT File,FileName,Directory_,Sequence,File.Attributes FROM File,Component WHERE Component_=Component ORDER BY " & orderBy) : CheckError
view.Execute : CheckError

' Create DDF file and write header properties
Dim FileSys : Set FileSys = CreateObject("Scripting.FileSystemObject") : CheckError
Dim outStream : Set outStream = FileSys.CreateTextFile(baseName & ".DDF", OverwriteIfExist, OpenAsASCII) : CheckError
outStream.WriteLine "; Generated from " & databasePath & " on " & Now
outStream.WriteLine ".Set CabinetNameTemplate=" & baseName & "*.CAB"
outStream.WriteLine ".Set CabinetName1=" & cabFile
outStream.WriteLine ".Set ReservePerCabinetSize=8"
outStream.WriteLine ".Set MaxDiskSize=" & cabSize
outStream.WriteLine ".Set CompressionType=" & compressType
outStream.WriteLine ".Set InfFileLineFormat=(*disk#*) *file#*: *file* = *Size*"
outStream.WriteLine ".Set InfFileName=" & baseName & ".INF"
outStream.WriteLine ".Set RptFileName=" & baseName & ".RPT"
outStream.WriteLine ".Set InfHeader="
outStream.WriteLine ".Set InfFooter="
outStream.WriteLine ".Set DiskDirectoryTemplate=."
outStream.WriteLine ".Set Compress=ON"
outStream.WriteLine ".Set Cabinet=ON"

' Fetch each file and request the source path, then verify the source path
Dim fileKey, fileName, folder, sourcePath, delim, message, attributes
Do
	Set record = view.Fetch : CheckError
	If record Is Nothing Then Exit Do
	fileKey    = record.StringData(1)
	fileName   = record.StringData(2)
	folder     = record.StringData(3)
	sequence   = record.IntegerData(4)
	attributes = record.IntegerData(5)
	If (attributes And msidbFileAttributesNoncompressed) = 0 Then
		If sequence <= lastSequence Then
			If Not sequenceFile Then Fail "Duplicate sequence numbers in File table, use /S option"
			sequence = lastSequence + 1
			record.IntegerData(4) = sequence
			view.Modify msiViewModifyUpdate, record
		End If
		lastSequence = sequence
		delim = InStr(1, fileName, "|", vbTextCompare)
		If delim <> 0 Then
			If shortNames Then fileName = Left(fileName, delim-1) Else fileName = Right(fileName, Len(fileName) - delim)
		End If
		sourcePath = session.SourcePath(folder) & fileName
		outStream.WriteLine """" & sourcePath & """" & " " & fileKey
		If installer.FileAttributes(sourcePath) = -1 Then message = message & vbNewLine & sourcePath
	End If
Loop
outStream.Close
REM Wscript.Echo "SourceDir = " & session.Property("SourceDir")
If Not IsEmpty(message) Then Fail "The following files were not available:" & message

' Generate compressed file cabinet
If makeCab Then
	Dim WshShell : Set WshShell = Wscript.CreateObject("Wscript.Shell") : CheckError
	Dim cabStat : cabStat = WshShell.Run("MakeCab.exe /f " & baseName & ".DDF", 7, True) : CheckError
	If cabStat <> 0 Then Fail "MAKECAB.EXE failed, possibly could not find source files, or invalid DDF format"
End If

' Update Media table and SummaryInformation if requested
If updateMsi Then
	Set view = database.OpenView("SELECT DiskId, LastSequence, Cabinet FROM Media ORDER BY DiskId") : CheckError
	view.Execute : CheckError
	updateMode = msiViewModifyUpdate
	Set record = view.Fetch : CheckError
	If record Is Nothing Then ' Media table empty
		Set record = Installer.CreateRecord(3)
		record.IntegerData(1) = 1
		updateMode = msiViewModifyInsert
	End If
	record.IntegerData(2) = lastSequence
	record.StringData(3) = cabName
	view.Modify updateMode, record
	Set sumInfo = database.SummaryInformation(3) : CheckError
	sumInfo.Property(11) = Now
	sumInfo.Property(13) = Now
	sumInfo.Property(15) = (shortNames And 1) + 2
	sumInfo.Persist
End If

' Embed cabinet if requested
If embedCab Then
	Set view = database.OpenView("SELECT `Name`,`Data` FROM _Streams") : CheckError
	view.Execute : CheckError
	Set record = Installer.CreateRecord(2)
	record.StringData(1) = cabFile
	record.SetStream 2, cabFile : CheckError
	view.Modify msiViewModifyAssign, record : CheckError 'replace any existing stream of that name
End If

' Commit database in case updates performed
database.Commit : CheckError
Wscript.Quit 0

' Extract argument value from command line, processing any option flags
Function NextArgument
	Dim arg
	Do  ' loop to pull in option flags until an argument value is found
		If iArg >= argCount Then Exit Function
		arg = Wscript.Arguments(iArg)
		iArg = iArg + 1
		If (AscW(arg) <> AscW("/")) And (AscW(arg) <> AscW("-")) Then Exit Do
		Select Case UCase(Right(arg, Len(arg)-1))
			Case "C" : makeCab      = True
			Case "E" : embedCab     = True
			Case "F" : cabSize      = "1.44M"
			Case "L" : compressType = "LZX"
			Case "R" : removeCab    = True
			Case "S" : sequenceFile = True
			Case "U" : updateMsi    = True
			Case Else: Wscript.Echo "Invalid option flag:", arg : Wscript.Quit 1
		End Select
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
	Fail message
End Sub

Sub Fail(message)
	Wscript.Echo message
	Wscript.Quit 2
End Sub

'' SIG '' Begin signature block
'' SIG '' MIImBwYJKoZIhvcNAQcCoIIl+DCCJfQCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' +3czCZ7bOIQLc7kJN8m8lyJpE9uBKr7KXjIe8c3/+0yg
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
'' SIG '' jgd7JXFEqwZq5tTG3yOalnXFMYIZ+DCCGfQCAQEwgZUw
'' SIG '' fjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWlj
'' SIG '' cm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMAITMwAA
'' SIG '' BVfPkN3H0cCIjAAAAAAFVzANBglghkgBZQMEAgEFAKCC
'' SIG '' AQQwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
'' SIG '' KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
'' SIG '' hvcNAQkEMSIEIMQxWWH3f0SoxzhOQoL/5WBpLL+Mh9+5
'' SIG '' YC4NsiyvzvYfMDwGCisGAQQBgjcKAxwxLgwsc1BZN3hQ
'' SIG '' QjdoVDVnNUhIcll0OHJETFNNOVZ1WlJ1V1phZWYyZTIy
'' SIG '' UnM1ND0wWgYKKwYBBAGCNwIBDDFMMEqgJIAiAE0AaQBj
'' SIG '' AHIAbwBzAG8AZgB0ACAAVwBpAG4AZABvAHcAc6EigCBo
'' SIG '' dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vd2luZG93czAN
'' SIG '' BgkqhkiG9w0BAQEFAASCAQB7Trb+Y7JytoV5SRlDxoxW
'' SIG '' q/KHcJNf7htYw1x/x/BTYNUVLKe1GIzoQOMhcokv/t8A
'' SIG '' PW2JIWLoK+x/Sd3l+IWsjUP6sKax+q7eXTlUl6cm6pNf
'' SIG '' M18XQybJfdBY6FV30ALS8ublPIJNpyujiKmdFIrUwiSg
'' SIG '' eZykVgttCAQmyH5uz3p3wUosVmdZIzJ0pIHilzsaQDb7
'' SIG '' izyA99/zzfNWfQdZkMyPDZUQ0lziHJiyg9R77gtN4KSb
'' SIG '' iNmEfZCWQkATeBp/RTjtcBV0JiqIqCkjpUWQWP05YsSb
'' SIG '' 74Ex3j+yLJv5AGVIZW1Bj+QKyOXGJCVn2BD15kpPbkta
'' SIG '' b8IhP1arlbpuoYIXKzCCFycGCisGAQQBgjcDAwExghcX
'' SIG '' MIIXEwYJKoZIhvcNAQcCoIIXBDCCFwACAQMxDzANBglg
'' SIG '' hkgBZQMEAgEFADCCAVkGCyqGSIb3DQEJEAEEoIIBSASC
'' SIG '' AUQwggFAAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUD
'' SIG '' BAIBBQAEICmB3wRmilA4FqVgxZOMoOawky3bh1B5T6FO
'' SIG '' kSEp0//XAgZl1gZE39EYEzIwMjQwMjIyMTA0NDIyLjkw
'' SIG '' NlowBIACAfSggdikgdUwgdIxCzAJBgNVBAYTAlVTMRMw
'' SIG '' EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
'' SIG '' b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
'' SIG '' b24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9w
'' SIG '' ZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVz
'' SIG '' IFRTUyBFU046MTc5RS00QkIwLTgyNDYxJTAjBgNVBAMT
'' SIG '' HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WgghF6
'' SIG '' MIIHJzCCBQ+gAwIBAgITMwAAAeDU/B8TFR9+XQABAAAB
'' SIG '' 4DANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzET
'' SIG '' MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
'' SIG '' bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
'' SIG '' aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
'' SIG '' cCBQQ0EgMjAxMDAeFw0yMzEwMTIxOTA3MTlaFw0yNTAx
'' SIG '' MTAxOTA3MTlaMIHSMQswCQYDVQQGEwJVUzETMBEGA1UE
'' SIG '' CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEe
'' SIG '' MBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0w
'' SIG '' KwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRp
'' SIG '' b25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
'' SIG '' RVNOOjE3OUUtNEJCMC04MjQ2MSUwIwYDVQQDExxNaWNy
'' SIG '' b3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkq
'' SIG '' hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEArIec86HFu9EB
'' SIG '' OcaNv/p+4GGHdkvOi0DECB0tpn/OREVR15IrPI23e2qi
'' SIG '' swrsYO9xd0qz6ogxRu96eUf7Dneyw9rqtg/vrRm4WsAG
'' SIG '' t+x6t/SQVrI1dXPBPuNqsk4SOcUwGn7KL67BDZOcm7Fz
'' SIG '' Nx4bkUMesgjqwXoXzv2U/rJ1jQEFmRn23f17+y81GJ4D
'' SIG '' mBSe/9hwz9sgxj9BiZ30XQH55sViL48fgCRdqE2QWArz
'' SIG '' k4hpGsMa+GfE5r/nMYvs6KKLv4n39AeR0kaV+dF9tDdB
'' SIG '' cz/n+6YE4obgmgVjWeJnlFUfk9PT64KPByqFNue9S18r
'' SIG '' 437IHZv2sRm+nZO/hnBjMR30D1Wxgy5mIJJtoUyTvsvB
'' SIG '' VuSWmfDhodYlcmQRiYm/FFtxOETwVDI6hWRK4pzk5Znb
'' SIG '' 5Yz+PnShuUDS0JTncBq69Q5lGhAGHz2ccr6bmk5cpd1g
'' SIG '' wn5x64tgXyHnL9xctAw6aosnPmXswuobBTTMdX4wQ7wv
'' SIG '' UWjbMQRDiIvgFfxiScpeiccZBpxIJotmi3aTIlVGwVLG
'' SIG '' fQ+U+8dWnRh2wIzN16LD2MBnsr2zVbGxkYQGsr+huKlf
'' SIG '' q7GMSnJQD2ZtU+WOVvdHgxYjQTbEj80zoXgBzwJ5rHdh
'' SIG '' YtP5pYJl6qIgwvHLJZmD6LUpjxkTMx41MoIQjnAXXDGq
'' SIG '' vpPX8xCj7y0CAwEAAaOCAUkwggFFMB0GA1UdDgQWBBRw
'' SIG '' Xhc/bp1X7xK6ygDVddDZMNKZ0jAfBgNVHSMEGDAWgBSf
'' SIG '' pxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSg
'' SIG '' UqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
'' SIG '' b3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIw
'' SIG '' UENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBe
'' SIG '' MFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29m
'' SIG '' dC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRp
'' SIG '' bWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNV
'' SIG '' HRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
'' SIG '' MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOC
'' SIG '' AgEAwBPODpH8DSV07syobEPVUmOLnJUDWEdvQdzRiO2/
'' SIG '' taTFDyLB9+W6VflSzri0Pf7c1PUmSmFbNoBZ/bAp0DDf
'' SIG '' lHG1AbWI43ccRnRfbed17gqD9Z9vHmsQeRn1vMqdH/Y3
'' SIG '' kDXr7D/WlvAnN19FyclPdwvJrCv+RiMxZ3rc4/QaWrvS
'' SIG '' 5rhZQT8+jmlTutBFtYShCjNjbiECo5zC5FyboJvQkF5M
'' SIG '' 4J5EGe0QqCMp6nilFpC3tv2+6xP3tZ4lx9pWiyaY+2xm
'' SIG '' xrCCekiNsFrnm0d+6TS8ORm1sheNTiavl2ez12dqcF0F
'' SIG '' LY9jc3eEh8I8Q6zOq7AcuR+QVn/1vHDz95EmV22i6Qej
'' SIG '' Xpp8T8Co/+yaYYmHllHSmaBbpBxf7rWt2LmQMlPMIVqg
'' SIG '' zJjNRLRIRvKsNn+nYo64oBg2eCWOI6WWVy3S4lXPZqB9
'' SIG '' zMaOOwqLYBLVZpe86GBk2YbDjZIUHWpqWhrwpq7H1DYc
'' SIG '' csTyB57/muA6fH3NJt9VRzshxE2h2rpHu/5HP4/pcq06
'' SIG '' DIKpb/6uE+an+fsWrYEZNGRzL/+GZLfanqrKCWvYrg6g
'' SIG '' kMlfEWzqXBzwPzqqVR4aNTKjuFXLlW/ID7LSYacQC4Dz
'' SIG '' m2w5xQ+XPBYXmy/4Hl/Pfk5bdfhKmTlKI26WcsVE8zlc
'' SIG '' KxIeq9xsLxHerCPbDV68+FnEO40wggdxMIIFWaADAgEC
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
'' SIG '' wXEGahC0HVUzWLOhcGbyoYIC1jCCAj8CAQEwggEAoYHY
'' SIG '' pIHVMIHSMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
'' SIG '' aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
'' SIG '' ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQL
'' SIG '' EyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExp
'' SIG '' bWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjE3
'' SIG '' OUUtNEJCMC04MjQ2MSUwIwYDVQQDExxNaWNyb3NvZnQg
'' SIG '' VGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoD
'' SIG '' FQBt89HV8FfofFh/I/HzNjMlTl8hDKCBgzCBgKR+MHwx
'' SIG '' CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
'' SIG '' MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
'' SIG '' c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3
'' SIG '' DQEBBQUAAgUA6YCEjjAiGA8yMDI0MDIyMTIyMTc1MFoY
'' SIG '' DzIwMjQwMjIyMjIxNzUwWjB2MDwGCisGAQQBhFkKBAEx
'' SIG '' LjAsMAoCBQDpgISOAgEAMAkCAQACARECAf8wBwIBAAIC
'' SIG '' EWAwCgIFAOmB1g4CAQAwNgYKKwYBBAGEWQoEAjEoMCYw
'' SIG '' DAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQAC
'' SIG '' AwGGoDANBgkqhkiG9w0BAQUFAAOBgQB2/cIqLDPOHnt8
'' SIG '' QbgxJC+grtvmkjGfymZ9k9mM/raoM37cTAuQH7ZHc11P
'' SIG '' 2AlTcPMw2+v0wJxnPKuXT45in6STaNH82WjCnLxAxnmC
'' SIG '' Sf9irJROKSNO6XKK8wBmQZjhOHjLdOreuJ4S7c/1tP/X
'' SIG '' AvHeYYjnHaYvVTh9QSELjXznMTGCBA0wggQJAgEBMIGT
'' SIG '' MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
'' SIG '' dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
'' SIG '' aWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
'' SIG '' Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB
'' SIG '' 4NT8HxMVH35dAAEAAAHgMA0GCWCGSAFlAwQCAQUAoIIB
'' SIG '' SjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwLwYJ
'' SIG '' KoZIhvcNAQkEMSIEIGDm1FfSXZILr++0GQXqc+DPC9hc
'' SIG '' jTplEKsvT+AVwGu1MIH6BgsqhkiG9w0BCRACLzGB6jCB
'' SIG '' 5zCB5DCBvQQg4+5Sv/I55W04z73O+wwgkm+E2QKWPZyZ
'' SIG '' ucIbCv9pCsEwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzET
'' SIG '' MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
'' SIG '' bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
'' SIG '' aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
'' SIG '' cCBQQ0EgMjAxMAITMwAAAeDU/B8TFR9+XQABAAAB4DAi
'' SIG '' BCC/ZsCBpGJaKDIHmmfGZb+90sLFjaszn4ArkzMFIXK6
'' SIG '' 9jANBgkqhkiG9w0BAQsFAASCAgBwgRpQDvTkg5O5Pmnr
'' SIG '' NRrWjeH/m3+FcsYbBCKVRiividCShLTXQ1lnonUsN7/R
'' SIG '' 4VPaJkmBQ4F96slIdJbP0TXUHe4ait5VYXHIpNTrtwdJ
'' SIG '' 18xFScvZzHcstJ4hAWmr1tH79PlTHNoL4WPrsfYIDEt5
'' SIG '' 5RpLg3naeLz21Y3+F1xpo80My5dSV0r0bo1aPqJGh4Q8
'' SIG '' C72ykQBKJzpDnllXdm55WXFt9Gcq9tGq2nCNH9UtJlAc
'' SIG '' o0cP2vzzHVx/EGPsvLaB9xH5ZtHZZEm2A//cRbxRaAxj
'' SIG '' tBEn+/QrzCMumrQIs2BjuqNvyZvHTK3RD/Jia3kxEwlr
'' SIG '' ud/uxEi0N3tzjUKf0uFh4NbTi6OOawgiLe+g3KyA/l+L
'' SIG '' r7/7wBa/JRj8nZdkYYs8+/CDv59fx2egKwZLXOkHbQhs
'' SIG '' FQt1Az8/6dbNVXtiaw3JZq0Abmmq9gI9zG1AUbSC34hW
'' SIG '' yTtgC2W2ODkXOfBiARfaNzj1IY6yPfNs5CDysRk8a4dP
'' SIG '' djbu3hrkK6KQNm9Rc93tht1XcqCGWSyg3cRFUSQT18R+
'' SIG '' 3Qj5eCxLShPK1mB59QWbGQ8fGLted520GOKLNaNYpn8C
'' SIG '' vzvPqJ6U7iAsAx8V5Zv58jO7/2rRoinMM1yhbgAb7DNW
'' SIG '' /M/cw6Djpe61JQ7XtkN0zmuS7Wnuo3LInDOdjRxOCQJ/
'' SIG '' H1iTjw==
'' SIG '' End signature block
