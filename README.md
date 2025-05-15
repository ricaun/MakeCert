# MakeCert

Sample to create a certificate and use `signtool` to sign files.

## Location

The `MakeCert.exe`, `pvk2pfx.exe` and `signtool.exe` are located in the Windows SDK folder. For example:
```
C:\Program Files (x86)\Windows Kits\10\bin\x64\
```

## Local Files

The `MakeCert.exe`, `pvk2pfx.exe` and `signtool.exe` was copy from the [Windows Software Development Kit (SDK)](https://developer.microsoft.com/pt-br/windows/downloads/windows-10-sdk/) version `10.0.22621.0` / `x64`.

* [Download Local Files in Latest release](../../releases/latest)

## Creating a self-signed certificate authority (CA)

[MakeCert](https://learn.microsoft.com/en-us/windows-hardware/drivers/devtest/makecert) is a command-line to create certificate with privaty key (`.pvk`).

The command below create a certificate and a popup to require the password. Use `signfile` in this sample.

```
.\x64\MakeCert.exe -r -pe -ss CA -a sha256 -sky signature -eku 1.3.6.1.5.5.7.3.3 -sv signfile.pvk -n "CN=signfile" signfile.cer -b 01/01/2020 -e 12/31/2050
```

The OID [1.3.6.1.5.5.7.3.3](https://www.alvestrand.no/objectid/1.3.6.1.5.5.7.3.3.html) indicates that a certificate can be used for code signing. 

[Pvk2Pfx](https://learn.microsoft.com/en-us/windows-hardware/drivers/devtest/pvk2pfx) is a command-line tool that copies public key and private key information contained in `.pvk` file to a Personal Information Exchange (`.pfx`) file. 

The command below create the `.pfx` with the personal password (`-po`) `signfile`. The `-pi` specifies the password for the `.pvk` file.

```
.\x64\pvk2pfx.exe -pvk signfile.pvk -pi signfile -spc signfile.cer -pfx signfile.pfx -po signfile
```

## Sign file

[SignTool](https://learn.microsoft.com/en-us/windows-hardware/drivers/devtest/signtool) is a command-line tool that digitally-signs files, verifies signatures in files, and time stamps files.

The command below sign the file `ConsoleApp.exe` using the Personal Information Exchange (`.pfx`) using the password `signfile`.

```
.\x64\signtool.exe sign /f "signfile.pfx" /t http://timestamp.digicert.com /p "signfile" /fd sha1 "ConsoleApp.exe"
```

### Verify file

Verify the `ConsoleApp.exe` file is signed. `The file should show an error that the signature is not trusted by a certificate authority.`

```
.\x64\signtool.exe verify /pa "ConsoleApp.exe"
```

## Importing the CA certificate

To import the certificate use the `certutil` windows tool. `Popup shows to confirm the installation of the CA certificate`.

```
certutil -user -addstore Root signtool.cer
```

---

Do you like this project? Please [star this project on GitHub](../../stargazers)!
