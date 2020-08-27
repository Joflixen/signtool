# SignTool for Macro Enabled Office Files

Signing Visual Basic Macros in Office Files

## What is this?
This module wraps the **Windows 10 SDK** code signing tool **[SignTool.exe]** in conjunction with the **Microsoft Office Subject Interface Packages for Digitally Signing VBA Projects** to perform the Digital Signing of Macro-Enabled Microsoft Office files.

## Installation Instructions
1. Download the [Windows SDK](https://go.microsoft.com/fwlink/p/?linkid=84091), and install the Code Signing component.
2. Once installed go to the install path, on my local workstation it was here: 
`C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x86` 
3. Copy the x86 version of **SignTool.exe** to your working path: 
`C:\[workingpath]\SignTool.exe` 
4. Download and run the [Microsoft Office Subject Interface Package](https://www.microsoft.com/en-us/download/details.aspx?id=56617), during the installer it will ask for an installation path, choose a path like: 
`C:\[workingpath]\MSOSIP` 
5. Refer to the readme.txt for additional instructions if required in "**HOW TO USE THESE COMPONENTS**"

#### HOW TO USE THESE COMPONENTS
Register the corresponding MSO SIP dll for the format you are working with:
- For VBA projects contained in legacy Office file formats use:
`C:\[workingpath]\MSOSIP>regsvr32.exe msosip.dll`

- For VBA projects contained in OOXML Office file formats use:
`C:\[workingpath]\MSOSIP>regsvr32.exe msosipx.dll`

## External Links
- https://developer.microsoft.com/en-us/windows/downloads/windows-10-sdk/
- https://www.microsoft.com/en-us/download/confirmation.aspx?id=56617


## Examples

#### Sign a file.
    `C:\PowerShell> Add-Signature -Filename "workbook.xlsm" -Certificate "certificate.pfx" -Passphrase "-Passphr4se"`

#### Check Signature.
    `C:\PowerShell> Get-Signature -Filename "document.docm"`

#### Generate a Self Signed Certficiate
    `C:\Powershell> New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=SignTool Demo" -TextExtension @("2.5.29.19={text}false") -KeyUsage DigitalSignature -KeyLength 2048 -NotAfter (Get-Date).AddMonths(36)`
