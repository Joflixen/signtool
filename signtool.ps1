<#
 .Synopsis
  Sign Macro-Enabled Office File Module v1.2

 .Description
  This module performs the signing of office files by wrapping the Windows SDK Signing Tools 
  and the Microsoft Office Subject Interface Packages for Digitally Signing VBA Projects.

  Installation Instructions:

  1. Ensure that you have obtained the x86 Windows SDK Signing tools "signtool.exe" and place in this directory.
  Additionally you can register the path to this tool using the environment variable: $env:SIGNTOOL

  2. Ensure that you have obtained the Microsoft Office Subject Interface Packages (SIP) for Digitally Signing VBA Projects.
  Then extract the SIP Package and register the msosip.dll on the signing platform.

  Links to Resources:
    - https://developer.microsoft.com/en-us/windows/downloads/windows-10-sdk/
    - https://www.microsoft.com/en-us/download/confirmation.aspx?id=56617


 .Parameter Filename
  The macro-enabled office file to sign (relative path).


 .Parameter Certificate
  The PFX certificate to sign the macros with (see example below)


 .Parameter Passphrase
  The certificate passphrase used to unlock the certificate.


 .Parameter Author
  Written by James Mouat, 2020.07


 .Example
    # Sign a file.
    Add-Signature -Filename "workbook.xlsm" -Certificate "certificate.pfx" -Passphrase "-Passphr4se"

    # Check Signature.
    Get-Signature -Filename "document.docm"

    # Generate a Self Signed Certficiate with Powershell for testing.
    New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=SignTool Demo" -TextExtension @("2.5.29.19={text}false") -KeyUsage DigitalSignature -KeyLength 2048 -NotAfter (Get-Date).AddMonths(36)

#>
[System.IO.DirectoryInfo]$Script:path_signtool = [System.IO.Path]::GetFullPath( (join-Path -Path $PSScriptRoot "\signtool.exe") )

function Add-Signature {
    param(
        [String] $Filename,
        [String] $Certificate,
        [string] $Passphrase
    )

    # Use environmental setting for Sigtool if set
    if ( (Test-Path 'env:SIGNTOOL') ){
        [System.IO.DirectoryInfo]$Script:path_signtool = $env:SIGNTOOL
    }

    if (-not (Test-Path $Script:path_signtool) ){
        Write-Error "SignTool is missing!"
        return $false
    }
    #Write-Host $Script:path_signtool


    # Get the path to the Office file to work on
    [System.IO.DirectoryInfo]$Script:path_officefile  = [System.IO.Path]::GetFullPath( $Filename )

    [System.IO.DirectoryInfo]$Script:path_certificate = [System.IO.Path]::GetFullPath( $Certificate )

    # Get the path to the Office file to work on
    if (-not (Test-Path $Script:path_officefile) ){
        Write-Warning "The office file [ $Script:path_officefile ] does not exist!"
        return $false
    }

    if (-not (Test-Path $Script:path_certificate) ){
        Write-Warning "The certificate [ $Script:path_certificate ] cannot be found!"
        return $false
    }

    if ((Test-Path $Script:path_signtool) -and 
        (Test-Path $Script:path_officefile) -and 
        (Test-Path $Script:path_certificate) 
    ){
        #$command = "Sign /q /f '$Script:path_certificate' /p `"$Passphrase`" /fd `"SHA256`" /td `"SHA256`" `"$Script:path_officefile`" ";
        $command = "`"$Script:path_signtool`" Sign /q /f `"$Script:path_certificate`" /p `"$Passphrase`" /fd `"SHA256`" /td `"SHA256`" `"$Script:path_officefile`" ";
        #$command = '"{0}" Sign /q /f "{1}" /p "{2}" /fd "{3}" /td "{3}" "{4}" ' -f $Script:path_signtool, $Script:path_certificate, $Passphrase, 'SHA256', $Script:path_officefile;
        $success = $null
        $cmdOutput = cmd /c $command '2>&1' | ForEach-Object {
            if ($_ -is [System.Management.Automation.ErrorRecord]) {
                $success = $false
                Write-Error $_
            } else {
                if ($_ -like '*PFX password is not correct.'){
                    Write-Warning "Incorrect Password [ $Passphrase ]"
                    $success = $false
                } elseif ($_ -like 'SignTool Error:*'){
                    Write-Warning "Unable to Sign [ $Script:path_officefile ]"
                    $success = $false
                } elseif ($_ -like 'Error information:*'){
                    Write-Warning "It is likely that this file contains no Macros."
                    $success = $false
                } elseif ($_ -like '*Adding Additional Store'){
                    $success = $true
#                    Write-Host "Sucessfully signed file [ $Script:path_officefile ]"
#                    #return $true
                } else {
                    write-Host $_
                }
            }
        }
        return $success
    } else {
        write-error("General Failure!")
    }
    return $false
}


function Get-Signature {
    param(
        [String] $Filename
    )

    # Use environmental setting for Sigtool if set
    if ( (Test-Path env:SIGNTOOL) ){
        [System.IO.DirectoryInfo]$Script:path_signtool = $env:SIGNTOOL
    }

    # Get the path to the Office file to work on
    [System.IO.DirectoryInfo]$Script:path_officefile  = [System.IO.Path]::GetFullPath( (Join-Path -path $pwd $Filename) )

    if (-not (Test-Path $Script:path_officefile) ){
        Write-Warning "The office file [ $Script:path_officefile ] does not exist!"
        return $false
    }

    if ((Test-Path $Script:path_signtool) -and 
        (Test-Path $Script:path_officefile)
    ){
        $command = "verify `"$Script:path_officefile`" ";
        $cmdOutput = cmd /c "`"$Script:path_signtool`" $command" '2>&1' | ForEach-Object {
            write-host $_
        }
    } else {
        write-error("General Failure!")
    }
}

#Export-ModuleMember -Function Add-Signature
#Export-ModuleMember -Function Get-Signature

#Add-Signature -Filename .\Workbook.xlsm -Certificate CodeSigning.pfx -Passphrase "Passphrase"
Get-Signature -Filename .\Workbook.xlsm
