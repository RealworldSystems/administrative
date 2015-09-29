<#
Copyright (c) 2015, Realworld Systems
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this
  list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice,
  this list of conditions and the following disclaimer in the documentation
  and/or other materials provided with the distribution.

* Neither the name of administrative nor the names of its
  contributors may be used to endorse or promote products derived from
  this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#>

param (
    [string]$AdminUrl,
    [string]$SMTPServer,
    [string]$From,
    [string]$CredentialUsername,
    [string]$CredentialFile
)

if ([bool]$AdminUrl) { $env:AdminUrl = $AdminUrl }
if ([bool]$SMTPServer) { $env:SMTPServer = $SMTPServer }
if ([bool]$From) { $env:From = $From }
if ([bool]$CredentialUsername) { $env:CredentialUsername = $CredentialUsername }
if ([bool]$CredentialFile) { $env:CredentialFile = $CredentialFile }


Import-Module "$PSScriptRoot\SharepointOnlineEx"


function Halt($why) {
    Write-Host $why
    if ($host.name -eq 'ConsoleHost' -and [bool]([Environment]::GetCommandLineArgs() -like '-NonInteractive')) {
        Exit 1
    }
    throw ""
}

try {
    if (![bool]$env:AdminUrl) {
        Halt "No Administrative URL set, please use argument -AdminUrl or environment variable AdminUrl"
    }

    if (![bool]$env:SMTPServer) {
        Halt "No SMTP ServerURL set, please use argument -SMTPServer or environment variable SMTPServer"
    }

    if (![bool]$env:From) {
        Halt "No FROM address set, please use arugument -From or environment variable From"
    }


    try {
        $cred = Get-SPOCredential
    } catch [Exception] {
        Halt "Can't login properly"
    }

    Write-Host "Retrieve all checked out files"

    $files = Get-AllCheckoutFiles -AdminUrl $env:AdminUrl -Credential $cred
    $emails = $files |% { $_.Email } | select -Unique
    # Scavenge emails from $files

    $emails |% {

        $forEmail = @()
        foreach ($file in $files) {
            if($file.Email -eq $_) {
                $forEmail += $file
            }
        }

        Write-Host "[MAIL]: Sending email to $_"

        $fileTable = Format-CheckoutFilesAsTable $forEmail | select * -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray | ConvertTo-HTML
        $html = @"
<head><title>Checked out Files</title></head><body>Please check in the following files at your convenience: <br/>$fileTable</body>
"@
        Send-MailMessage -From $env:From -To $_ -SmtpServer $env:SMTPServer -Subject "Checked out files" -BodyAsHtml $html
    }

} catch [Exception] {Write-Host $_.Exception.Message }