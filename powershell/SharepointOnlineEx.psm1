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

[void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Sharepoint.Client")

function Get-SPOCredential () {
   try {
        if ([bool]$env:CredentialFile -and [bool]$env:CredentialUsername) {
            $pwd = cat $env:CredentialFile | ConvertTo-SecureString
            $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $env:CredentialUsername, $pwd
            return $cred
        } else {
            $cred = Get-Credential
            return $cred
        }
    } catch [Exception] {
        throw "No credentials supplied, bailing out."
    } finally {
        if (![bool]$cred) {
            throw "No credentials supplied, bailing out."
        }
    }
}

function Get-Lists () {
    param (
        $Url = $(throw "Provide a non-administrative sharepoint Url, e.g.: https://....sharepoint.com/sites/..."),
        $Credential = $(throw "Provide a PSCredential object")
    )

    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $spoCred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName, $Credential.Password)
    $ctx.Credentials = $spoCred

    $ctx.Load($ctx.Web)
    $lists = $ctx.Web.Lists
    $ctx.Load($lists)
    try {
        $ctx.ExecuteQuery()

        $arr = @()
        $lists |% { 
            $title = $_.Title
            Write-Host "[SITE]: $title --> {$Url}"
            $arr += $title
        }
        return $arr
    } catch [Exception] {
        return @()
    }
}

function Get-CheckoutFiles () {
    param (
        $Url = $(throw "Provide a non-administrative sharepoint Url, e.g.: https://....sharepoint.com/sites/..."),
        $Credential = $(throw "Provide a PSCredential object"),
        $ListTitle = $(throw "Provide the title of the library")
    )

    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $spoCred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName, $Credential.Password)
    $ctx.Credentials = $spoCred

    $ctx.Load($ctx.Web)
    $ctx.Load($ctx.Web.Webs)
    $list = $ctx.Web.Lists.GetByTitle($ListTitle)

    $table = @()

    Write-Host "[LIST]: Scanning for checkout out documents in $ListTitle --> $Url"

    $query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(1000000, "FileLeafRef", "Title", "CheckoutUser")
    $results = $list.GetItems($query)
    $ctx.load($list)
    $ctx.Load($results)
    try {
        $ctx.ExecuteQuery()
        $results |%  {
            $userName = $_["CheckoutUser"]    
            if ($userName) {
                $fileRef = $_["FileRef"]
                $ru = $userName.LookupValue
                $email = $userName.Email
                $table += @{Document=$fileRef; User=$ru; Email=$email}
            }
    
        }
    } catch [Exception] {}
    return $table
}


function Get-Sites() {
    param (
        $AdminUrl = $(throw "Provide an administrative URL http://...-admin.sharepoint.com"),
        $Credential = $(throw "Provide a PSCredential object")
    )

    function NextSites($Url) {
        $arr = @()
        $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
        $spoCred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName, $Credential.Password)
        $ctx.Credentials = $spoCred
        $web = $ctx.Web
        $ctx.Load($web)
        $ctx.Load($web.Webs)
        try {
            $ctx.ExecuteQuery()

            foreach ($w in $web.Webs) {
                $url = $w.Url
                Write-Host "[SCAN] $url"
                $arr += $url
                NextSites $url |% {$arr += $_ }
            }
        } catch [Exception] {Write-Host "Could not access this site ${web.Url}"} # swallow all exceptions
        return $arr
    }

    Connect-SPOService -Url $AdminUrl -Credential $Credential
    $sitesInfo = Get-SPOSite
    $arr = @()
    $sitesInfo |% {
        $url = $_.Url
        Write-Host "[SCAN] $url"
        $arr += $url
        $sites = NextSites $url
        foreach($site in $sites) {
            $arr += $site
        }
    }
    return $arr
}

# Retrieve all site collections known to the sharepoint online instance


function Get-AllCheckoutFiles () {
    param (
        $AdminUrl = $(throw "The Administrative URL"),
        $Credential = $(throw "A PSCredential")
    )
    $sites = Get-Sites -AdminUrl $AdminUrl -Credential $Credential

    $charr = @()

    foreach ($site in $sites) {
        Get-Lists -Url $site -Credential $Credential |% {
            $res = Get-CheckoutFiles -Url $site -ListTitle $_ -Credential $Credential
            foreach($r in $res) {
                $charr += $r
            }
        }
    }

    return $charr

}

function Format-CheckoutFilesAsTable ($files) {

    $table = New-Object system.Data.DataTable "Checked out documents"

    $col1 = New-Object system.Data.DataColumn "Document location",([string])
    $col2 = New-Object system.Data.DataColumn "Checkout lock held by",([string])
    $col3 = New-Object system.Data.DataColumn "Email",([string])

    $table.columns.Add($col1)
    $table.columns.Add($col2)
    $table.columns.Add($col3)

    $files |% {
        if (![string]::IsNullOrEmpty($_.Document)) {
            $row = $table.NewRow()
            $row[0] = $_.Document
            $row[1] = $_.User 
            $row[2] = $_.Email
            $table.Rows.Add($row)
        }
    }

    return $table
}

