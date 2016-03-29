<#
.SYNOPSIS
SharePointMiner was created to address a need to programmtically search SharePoint sites for interesting documents during Red Team assessments.

Author: Andrew @ch33kyf3ll0w Bonstrom
License: BSD 3-Clause
Required Dependencies: Microsoft.SharePoint.Client.dll, Microsoft.SharePoint.Client.Runtime.dll
Optional Dependencies: None

.LINK
http://blog.vgrem.com/2014/08/24/working-with-folders-and-files-via-sharepoint-2013-rest-in-powershell/ -Awesome powershell examples for accessing sharepoint via REST api
http://www.slideshare.net/spcadriatics/developing-searchdriven-application-in-sharepoint-2013 - Great presentation about the REST API and using search query terms

#>
<#        
.DESCRIPTION  
Invoke-SPSearch uses the SharePoint REST API to issue search queries for files stored on a sharepoint site.

.PARAMETER SPUrl

The sharepoint site URL, ie. https://exist.sharepoint.com

.PARAMETER UserName

The sharepoint username, ie.  andrew.bonstrom@exist.sharepoint.com 

.PARAMETER Password  

The SharePoint user's password.

.PARAMETER QueryText

Text you wish to query the SharePoint for, ie. Title:ssn*   
   
.EXAMPLE  
PS C:\>Invoke-SPSearch -SPUrl https://exist.sharepoint.com -UserName andrew.bonstrom@exist.sharepoint.com -Password n0pa$$w0rd1 -QueryText "Title:ssn*" | Format-Table -AutoSize    
#>
Function Invoke-SPSearch(){
 
Param(
  [Parameter(Mandatory=$True)]
  [String]
  $SPUrl,
 
  [Parameter(Mandatory=$True)]
  [String]
  $UserName,
 
  [Parameter(Mandatory=$True)]
  [String]
  $Password, 
 
  [Parameter(Mandatory=$True)]
  [String]
  $QueryText,
  
  [Parameter(Mandatory=$True)]
  [String]
  $DLLFolderPath
 
)
   #Add the assembly type for the API access
   Add-Type -Path "$DLLFolderPath\Microsoft.SharePoint.Client.dll"
   Add-Type -Path "$DLLFolderPath\Microsoft.SharePoint.Client.Runtime.dll"
   
   #Builds out full query url and assigns to $Url
   $Url = $SPUrl + "/_api/search/query?querytext='" + $QueryText + "'"

 
   Write-Verbose "[+] Now issuing the following API query: $Url"

   try{
    #Sends query url, method, and credential data to REST helper function
    #Stops execution in the even of error
    $queryUrl = Invoke-RestSP -Url $Url -Method Get -UserName $UserName -Password $Password -ErrorAction Stop

   }
   catch{

    Write-Output "[-] Please review commandline parameters. Error: $_"

   }
   #Assigns the json section that contains our results to the var $results
   $results = $queryUrl.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.Cells.results

   Write-Verbose "[+] Now parsing results..."

   #Parses the data and places into a new PSObject with note properties containg results
   for ($i=0; $i -lt $results.length; $i++){

    $Out = New-Object PSObject 
    $Out | Add-Member Noteproperty 'Key' $results[$i].Key
    $Out | Add-Member Noteproperty 'Value' $results[$i].Value
    $Out 
  }
  Write-Verbose "[+] Results parsing complete! "


}

<#          
.DESCRIPTION  
 Get-SPFile downloads a file from a user provided URL.

.PARAMETER SPUrl

The sharepoint site URL, ie. https://exist.sharepoint.com

.PARAMETER UserName

The sharepoint username, ie.  andrew.bonstrom@exist.sharepoint.com 

.PARAMETER Password  

The SharePoint user's password.

.PARAMETER  RelativeFileUrl

Path to the file you wish to download, ie. /Shared Documents/Folder/File.docx

.PARAMETER  DownloadPath

Local directory you wish to place the downloaded file into.

.EXAMPLE  
 PS C:\> Get-SPFile -SPUrl  https://exist.sharepoint.com -UserName andrew.bonstrom@exist.sharepoint.com -Password n0pa$$w0rd1 -RelativeFileUrl '/Shared Documents/Folder/File To Download' -DownloadPath 'c:\downloads'     
#>
Function Get-SPFile(){
 
Param(
  [Parameter(Mandatory=$True)]
  [String]$SPUrl,
 
  [Parameter(Mandatory=$True)]
  [String]$UserName,
 
  [Parameter(Mandatory=$True)]
  [String]$Password, 
 
  [Parameter(Mandatory=$True)]
  [String]$RelativeFileUrl,
 
  [Parameter(Mandatory=$True)]
  [String]$DownloadPath
 
  [Parameter(Mandatory=$True)]
  [String]
  $DLLFolderPath
 
)
   #Add the assembly type for the API access
   Add-Type -Path "$DLLFolderPath\Microsoft.SharePoint.Client.dll"
   Add-Type -Path "$DLLFolderPath\Microsoft.SharePoint.Client.Runtime.dll"
  
   $Url = $SPUrl + "/_api/web/GetFileByServerRelativeUrl('" + $RelativeFileUrl + "')"

   Write-Verbose "[+] Now issuing the following API query: $Url"

   try{
   #Issues query to REST API helper function and returns binary string to variable $filecontent
   $fileContent = Invoke-RestSP -Url $Url -Method Get -UserName $UserName -Password $Password -BinaryStringResponseBody $True -ErrorAction Stop

   }
   catch{
        Write-Output "[-] Please review commandline parameters. Error: $_"
   }

   Write-Verbose "[+] Query results retrieved. Now attempting to write file out to directory: $DownloadPath"

   try{

   #Write out file to speciifed directory
   $fileName = [System.IO.Path]::GetFileName($FileUrl)
   $downloadFilePath = [System.IO.Path]::Combine($DownloadPath,$fileName)
   [System.IO.File]::WriteAllBytes($downloadFilePath,$fileContent)

   }

   catch{
        Write-Output "[-] Please review commandline parameters. Error: $_"
   }
   Write-Verbose "[+] File located at $FileUrl successfully written to $DownloadPath!"
}

<#
.DESCRIPTION  
Get-SPFolderContentList uses the SharePoint REST API to issue list out the contents of a user provided folder.

.PARAMETER SPUrl

The sharepoint site URL, ie. https://exist.sharepoint.com

.PARAMETER UserName

The sharepoint username, ie.  andrew.bonstrom@exist.sharepoint.com 

.PARAMETER Password  

The SharePoint user's password.

.PARAMETER RelativFoldereUrl

Relative Url of the folder you wish to list the contents of.  
   
.EXAMPLE  
PS C:\>Get-SPFolderContentList -SPUrl  https://exist.sharepoint.com -UserName andrew.bonstrom@exist.sharepoint.com -Password n0pa$$w0rd -RelativeFolderUrl "/Shared Documents/SocialSecurityNumbers"| format-table -AutoSize      
#>
Function Get-SPFolderContentList(){
 
Param(
  [Parameter(Mandatory=$True)]
  [String]$SPUrl,
 
  [Parameter(Mandatory=$True)]
  [String]$UserName,
 
  [Parameter(Mandatory=$False)]
  [String]$Password, 
 
  [Parameter(Mandatory=$True)]
  [String]$RelativeFolderUrl
 
  [Parameter(Mandatory=$True)]
  [String]
  $DLLFolderPath
 
)
   #Add the assembly type for the API access
   Add-Type -Path "$DLLFolderPath\Microsoft.SharePoint.Client.dll"
   Add-Type -Path "$DLLFolderPath\Microsoft.SharePoint.Client.Runtime.dll"
 
   #Create URL for query
   $Url = $SPUrl + "/_api/web/GetFolderByServerRelativeUrl('" + $RelativeFolderUrl + "')/Files"

   try{

   $folderContent = Invoke-RestSP $Url Get $UserName $Password
    
   }
   catch{

    Write-Output "[-] Please review commandline parameters. Error: $_"

   }

   #Assigns the names of all files to $fIleNames to be used as an iterator
   $fileNames = $folderContent.results.Name

   try{

   for ($i=0; $i -lt $fileNames.length; $i++){

     #Creates a new PSObject to organize relevant data from request
    $Out = New-Object PSObject
    $Out | Add-Member Noteproperty 'FileName' $fileNames[$i]
    $Out | Add-Member Noteproperty 'RelativeURL' $folderContent.results.ServerRelativeUrl[$i]
    $Out     

   }
 }

 catch{

    Write-Output "[-] Please review commandline parameters. Error: $_"

 }
}

#SharePointRest API Helper function 
Function Invoke-RestSP(){
 
Param(
[Parameter(Mandatory=$True)]
[String]$Url,
 
[Parameter(Mandatory=$False)]
[Microsoft.PowerShell.Commands.WebRequestMethod]$Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Get,
 
[Parameter(Mandatory=$True)]
[String]$UserName,
 
[Parameter(Mandatory=$False)]
[String]$Password,

[Parameter(Mandatory=$False)]
[System.Byte[]]$Body,

[Parameter(Mandatory=$False)]
[Boolean]$BinaryStringResponseBody = $False

)
 

    #Credential check
   if([string]::IsNullOrEmpty($Password)) {
      $SecurePassword = Read-Host -Prompt "Enter the password" -AsSecureString 
   }
   else {
      $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
   }
 
    #Establish new sharepoint credential object
   $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
   $request = [System.Net.WebRequest]::Create($Url)
   $request.Credentials = $credentials
   $request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
   $request.ContentType = "application/json;odata=verbose"
   $request.Accept = "application/json;odata=verbose"
   $request.Method=$Method
   $request.ContentLength = 0


   #Porcess query response
   $response = $request.GetResponse()
   try {
       if($BinaryStringResponseBody -eq $False) {    
           $streamReader = New-Object System.IO.StreamReader $response.GetResponseStream()
           try {
              $data=$streamReader.ReadToEnd()
              $results = $data | ConvertFrom-Json
              return $results.d 
           }
           finally {
              $streamReader.Dispose()
           }
        }
        else {
           $dataStream = New-Object System.IO.MemoryStream
           try {
           Stream-CopyTo -Source $response.GetResponseStream() -Destination $dataStream
           $dataStream.ToArray()
           }
           finally {
              $dataStream.Dispose()
           } 
        }
    }
    finally {
        $response.Dispose()
    }
   
}
 
 
# Get Context Info 
Function Get-SPOContextInfo(){
 
Param(
[Parameter(Mandatory=$True)]
[String]$WebUrl,
 
[Parameter(Mandatory=$True)]
[String]$UserName,
 
[Parameter(Mandatory=$False)]
[String]$Password
)
 
   
   $Url = $WebUrl + "/_api/contextinfo"
   Invoke-RestSPO $Url Post $UserName $Password
}
 


Function Stream-CopyTo([System.IO.Stream]$Source, [System.IO.Stream]$Destination)
{
    $buffer = New-Object Byte[] 8192 
    $bytesRead = 0
    while (($bytesRead = $Source.Read($buffer, 0, $buffer.Length)) -gt 0)
    {
         $Destination.Write($buffer, 0, $bytesRead)
    }
}

