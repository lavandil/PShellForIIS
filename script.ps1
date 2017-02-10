
Cd C:\

cls
Add-Type -AssemblyName System.IO.Compression.FileSystem

Import-Module WebAdministration

$Path = "C:\"
$gitUrl = "https://codeload.github.com/TargetProcess/DevOpsTaskJunior/zip/master"
$gitRssUrl = "https://github.com/TargetProcess/DevOpsTaskJunior/commits/master.atom" 

$shortUrl = "https://goo.gl/fu879a"
$slackUri = "https://hooks.slack.com/services/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"


$iisAppPoolName = "TestPool"
$iisAppPoolDotNetVersion = "v4.0"
$HostFilePath = "$env:windir\System32\drivers\etc\hosts"
$global:ErrorMessage =""
$global:commits = 7

function Get-NameFromUrl {

    Param (
        [Parameter(Mandatory=$true)]
            [String]$URL
)
    $request = [System.Net.WebRequest]::Create($URL)    
    $response=$request.GetResponse()


    return $response.GetResponseHeader("Content-Disposition")
}

function Unzip {
    param([string]$zipfile, [string]$outpath)

    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}
function Test-Connection(){
    param([string]$adress , [string] $Method )
    try {
        
        if(!$Method){
        $Method = "Head" 
        }
      
        if((Invoke-WebRequest $adress -Method $Method |select -ExpandProperty StatusCode) -eq 200){           
            return $TRUE
         }
    }
    catch {
        $global:ErrorMessage = $_.ErrorDetails.Message                
        return $FALSE
}
}
function CreateWebSiteAndPool(){

param(
[string] $iisAppPoolName,
[string] $iisAppPoolDoNetVersion,
[string] $iisAppName,
[string] $directoryPath
)


#navigate to the app pools root
cd IIS:\AppPools\

#check if the app pool exists
if (!(Test-Path $iisAppPoolName -pathType container))
{
    #create the app pool
    $appPool = New-Item $iisAppPoolName
    $appPool | Set-ItemProperty -Name "managedRuntimeVersion" -Value $iisAppPoolDotNetVersion
}

#navigate to the sites root
cd IIS:\Sites\

#check if the site exists
if (Test-Path $iisAppName -pathType container)
{
    return
}

#create the site
$iisApp = New-Item $iisAppName -bindings @{protocol="http";bindingInformation=":80:" + $iisAppName} -physicalPath $directoryPath
$iisApp | Set-ItemProperty -Name "applicationPool" -Value $iisAppPoolName
}
function SendToSlack(){
param([string] $URI,[object]$payload )

$objectToPayload = @{		
	"text" = $payload;	
}
    Invoke-WebRequest -URI $URI -Method Post -ContentType "application/json" -Body (ConvertTo-Json -Compress -InputObject $objectToPayload)
}
function Get-RedirectedUrl {

    Param (
        [Parameter(Mandatory=$true)]
        [String]$URL
    )

    $request = [System.Net.WebRequest]::Create($url)
    $request.AllowAutoRedirect=$false
    $response=$request.GetResponse()

  If ($response.StatusCode -eq "MovedPermanently")
    {
        $response.GetResponseHeader("Location")
    }
}
function IsGitUpdated(){
param([string]$url)

Write-host "IS Git Updated function"
    $response = Invoke-WebRequest -Uri $url
    $doc = [xml]$response.Content
 
    if($doc.feed.entry.count -gt $global:commits){
    $global:commits = $doc.feed.entry.count
    return $FALSE
    }
    else {
    return $TRUE
    }
}
function DownloadProject(){
param([string]$Url, [string] $FileName)
  
    #Write-Host "Downloading from github.."
    #Invoke-WebRequest $gitUrl -OutFile $FileName
    write-Host "Download complete.."
    return $TRUE
}

function CheckSiteWorking(){

    
}

function MainAction(){

Write-host "MAIN ACTION!"
Write-host (Test-Connection($gitUrl))
Write-host (IsGitUpdated($gitRssUrl))

If((Test-Connection($gitUrl)) -and (IsGitUpdated($gitRssUrl))){

Write-host "Before try"
try{
   
Write-host "try block"

    $FileName = Get-NameFromUrl($gitUrl) |% {($_ -split "=")[1]}    
    $BaseName = (Get-Item  $FileName).BaseName


   $isDoanloaded = DownloadProject -RssUrl $gitRssUrl -gitUrl $Url -FileName $FileName
    
   }
catch
    {
    write-host "first catch"
   $ErrorMessage = $_.ErrorDetails.Message
    }

   
   if(isDownloaded){
    If( Test-Path $Path$BaseName){
        Write-Host "removing folder"
        Remove-Item -Path $Path$BaseName -Recurse

    }

    Unzip $Path$FileName $Path
}

if((Get-WindowsFeature -Name web-server| select -ExpandProperty InstallState) -ne "Installed"){
    Write-Output "Installing IIS and ASP.NET.."
    Install-WindowsFeature -Name Web-Server -includeManagementTools
    dism /online /enable-feature /all /featurename:IIS-ASPNET45
}


$oldPath = Get-Location
createWebSiteAndPool -iisAppPoolName $iisAppPoolName -iisAppPoolDoNetVersion $iisAppPoolDotNetVersion `
-iisAppName $BaseName -directoryPath $Path$BaseName 
cd $oldPath


if ((Test-Path $HostFilePath) -eq $TRUE){
Write-Output "Changing host file.."
Write-output  "127.0.0.1 $BaseName"| Out-File $HostFilePath  -encoding ASCII -append
}

Write-Output "Test new web-site.."


If((Test-Connection -adress $BaseName -Method "GET") -eq $TRUE){
    $result =  sendToSlack -URI $slackUri  -payload "Site is working!"  
}
else {
  Write-Output NO!
 
   $result = sendToSlack -URI $slackUri  -payload $ErrorMessage  
    }
}


write-host "After main IF"
$originalUrl = Get-RedirectedUrl($shortUrl)
}

MainAction

$action = {  
try{
    Write-Host "Action execution"
    #throw "action error"

    #DownloadProject -gitRssUrl $gitUrl -gitUrl $Url -FileName $FileName

    MainAction
}
catch{
    Write-Host $Error.Gettype()
    Write-Host $Error.Count
    Write-Host $Error
    #to stop run 
    $timer.stop() 
    #cleanup 
    Unregister-Event thetimer
}
} 

$timer = New-Object System.Timers.Timer

#$event = Register-ObjectEvent -InputObject $timer -EventName elapsed -Action $action
Register-ObjectEvent -InputObject $timer -EventName elapsed –SourceIdentifier  thetimer -Action $action -OutVariable out



$timer.Interval = 5000
$timer.AutoReset = $true
$timer.start()


