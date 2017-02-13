
Cd C:\
cls

Add-Type -AssemblyName System.IO.Compression.FileSystem
Import-Module WebAdministration

New-Module -name MainDeclaration {
    $Path = "C:\"
 

    $projectPath= "https://github.com/TargetProcess/DevOpsTaskJunior"
    $gitUrl = "$($projectPath -replace '^https:\/\/' , 'https://codeload.')/zip/master"
    $gitRssUrl = "$projectPath/commits/master.atom"
   

    $shortUrl = "https://goo.gl/fu879a"
    $slackUri = "https://hooks.slack.com/services/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"


    $iisAppPoolName = "TestPool"
    $iisAppPoolDotNetVersion = "v4.0"
    $HostFilePath = "$env:windir\System32\drivers\etc\hosts"
    $ErrorMessage =""
    $commits = 7

    $logFile = "C:\log.txt"

#---------------------------------------------------------------------------------------------------------
function Get-NameFromUrl {

        Param (
            [Parameter(Mandatory=$true)]
                [String]$URL
    )
        $request = [System.Net.WebRequest]::Create($URL)    
        $response=$request.GetResponse()


        return $response.GetResponseHeader("Content-Disposition")
    }
#---------------------------------------------------------------------------------------------------------
function Unzip {

        param([string]$zipfile, [string]$outpath)

        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
    }
#---------------------------------------------------------------------------------------------------------
function Test-Connection(){

        param([string]$adress , [string] $Method )
        try {     
            if(!$Method){
            $Method = "Head" 
            }
           write-host $Method
           #write-host $adress
            if((Invoke-WebRequest $adress -Method $Method -DisableKeepAlive -TimeoutSec 1000  |
                select -ExpandProperty StatusCode) -eq 200){           
                return $TRUE
             }
        }
        catch {
            $global:ErrorMessage = $_.ErrorDetails.Message                
            return $FALSE
    }
    }
#---------------------------------------------------------------------------------------------------------
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
    $appPool | Set-ItemProperty -Name "managedPipelineMode" -Value "Integrated"
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
#---------------------------------------------------------------------------------------------------------
function SendToSlack(){
param([string] $URI,[object]$payload )


$objectToPayload = @{		
	"text" = $payload;	
}
    Invoke-WebRequest -URI $URI -Method Post -ContentType "application/json" -Body (ConvertTo-Json -Compress -InputObject $objectToPayload)
    }
#---------------------------------------------------------------------------------------------------------
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
#---------------------------------------------------------------------------------------------------------
function IsGitUpdated(){
    param([string]$url)

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
#---------------------------------------------------------------------------------------------------------
function DownloadProject(){
    param([string]$Url, [string] $FileName)
  
        #Write-Host "Downloading from github.."
        #Invoke-WebRequest $gitUrl -OutFile $FileName
        write-Host "Download complete.."
        return $TRUE
    }
#---------------------------------------------------------------------------------------------------------
function ToLog(){
param([string]$message,[bool]$noDatePrefix = $false, [string]$color= "Green")
    if($noDatePrefix){
  
        $message |
         %{write-host $_ -ForegroundColor Magenta; out-file -filepath $logFile -inputobject $_ -append}
    }
    else {
     
         "$(get-date -Format ‘HH:mm:ss’):" |
            %{Write-Host $_ -ForegroundColor $color -NoNewline; Write-Host $message;$_ = $_+ $($message);
             out-file -filepath $logFile -inputobject $_ -append}     
    }
}
#---------------------------------------------------------------------------------------------------------
function FixConfig(){
$configPath = "C:\DevOpsTaskJunior-master\Web.config"
$config = Get-Content $configPath
$config | %{
    $_ -replace '\.>' , '>' `
       -replace '4\.5\.2', '4.5'
} | Set-Content $configPath
}
#---------------------------------------------------------------------------------------------------------
function Stop(){
$timer.Stop()
Unregister-Event thetimer
}
#---------------------------------------------------------------------------------------------------------
function MainAction(){

$gitUrlTest = Test-Connection($gitUrl)
$needUpdate = IsGitUpdated($gitRssUrl)

If($gitUrlTest -and $needUpdate){

ToLog -message  $(get-date)  -noDatePrefix $TRUE 

ToLog "Zip path is $gitUrl"  
ToLog "Rss path is $gitRssUrl"
ToLog "Testing coonection to zip. Accessible? $gitUrlTest"
ToLog "Require to upload project? $needUpdate"

try{ 

    $FileName = Get-NameFromUrl($gitUrl) |% {($_ -split "=")[1]}    
    $BaseName = (Get-Item  $FileName).BaseName
    ToLog "$FileName unzip in folder $BaseName"  
    $isDownloaded = DownloadProject -RssUrl $gitRssUrl -gitUrl $Url -FileName $FileName
    
   }
catch{

    ToLog "Error occured at download time"
    $ErrorMessage = $_.ErrorDetails.Message
    ToLog $ErrorMessage
     #>>>?????
    }
  
   if($isDownloaded){
    ToLog "$FileName downloaded"
       if( Test-Path $Path$BaseName){
           ToLog "Folder $Path$BaseName exists. Removing.."
            Remove-Item -Path $Path$BaseName -Recurse
          }

    Unzip $Path$FileName $Path
    ToLog("Unzipping in $Path$BaseName")
    FixConfig
}


if((Get-WindowsFeature -Name web-server| select -ExpandProperty InstallState) -ne "Installed"){
    ToLog "Installing IIS and ASP.NET.."
    Install-WindowsFeature -Name Web-Server -includeManagementTools
    dism /online /enable-feature /all /featurename:IIS-ASPNET45
}

ToLog "Creating web-site and pool"
$oldPath = Get-Location
createWebSiteAndPool -iisAppPoolName $iisAppPoolName -iisAppPoolDoNetVersion $iisAppPoolDotNetVersion `
-iisAppName $BaseName -directoryPath $Path$BaseName 
cd $oldPath


if ((Test-Path $HostFilePath) -eq $TRUE -and !$(Get-Content -Path $HostFilePath| Select-String -pattern "127.0.0.1 $BaseName" -Quiet)){

ToLog "Changing host file.."
Write-output  "127.0.0.1 $BaseName"| Out-File $HostFilePath  -encoding ASCII -append
}

ToLog "Test new web-site.."


If((Test-Connection -adress $BaseName -Method "GET") -eq $TRUE){

    ToLog "Site $BaseName on $env:computername is working!"
    $result =  sendToSlack -URI $slackUri  -payload "Site $BaseName on $env:computername  is working!"  
}
else {

    ToLog "Site $BaseName responded with errors"
    ToLog $ErrorMessage
    ToLog "Sending message to Slack $slackUri"
    $result = sendToSlack -URI $slackUri  -payload $ErrorMessage     
    }

     if($result.StatusCode -eq 200){
        ToLog "Slack got Message"
    }

    ToLog $("-"*20) -noDatePrefix $TRUE
}


#write-host "After main IF"
$originalUrl = Get-RedirectedUrl($shortUrl)
    }

Export-ModuleMember -Function * -Variable *
} | Import-Module
#---------------------------------------------------------------------------------------------------------


MainAction

$action = {  
try{

    #throw "action error" 
    MainAction
}
catch{
    ToLog "Error occured.Stopping script"
    $timer.stop()       
    Unregister-Event thetimer

    ToLog $($Error) -color Red   
    ToLog "Script stopped. Sending message to Slack"
       
    $result =  sendToSlack -URI $slackUri  -payload "$Error"
  
    ToLog "Slack get Message" 

  
}
finally{ 
    
    Write-Host "IN FINALLY"
    Write-Host $Error
}
} 

$timer = New-Object System.Timers.Timer

#$event = Register-ObjectEvent -InputObject $timer -EventName elapsed -Action $action
$EventJob = Register-ObjectEvent -InputObject $timer -EventName elapsed –SourceIdentifier  thetimer -Action $action -OutVariable out


$timer.Interval = 15000
$timer.AutoReset = $true
$timer.start()



