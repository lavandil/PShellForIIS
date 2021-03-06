﻿Cd C:\
cls

Add-Type -AssemblyName System.IO.Compression.FileSystem

New-Module -name MainDeclaration {
    $Path = "C:\"

    $projectPath= "https://github.com/TargetProcess/DevOpsTaskJunior"
    $gitUrl = "$($projectPath -replace '^https:\/\/' , 'https://codeload.')/zip/master"
    $gitRssUrl = "$projectPath/commits/master.atom"

    $shortUrl = "https://goo.gl/fu879a"
    #$shortUrl = "https://goo.gl/LWjuda"

    $iisAppPoolName = "TestPool"
    $iisAppPoolDotNetVersion = "v4.0"
    $HostFilePath = "$env:windir\System32\drivers\etc\hosts"
    $script:ErrorMessage =""
    $script:commits = 0

 
    $logFile = "C:\log.txt"
    $EventExecuting = $TRUE
    $debug = $FALSE
    $consoleOutput = $TRUE
    $timerInterval = 30000
    $firstRun = $TRUE
    
    $eventQuery = @" 
     Select * From __InstanceOperationEvent Within 1 
     Where TargetInstance Isa 'Win32_Service' 
     and 
     TargetInstance.Name='w3svc'
     and
     TargetInstance.State="Stopped"
"@

    $featureQuery = @" 
     Select * From __InstanceModificationEvent Within 5 where TargetInstance Isa 'Win32_OptionalFeature'
     and
     TargetInstance.InstallState != '1'
     GROUP Within 5
"@

#---------------------------------------------------------------------------------------------------------
function Get-NameFromUrl {

        Param (
            [Parameter(Mandatory=$true)]
                [String]$URL
    )
        $request = [System.Net.WebRequest]::Create($URL)    
        $response=$request.GetResponse()
        $result = $response.GetResponseHeader("Content-Disposition")|% {($_ -split "=")[1]} 
        $response.Dispose()
        return $result
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
           #debug $Method
           #debug "Adress $adress"
          
           $httpResponse = Invoke-WebRequest $adress -Method $Method            
           
           $httpResponse.BaseResponse.Close()
            if( ($httpResponse |select -ExpandProperty StatusCode) -eq 200){
                       
                return $TRUE
             }
             
        }
        catch {
            $script:ErrorMessage = $_   
            debug  "in Test-Connection $_ with url $adress"  
            debug $_          
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
$oldPath = Get-Location
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
    cd $oldPath
    return
    
}

#create the site
$iisApp = New-Item $iisAppName -bindings @{protocol="http";bindingInformation=":80:" + $iisAppName} -physicalPath $directoryPath
$iisApp | Set-ItemProperty -Name "applicationPool" -Value $iisAppPoolName
cd $oldPath
}
#---------------------------------------------------------------------------------------------------------
function SendToSlack(){
param(
[string] $URI,
[Parameter(ValueFromPipeline)]
[object]$payload )
try{
debug $payload
$objectToPayload = @{
    "username" = "$BaseName";
    "icon_emoji" = ":necktie:";
	"text" = $payload;
}
    $result = Invoke-WebRequest -URI $URI -Method Post -ContentType "application/json" -Body (ConvertTo-Json -Compress -InputObject $objectToPayload)

    if($result.StatusCode -eq 200){

        ToLog "Slack got message"
        return $TRUE
    }
}
catch{
    ToLog "Slack didn't get message" -color Red
    ToLog $_
    return $FALSE
}
}
#$x = 123 |tolog -PassThru | sendtoslack -URI "$slackUri"
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
            $response.Dispose()
        }
    }
$slackUri = Get-RedirectedUrl($shortUrl) 
#---------------------------------------------------------------------------------------------------------
function IsGitUpdated(){
    param([string]$url)

        $response = Invoke-WebRequest -Uri $url
        $doc = [xml]$response.Content
 
        if($doc.feed.entry.count -gt $commits){
        $script:commits = $doc.feed.entry.count
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
        $response = Invoke-WebRequest $gitUrl -OutFile $FileName
        #write-Host "Download complete.."
        return $TRUE
    }
#---------------------------------------------------------------------------------------------------------
function ToLog(){
param(
[Parameter(ValueFromPipeline)]
[string]$message,
[bool]$noDatePrefix = $false,
[string]$color= "Green",
[switch]$PassThru
)
 
    if($noDatePrefix){

        $message |
         %{if($consoleOutput){write-host $_ -ForegroundColor Magenta}; out-file -filepath $logFile -inputobject $_ -append}
    }
    else {     
         "$(get-date -Format ‘HH:mm:ss’):" |
            %{if($consoleOutput){Write-Host $_ -ForegroundColor $color -NoNewline; Write-Host $message};$_ = $_+ $($message);
             out-file -filepath $logFile -inputobject $_ -append}     
    }
    if($PassThru){
        return $message
    }
}
#---------------------------------------------------------------------------------------------------------
function Debug(){
param([string]$message)
if($debug){
write-host "debug:$message" -ForegroundColor Yellow
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
function StopAll(){
debug InStopALL
$script:timer.Stop()
"timerEvent", "processEvent", "featureEvent" |%{Unregister-Event $_ -ErrorAction SilentlyContinue}
}
#---------------------------------------------------------------------------------------------------------
function StartAll(){
debug StartAll
<#
if($script:timer){
write-host "Timer already exists.Removing.."
stopall
}#>
$script:timer = New-Object System.Timers.Timer
$EventJob = Register-ObjectEvent -InputObject $timer -EventName elapsed –SourceIdentifier  timerEvent -Action {timerhandler}
$timer.Interval = $timerInterval
$timer.AutoReset = $true
$timer.start()

$eventJob2 = Register-WmiEvent -Query $eventQuery -Action {ProcessStopHandler} -SourceIdentifier processEvent
$eventJob3 = Register-WmiEvent -Query $featureQuery -Action {FeatureRemoveHandler($Event)} -SourceIdentifier featureEvent
}
#---------------------------------------------------------------------------------------------------------
function StopIIS(){
  #Stop-Service iisadmin,was,w3svc 
  Stop-Service iisadmin,was,w3svc 
  #Stop-Service w3svc
}
#---------------------------------------------------------------------------------------------------------
function StartIIS(){
  Start-Service was,w3svc
}
#---------------------------------------------------------------------------------------------------------
function StopPool(){
param([string]$poolName)
try{
 debug "StopPool"
 $state= (Get-WebAppPoolState -Name $poolName).Value

     if($state -ne "Stopped"){
      Stop-WebAppPool  -Name $poolName
     }


 while((Get-WebAppPoolState -Name $poolName).Value -ne "Stopped"){
    Start-Sleep -s 2
    ToLog "Stopping $poolName, state: $((Get-WebAppPoolState -Name $poolName).Value)"    
 }
 }
 catch{
    
    ToLog  $_
 }
}
#---------------------------------------------------------------------------------------------------------
function StartPool(){
param([string]$poolName)
try{
 $state= (Get-WebAppPoolState -Name $poolName).Value

     if($state -eq "Stopped"){
      Start-WebAppPool  -Name $poolName
     } 

 while((Get-WebAppPoolState -Name $poolName).Value -ne "Started"){
    Start-Sleep -Seconds 2
    ToLog "Starting $poolName, state: $((Get-WebAppPoolState -Name $poolName).Value)"
 }
 }
 catch{
    ToLog $_
 }
}
#---------------------------------------------------------------------------------------------------------
function TestSite(){
param([string]$name)

debug "TestSite"
#start-sleep -Seconds 10
If((Test-Connection -adress $name -Method "GET") -eq $TRUE){

    ToLog "Site $name on $env:computername is working!"
    $result =  sendToSlack -URI $slackUri  -payload "Site $BaseName on $env:computername is working!" 
  
}
else {

    ToLog "Site $name responded with errors" -color Red 
    
    ToLog $ErrorMessage -color Red  #remove from testconnection tolog
    ToLog "Sending message to Slack $slackUri"
    $result = sendToSlack -URI $slackUri  -payload $ErrorMessage   
    debug "123"
    $script:ErrorMessage =""  
    }

    ToLog $("-"*20) -noDatePrefix $TRUE
 
}
#---------------------------------------------------------------------------------------------------------
function TimerHandler(){
try{
    debug "TimerHandler"
    #throw "action error" 
    
    if(!$eventExecuting){
    MainAction
    }
}
catch{
    ToLog "Error occured.Stopping script" -color Red 
    #stopALL
    $timer.Stop()

    ToLog $($Error) -color Red   
    ToLog "Script stopped. Sending message to Slack"
       
    $result =  sendToSlack -URI $slackUri  -payload "$Error"  
    
    debug $_.Exception
  
}
finally{ 
    
    debug "IN FINALLY"
    #Write-Host $Error

}
}
#---------------------------------------------------------------------------------------------------------
function ProcessStopHandler(){
$timer.Stop()
#write-host action-2
$message = "IIS stopped. Trying to restart. " 
ToLog $message -color Red 
$SlackPayload = $message

    if(((Get-WindowsFeature -Name web-server| select -ExpandProperty InstallState) -ne "Installed") -or ((Get-WindowsFeature Web-Asp-Net45|select -ExpandProperty InstallState) -ne "Installed")){
      
      $message = "IIS was uninstalled. Trying to reload project" 
      ToLog $message -color Red 
      $slackPayload+= $message     
      $commits = 0
      MainAction    
    }
    else{

    startIIS
    #Start-Sleep -Seconds 5
        if((Get-Service w3svc).Status -eq "Running"){
            
            $message = "IIS working. Testing $BaseName"
            ToLog $message
            $SlackPayload+=$message
            $result =  sendToSlack -URI $slackUri  -payload $slackPayload 
            TestSite $BaseName
        }
        else{
         $message ="IIS couldn't be restarted on $env:computername"
         ToLog $message -color Red 
         $SlackPayload += $message
                
         $result =  sendToSlack -URI $slackUri  -payload $slackPayload 
        }
    }
    $timer.Start()       

}
#---------------------------------------------------------------------------------------------------------
function FeatureRemoveHandler(){
param($Event)
try{
$timer.Stop()
debug FeatureHandler
ToLog "Features removed:$($Event.SourceEventArgs.NewEvent.NumberOfEvents)" -color red

If((Test-Connection -adress $script:BaseName -Method "GET") -eq $TRUE){

  ToLog "Site $name on $env:computername is working!"
  ToLog "Sending message to Slack $slackUri"
  $result = sendToSlack -URI $slackUri -payload "Site $BaseName on $env:computername is working!"
  debug "after message"
  $timer.Start()    
}
else{
    ToLog "Site $name responded with errors" -color Red 
    ToLog $ErrorMessage -color Red  #remove from testconnection tolog
    ToLog "Trying to reinstall IIS"
     $result =  sendToSlack -URI $slackUri  -payload "Site $script:BaseName on $env:computername is working after removing features!" 
    Install-IISASP4

    If((Test-Connection -adress $script:BaseName -Method "GET") -eq $TRUE){

      ToLog "Site $name on $env:computername is working!"
      $result =  sendToSlack -URI $slackUri  -payload "Site $script:BaseName on $env:computername is working after removing features!" 
      $timer.Start()
    }
    else {
         ToLog $ErrorMessage -color Red
         ToLog "Sending message to Slack $slackUri"
         $result = sendToSlack -URI $slackUri  -payload $ErrorMessage 
         ToLog "Stopping script"
         StopAll     
    }
}
}
finally{
ToLog $("-"*20) ToLog $("-"*20) -noDatePrefix $TRUE
}
}
#---------------------------------------------------------------------------------------------------------
function Install-IISASP4(){
   ToLog "Installing IIS and ASP.NET.."
   $iisInstallResult = Install-WindowsFeature -Name Web-Server  -includeManagementTools -WarningAction SilentlyContinue
   debug "before asp.net installation"
   $aspInstallResult = dism /online /enable-feature /all /featurename:IIS-ASPNET45 /norestart

   if($iisInstallResult.Restart -eq "YES"){

        ToLog "Restart required after IIS installation. InstallationSucces = $($iisInstalResult.Success)"
        }
       
   if($iisInstallResult.Success -eq "True"){
        ToLog "Installation was successfull"
        return $TRUE
   }
   else {

        ToLog "Something gone wrong."
        ToLog "$($iisInstallResult)"
        return $FALSE
   }
}
#---------------------------------------------------------------------------------------------------------
function AllowDownloading(){
$oldPath = Get-Location
cd "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\zones\3"
if((Get-ItemProperty . -Name 1803) -ne 0){
  $result = new-itemproperty . -Name 1803 -Value 0 -Type DWORD -Force
}
cd $oldPath
}
#---------------------------------------------------------------------------------------------------------
function AddToTrusted(){
param([string]$domain)

    $oldPath = Get-Location
    set-location "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    set-location ZoneMap\EscDomains
    new-item $domain/ -Force > $null
    set-location $domain/
    new-itemproperty . -Name * -Value 2 -Type DWORD -Force > $null
    new-itemproperty . -Name http -Value 2 -Type DWORD -Force > $null
    new-itemproperty . -Name https -Value 2 -Type DWORD -Force > $null
    cd $oldPath
}
#---------------------------------------------------------------------------------------------------------
function IsNull(){
param([string]$string)

return [String]::IsNullOrWhiteSpace($string)
}
#---------------------------------------------------------------------------------------------------------
function MainAction(){
debug "Debug messages enabled:$debug"
debug Main
$gitUrlTest = Test-Connection($gitUrl)
debug "ErrorMessage - $ErrorMessage"
$needUpdate = !(IsGitUpdated($gitRssUrl))

debug "git url accessible: $gitUrlTest"
debug "Need update? $($needUpdate)"
debug "Commits: $commits"

If($gitUrlTest -and $needUpdate){

ToLog -message  $(get-date)  -noDatePrefix $TRUE 

ToLog "Zip path is $gitUrl"  
ToLog "Rss path is $gitRssUrl"
ToLog "Testing connection to zip. Accessible? $gitUrlTest"
ToLog "Require to upload project? $(!$needUpdate)"

try{ 

    $FileName = Get-NameFromUrl($gitUrl)    
    #$script:BaseName = (Get-Item  $FileName).BaseName
    $script:BaseName = $FileName -replace ".zip" , ""
    debug "Basename - $BaseName"

    if(isNull ($BaseName)) {
        $BaseName = "TestSite"
        ToLog "BaseName is empty. Current Name - TestSite"
        
    }
    
    $isDownloaded = DownloadProject -RssUrl $gitRssUrl -gitUrl $Url -FileName $FileName
    
   }
catch{

    ToLog "Error occured at download time"
    $ErrorMessage = $_.ErrorDetails.Message
    ToLog $ErrorMessage
     #>>>?????
    }
ToLog "$FileName will unzip in folder $BaseName"  
debug "Before replacing project"
   if($isDownloaded){
    ToLog "$FileName downloaded"
       if( Test-Path $Path$BaseName){
           ToLog "Folder $Path$BaseName exists. Removing.."

            StopPool -poolName $iisAppPoolName
            Remove-Item -Path $Path$BaseName -Recurse -Force  #still error with small timer time  
          }

    Unzip $Path$FileName $Path
    ToLog("Unzipping in $Path$BaseName")
    FixConfig  
    StartPool -poolName $iisAppPoolName  
}

#add test of asp-net installed
if(((Get-WindowsFeature -Name web-server| select -ExpandProperty InstallState) -ne "Installed") -or ((Get-WindowsFeature Web-Asp-Net45|select -ExpandProperty InstallState) -ne "Installed")){
  debug "Install IIS"
  $iisInstallResult = Install-IISASP4
}

ToLog "Creating web-site and pool"

createWebSiteAndPool -iisAppPoolName $iisAppPoolName -iisAppPoolDoNetVersion $iisAppPoolDotNetVersion `
-iisAppName $BaseName -directoryPath $Path$BaseName 
addToTrusted($BaseName)

debug $((Get-Service w3svc).Status)
if((Get-Service w3svc).Status -ne "Running"){
ToLog "Service w3svc not running. Starting IIS" 
StartIIS
}

debug $((Get-WebItemState "IIS:\sites\$BaseName").Value)
if((Get-WebItemState "IIS:\sites\$BaseName").Value -ne "Started"){
ToLog "Website $BaseName not started. Starting"
    Start-WebSite $BaseName
}

if ((Test-Path $HostFilePath) -eq $TRUE -and !$(Get-Content -Path $HostFilePath| Select-String -pattern "127.0.0.1 $script:BaseName" -Quiet)){

ToLog "Changing host file.."
Write-output  "127.0.0.1 $script:BaseName"| Out-File $HostFilePath  -encoding ASCII -append
}

ToLog "Testing $script:BaseName"
#start-sleep -Seconds 2
TestSite $script:BaseName
}

debug "After main IF"

$script:eventExecuting = $FALSE
}

Export-ModuleMember -Function * -Variable * 
} | Import-Module
#---------------------------------------------------------------------------------------------------------

try{
Install-IISASP4 >$null
Import-Module WebAdministration
MainAction
StartAll

}
catch{
ToLog $_ -color Red
stopall
}


