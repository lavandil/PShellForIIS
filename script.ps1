
Cd C:\
cls

Add-Type -AssemblyName System.IO.Compression.FileSystem


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
    $EventExecuting = $TRUE
  

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
           write-host "Adress $adress"
           $httpResponse = Invoke-WebRequest $adress -Method $Method 
           $httpResponse.BaseResponse.Close()
           
            if( ($httpResponse |select -ExpandProperty StatusCode) -eq 200){           
                return $TRUE
             }
        }
        catch {
            $global:ErrorMessage = $_   
            Write-Host  "in test $_"            
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

Import-Module WebAdministration
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
    $result = Invoke-WebRequest -URI $URI -Method Post -ContentType "application/json" -Body (ConvertTo-Json -Compress -InputObject $objectToPayload)

    if($result.StatusCode -eq 200){

        ToLog "Slack got Message"
        return $TRUE
    }
    return $FALSE
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
        #write-Host "Download complete.."
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
function StopAll(){
$timer.Stop()
Unregister-Event thetimer

Get-EventSubscriber| Unregister-Event
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
 $state= (Get-WebAppPoolState -Name $poolName).Value

     if($state -ne "Stopped"){
      Stop-WebAppPool  -Name $poolName
     }


 while((Get-WebAppPoolState -Name $poolName).Value -ne "Stopped"){
    Start-Sleep -s 5
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
    Start-Sleep -s 2
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

If((Test-Connection -adress $name -Method "GET") -eq $TRUE){

    ToLog "Site $name on $env:computername is working!"
    $result =  sendToSlack -URI $slackUri  -payload "Site $global:BaseName on $env:computername  is working!"  
}
else {

    ToLog "Site $name responded with errors"
    ToLog $ErrorMessage #remove from testconnection tolog
    ToLog "Sending message to Slack $slackUri"
    $result = sendToSlack -URI $slackUri  -payload $ErrorMessage     
    }

    ToLog $("-"*20) -noDatePrefix $TRUE
}
#---------------------------------------------------------------------------------------------------------
function TimerHandler(){
try{
    write-host in ACTION
    #throw "action error" 
    write-host "bool  - $eventExecuting"
    if(!$eventExecuting){
    MainAction
    }
}
catch{
    ToLog "Error occured.Stopping script"
    stop

    ToLog $($Error) -color Red   
    ToLog "Script stopped. Sending message to Slack"
       
    $result =  sendToSlack -URI $slackUri  -payload "$Error"
  
    ToLog "Slack get Message" 
    write-host $_.Exception
  
}
finally{ 
    
    Write-Host "IN FINALLY"
    #Write-Host $Error

}
}
#---------------------------------------------------------------------------------------------------------
function ProcessStopHandler(){
$timer.Stop()
#write-host action-2
$message = "IIS stopped. Trying to restart. " 
ToLog $message
$SlackPayload = $message

    if((Get-WindowsFeature -Name web-server| select -ExpandProperty InstallState) -ne "Installed"){
      
      $message = "IIS was uninstalled. Trying to reload project"
      ToLog $message
      $slackPayload+= $message
     
      $commits = 0
      MainAction
    
    }

    else{

    startIIS
    Start-Sleep -Seconds 5
        if((Get-Service w3svc).Status -eq "Running"){
            
            $message = "IIS working. Testing $global:BaseName"
            ToLog $message
            $SlackPayload+=$message
            $result =  sendToSlack -URI $slackUri  -payload $slackPayload 
            TestSite $global:BaseName
        }
        else{
         $message ="IIS couldn't be restarted on $env:computername"
         ToLog $message
         $SlackPayload += $message
                
         $result =  sendToSlack -URI $slackUri  -payload $slackPayload 
        }
    }
    $timer.Start()       

}
#---------------------------------------------------------------------------------------------------------
function FeatureUninstallHandler(){

}
#---------------------------------------------------------------------------------------------------------
function MainAction(){
Write-host Main
$gitUrlTest = Test-Connection($gitUrl)
$needUpdate = IsGitUpdated($gitRssUrl)

write-host $gitUrlTest
If($gitUrlTest -and $needUpdate){

ToLog -message  $(get-date)  -noDatePrefix $TRUE 

ToLog "Zip path is $gitUrl"  
ToLog "Rss path is $gitRssUrl"
ToLog "Testing coonection to zip. Accessible? $gitUrlTest"
ToLog "Require to upload project? $needUpdate"

try{ 

    $FileName = Get-NameFromUrl($gitUrl) |% {($_ -split "=")[1]}    
    $global:BaseName = (Get-Item  $FileName).BaseName
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

            StopPool -poolName $iisAppPoolName
            Remove-Item -Path $Path$BaseName -Recurse -Force  #still error with small timer time  
          }

    Unzip $Path$FileName $Path
    ToLog("Unzipping in $Path$BaseName")
    FixConfig  
    StartPool -poolName $iisAppPoolName
  
}

#add test of asp-net installed
if((Get-WindowsFeature -Name web-server| select -ExpandProperty InstallState) -ne "Installed"){
    ToLog "Installing IIS and ASP.NET.."
   ToLog $(Install-WindowsFeature -Name Web-Server -includeManagementTools)
   ToLog $(dism /online /enable-feature /all /featurename:IIS-ASPNET45)
}

ToLog "Creating web-site and pool" #move to function
$oldPath = Get-Location
createWebSiteAndPool -iisAppPoolName $iisAppPoolName -iisAppPoolDoNetVersion $iisAppPoolDotNetVersion `
-iisAppName $BaseName -directoryPath $Path$BaseName 
cd $oldPath


start-sleep -Seconds 5

if ((Test-Path $HostFilePath) -eq $TRUE -and !$(Get-Content -Path $HostFilePath| Select-String -pattern "127.0.0.1 $BaseName" -Quiet)){

ToLog "Changing host file.."
Write-output  "127.0.0.1 $BaseName"| Out-File $HostFilePath  -encoding ASCII -append
}

ToLog "Testing $BaseName"

    TestSite $BaseName
}


write-host "After main IF"
$originalUrl = Get-RedirectedUrl($shortUrl)

#Write-Host "in main $eventExecuting"
$global:eventExecuting = $FALSE
    }

Export-ModuleMember -Function * -Variable *
} | Import-Module
#---------------------------------------------------------------------------------------------------------

try{
MainAction
}
catch{
ToLog $_
}


$timer = New-Object System.Timers.Timer

#$event = Register-ObjectEvent -InputObject $timer -EventName elapsed -Action $action
$EventJob = Register-ObjectEvent -InputObject $timer -EventName elapsed –SourceIdentifier  thetimer -Action {timerhandler} -OutVariable out


$timer.Interval = 30000
$timer.AutoReset = $true
$timer.start()

$query1 = @" 
 Select * From __InstanceOperationEvent Within 1 
    Where TargetInstance Isa 'Win32_Service' 
   and 
   TargetInstance.Name='w3svc'
   and
   TargetInstance.State="Stopped"
"@

$eventJob2 = Register-WmiEvent -Query $query1 -Action {ProcessStopHandler}

# Get-EventSubscriber| Unregister-Event