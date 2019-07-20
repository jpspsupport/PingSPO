
param(
  [Parameter(Mandatory=$true)]
  $SiteUrl,
  $AlertMe = $false,
  $Number = 4,
  $ResetCred = $false,
  $IntervalSecond = 1
)


### Alert Settings ###
$durationAlert = 30000 # 30 seconds (30 sec x 1000 millisec)
$healthAlert = 9
$statusAlert = 400
###

if (($global:PingCred -eq $null) -or $ResetCred)
{
  $global:PingCred = Get-Credential
}


function GetHostAddress($strUrl)
{
  $spaddr = "sharepoint.com"
  return $strUrl.SubString(0, $strUrl.IndexOf($spaddr) + $spaddr.Length)
}

function SendAlert($ToAddress, $Subject, $Body)
{

  $emailProperties = New-Object Microsoft.SharePoint.Client.Utilities.EmailProperties
  $emailProperties.To = [String[]]($ToAddress)
  $emailProperties.Subject = $Subject
  $emailProperties.Body = $Body
  [Microsoft.SharePoint.Client.Utilities.Utility]::SendEmail($script:context,$emailProperties)

  $script:context.ExecuteQuery()
}

[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

$script:context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)

$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($global:PingCred.UserName, $global:PingCred.Password) 
$script:context.Credentials = $credentials

$script:mySession = New-Object -TypeName Microsoft.PowerShell.Commands.WebRequestSession
$myCookie = New-Object -TypeName System.Net.Cookie("SPOIDCRL", $credentials.GetAuthenticationCookie($SiteUrl).Replace("SPOIDCRL=", ""))
$script:mySession.Cookies.Add((GetHostAddress -strUrl $SiteUrl), $myCookie)

for ($i = 0; $i -lt $Number; $i++)
{
  $d = Get-Date
  $data = Invoke-WebRequest -Uri $SiteUrl -WebSession $mySession -MaximumRedirection 0
  [int]$duration = ((Get-Date) - $d).TotalMilliseconds
  [int]$statuscode = $data.StatusCode
  [int]$healthscore = $data.Headers["x-sharepointhealthscore"]
  [string]$correlationid = $data.Headers["sprequestguid"]
  $out = ("StatusCode: " + $statuscode + "; Duration (msec): " + $duration + "; SharePointHealthScore: " + $healthscore)

  if (($statuscode -ge $statusAlert) -or ($duration -gt $durationAlert) -or ($healthscore -ge $healthAlert))
  {
    $out += ("; CorrelationId: " + $correlationid)
    if ($AlertMe)
    {
      $subject = ("PingSPO Alert " + ((Get-Date).ToString((Get-Culture).DateTimeFormat.ShortDatePattern + " hh:mm:ss")))
      $body = ("<B>Alert Detail</B><BR>" + $out.Replace("; ", "<BR>"))
      SendAlert -ToAddress $global:PingCred.UserName -Subject $subject -Body $body
    }
  }
  $out  

  Start-Sleep -Second $IntervalSecond
}

