
param(
  [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
  $SiteUrl,
  $Number = 4,
  $IntervalSecond = 1,
  [switch]$ResetCred,
  [switch]$AlertMe,
  [switch]$ReturnPSObject
)

$PrevProgressPreference = $ProgressPreference
$ProgressPreference = "silentlyContinue"

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

function WriteOut($InputObject)
{
  if ($ReturnPSObject)
  {
    Write-Host $InputObject
  }
  else {
    Write-Output $InputObject
  }

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

$ResultRows = New-Object System.Collections.ArrayList

$ResultSummary = New-Object PSObject
$ResultSummary | Add-Member -MemberType NoteProperty -Name URL -Value $SiteUrl
$ResultSummary | Add-Member -MemberType NoteProperty -Name StartTime -Value (Get-Date)

WriteOut ""
WriteOut ("Requesting URL (" + $ResultSummary.URL + ")")

for ($i = 0; $i -lt $Number; $i++)
{
  if ($i -ne 0)
  {
    Start-Sleep -Second $IntervalSecond
  }

  $d = Get-Date
  try {
    $data = Invoke-WebRequest -Uri $SiteUrl -WebSession $mySession -MaximumRedirection 0
    $statuscode = $data.StatusCode
  }
  catch {
    $data = ([system.net.webexception]$_.Exception.GetBaseException()).Response;
    $statuscode = $data.StatusCode.value__
  }

  $ResultRow = New-Object psobject
  $ResultRow | Add-Member -MemberType NoteProperty -Name EventTime -Value (Get-Date)
  $ResultRow | Add-Member -MemberType NoteProperty -Name Duration -Value ([int]((Get-Date) - $d).TotalMilliseconds)
  $ResultRow | Add-Member -MemberType NoteProperty -Name StatusCode -Value $statuscode
  $ResultRow | Add-Member -MemberType NoteProperty -Name HealthScore -Value  $data.Headers["x-sharepointhealthscore"]
  $ResultRow | Add-Member -MemberType NoteProperty -Name CorrelationId -Value $data.Headers["sprequestguid"]

  $out = ("Response: StatusCode=" + $ResultRow.StatusCode + ", Duration=" + $ResultRow.Duration + "ms, HealthScore=" + $ResultRow.HealthScore)

  if (($ResultRow.StatusCode -ge $statusAlert) -or ($ResultRow.Duration -ge $durationAlert) -or ($ResultRow.HealthScore -ge $healthAlert))
  {
    $out += (", CorrelationId=" + $ResultRow.CorrelationId)
    if ($AlertMe)
    {
      $subject = ("PingSPO Alert " + ((Get-Date).ToString((Get-Culture).DateTimeFormat.ShortDatePattern + " hh:mm:ss")))
      $body = ("<B>Alert Detail</B><BR>" + $out.Replace("; ", "<BR>"))
      SendAlert -ToAddress $global:PingCred.UserName -Subject $subject -Body $body
    }
  }
  WriteOut $out
  [void]$ResultRows.Add($ResultRow)
}

$ResultSummary | Add-Member -MemberType NoteProperty -Name EndTime -Value (Get-Date)
$ResultSummary | Add-Member -MemberType NoteProperty -Name TotalCount -Value $ResultRows.Count
$ResultSummary | Add-Member -MemberType NoteProperty -Name StatusCodeAlertCount -Value (($ResultRows.StatusCode | where { $_ -ge $statusAlert}).Count)
$ResultSummary | Add-Member -MemberType NoteProperty -Name HealthScoreAlertCount -Value (($ResultRows.HealthScore | where { $_ -ge $healthAlert}).Count)
$ResultSummary | Add-Member -MemberType NoteProperty -Name DurationAlertCount -Value (($ResultRows.Duration | where { $_ -ge $durationAlert}).Count)
$ResultSummary | Add-Member -MemberType NoteProperty -Name DurationMin -Value (($ResultRows.Duration | Measure-Object -Minimum).Minimum)
$ResultSummary | Add-Member -MemberType NoteProperty -Name DurationMax -Value (($ResultRows.Duration | Measure-Object -Maximum).Maximum)
$ResultSummary | Add-Member -MemberType NoteProperty -Name DurationAvg -Value (($ResultRows.Duration | Measure-Object -Average).Average)

$ResultSummary | Add-Member -MemberType NoteProperty -Name HealthScoreMin -Value (($ResultRows.HealthScore | Measure-Object -Minimum).Minimum)
$ResultSummary | Add-Member -MemberType NoteProperty -Name HealthScoreMax -Value (($ResultRows.HealthScore | Measure-Object -Maximum).Maximum)
$ResultSummary | Add-Member -MemberType NoteProperty -Name HealthScoreAvg -Value (($ResultRows.HealthScore | Measure-Object -Average).Average)


$ResultSummary | Add-Member -MemberType NoteProperty -Name ResultRows -Value $ResultRows

WriteOut ""
WriteOut "Alert Statistics"
WriteOut (Write-Output("    StatusCode = {0} ({1}%), HealthScore = {2} ({3}%), Duration = {4} ({5}%)" -f $ResultSummary.StatusCodeAlertCount, ($ResultSummary.StatusCodeAlertCount / $ResultSummary.TotalCount * 100), $ResultSummary.HealthScoreAlertCount, ($ResultSummary.HealthScoreAlertCount / $ResultSummary.TotalCount * 100), $ResultSummary.DurationAlertCount, ($ResultSummary.DurationAlertCount / $ResultSummary.TotalCount * 100)))
WriteOut "HealthScore Statistics"
WriteOut (Write-Output("    Minimum = {0}, Maximum = {1}, Average = {2}" -f $ResultSummary.HealthScoreMin, $ResultSummary.HealthScoreMax, $ResultSummary.HealthScoreAvg))
WriteOut "Duration Statistics"
WriteOut (Write-Output("    Minimum = {0}ms, Maximum = {1}ms, Average = {2}ms" -f $ResultSummary.DurationMin, $ResultSummary.DurationMax, $ResultSummary.DurationAvg))
WriteOut ""

if ($ReturnPSObject)
{
  return $ResultSummary
}
$ProgressPreference = $PrevProgressPreference
