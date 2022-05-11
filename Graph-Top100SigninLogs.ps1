##############################################################################################
#
# This script is not officially supported by Microsoft, use it at your own risk.
# Microsoft has no liability, obligations, warranty, or responsibility regarding
# any result produced by use of this file.
#
##############################################################################################
# The sample scripts are not supported under any Microsoft standard support
# program or service. The sample scripts are provided AS IS without warranty
# of any kind. Microsoft further disclaims all implied warranties including, without
# limitation, any implied warranties of merchantability or of fitness for a particular
# purpose. The entire risk arising out of the use or performance of the sample scripts
# and documentation remains with you. In no event shall Microsoft, its authors, or
# anyone else involved in the creation, production, or delivery of the scripts be liable
# for any damages whatsoever (including, without limitation, damages for loss of business
# profits, business interruption, loss of business information, or other pecuniary loss)
# arising out of the use of or inability to use the sample scripts or documentation,
# even if Microsoft has been advised of the possibility of such damages
##############################################################################################

#documentation for query: https://docs.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0&tabs=http

#######################################################################
#######################################################################
#############--Please update the script variables below--##############
#App ID of app registration
$AppId = "xxxxxxxxxxxxx" 

#Client Secret for App Registration
$client_secret = 'xxxxxxxxxxxxx'

#Tenant ID of app registration
$TenantId = "xxxxxxxxxxxxx"
#######################################################################
#######################################################################
 
#######################################################################
#Get access token for App Registration
#######################################################################
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $client_secret
    grant_type    = "client_credentials"
}

try 
{
    $tokenRequest = Invoke-WebRequest -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop
}
catch 
{ 
    Write-Host "Unable to obtain access token" -ForegroundColor Red; return 
}

$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

#######################################################################
#Auth header Details
#######################################################################
$authHeader = @{
   'Content-Type'='application\json'
   'Authorization'="Bearer $token"
}

#######################################################################
#Execute the Microsoft Graph query
#######################################################################
$Uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?$top=100"

$retryCount = 0
$maxRetries = 3
$pauseDuration = 5

write-host "Running Graph Query..." -ForegroundColor Yellow
try
{
    $SignInActivity = Invoke-WebRequest -Headers $AuthHeader -Uri $Uri
}
catch
{
    Write-Host "StatusCode: " $_.Exception.Response.StatusCode.value__
    Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
        if($_.ErrorDetails.Message)
            {
                Write-Host "Inner Error: $_.ErrorDetails.Message"
            }
        # check for a specific error so that we can retry the request otherwise, set the url to null so that we fall out of the loop
         if($_.Exception.Response.StatusCode.value__ -eq 429 )
         {
            # just ignore, leave the url the same to retry but pause first
            if($retryCount -ge $maxRetries)
            {
            # not going to retry again
            $Uri = $null
            Write-Host 'Not going to retry...'
            }   
            else 
            {
            $retryCount += 1
            Write-Host "Retry attempt $retryCount after a $pauseDuration second pause..."
            Start-Sleep -Seconds $pauseDuration
            }

        } 
        else 
        {
            # not going to retry -- set the url to null to fall back out of the while loop
            $Uri = $null
        }       

}


Write-host "Completed Graph Query" -ForegroundColor Green
Write-Host ""

#######################################################################
#Format/Export results to csv
#######################################################################
Write-Host "Exporting results to CSV..." -ForegroundColor Yellow
$path = "c:\temp\SignInLogs_$(get-date -Format MM-dd-yyyy--HH.mm.ss).csv"

$result = ($SignInActivity.Content | ConvertFrom-Json).Value
$result | Export-Csv -Path  $path -NoTypeInformation -NoClobber -Encoding UTF8 -UseCulture

Write-Host "Completed exporting result" -ForegroundColor Green
Write-Host ""
Write-Host "Top 100 Sign in Logs File Output: " $path -ForegroundColor Green -BackgroundColor Black