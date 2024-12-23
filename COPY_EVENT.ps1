Function Copy-Event {
param   (
         [string]$SourceMail,
         [string]$DestinationMail,
         [string]$EventSendTo,
         [int]$ThresholdInDays

  )

   $emailRegex = ‘^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$’
   if (($SourceMail -notmatch $emailRegex) -or ($DestinationMail -notmatch $emailRegex) -or ($EventSendTo -notmatch $emailRegex)) {
   Write-Host "Please provide valid email address"
   return}

   Write-Host "Connecting "
#Change According to your Tennat !!!

$clientId =""
$tenantId = ""
$clientSecret = "" 

# Get access token
$body = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}

$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body
$global:accessToken = $tokenResponse.access_token
# Get user information
$headers = @{
    Authorization = "Bearer $accessToken"
}

$userResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/users/$DestinationMail" -Headers $headers -ErrorAction SilentlyContinue
if (-not $userResponse) {
                        write-host  "Destination User Not Found" 
                        return
                           }
$userResponse2 = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/users/$SourceMail" -Headers $headers -ErrorAction SilentlyContinue
if (-not $userResponse2) {
                        write-host  "Source User Not Found" 
                        return
                           }


$Name = $userResponse.displayName
$createdDate = $userResponse.createdDateTime
Write-Host "$name Created in $createdDate"
if ($ThresholdInDays -eq $null -or $ThresholdInDays -eq "") {$ThresholdInDays = 90}
$startDate =  ((get-date $createdDate).adddays(-$ThresholdInDays)).ToString("yyyy-MM-ddTHH:mm:ssZ")
$endDate = (Get-Date $createdDate).ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Host "Searching all events sent to $EventSendTo from $startDate to $endDate"

# Retrieve events within the date range
$headers = @{
    Authorization = "Bearer $accessToken"
}
$eventsUri = "https://graph.microsoft.com/v1.0/users/$SourceMail/calendar/events?`$filter=start/dateTime ge '$startDate' and end/dateTime le '$endDate'"
$events = Invoke-RestMethod -Uri $eventsUri -Headers $headers -Method Get

# Filter events that include the AllCompany group as an attendee
$eventsSentToAllCompany = $events.value | Where-Object {
    $_.attendees -ne $null -and
    ($_.attendees | Where-Object { $_.emailAddress.address -eq $EventSendTo})
}
$C = $eventsSentToAllCompany.count
Write-Host "Found $C Events"

# Output results
$eventsSentToAllCompany | ForEach-Object {
$T+=1
    [PSCustomObject]@{
        Subject   = $_.subject
        StartTime = $_.start.dateTime
        EndTime   = $_.end.dateTime
        Organizer = $_.organizer.emailAddress.name
        Location  = $_.location.displayName
    }
}

# Copy events to the target user's calendar
foreach ($event in $eventsSentToAllCompany) {
    $newEvent = @{
        subject     = $event.subject
        start       = $event.start
        end         = $event.end
        location    = $event.location
        attendees   = $event.attendees
        isOnlineMeeting = $event.isOnlineMeeting
        onlineMeetingUrl = $event.onlineMeetingUrl
        organizer   = $event.organizer
    }

    # Include recurrence pattern if the event is recurring
    if ($event.recurrence -ne $null) {
        $newEvent.recurrence = $event.recurrence
    }

    $createEventUri = "https://graph.microsoft.com/v1.0/users/$DestinationMail/calendar/events"
    # Send the request and capture the response
    try {
        $response = Invoke-RestMethod -Uri $createEventUri -Headers $headers -Method Post -Body ($newEvent | ConvertTo-Json -Depth 10) -ContentType "application/json"
        Write-Output "$T/$C Event created successfully "
    } catch {
        Write-Output "Error: $_"
    }
}

}



