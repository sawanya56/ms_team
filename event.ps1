# Import the required modules
Import-Module MicrosoftTeams

# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Define the event details
$roomName = "Meeting"
$startTime = "09:00"
$endTime = "12:00"
$startDate = Get-Date "2023-06-01"
$endDate = Get-Date "2023-07-31"
$daysOfWeek = @("Tuesday")

# Loop through each day within the date range
while ($startDate -le $endDate) {
    # Check if the current day is one of the scheduled days of the week
    if ($daysOfWeek -contains $startDate.DayOfWeek.ToString()) {
        # Create the start and end date/time strings
        $startDateTime = $startDate.ToString("yyyy-MM-dd") + "T" + $startTime + ":00"
        $endDateTime = $startDate.ToString("yyyy-MM-dd") + "T" + $endTime + ":00"

        $GroupId = '59a783dd-9276-4fb0-b6d0-8d7a8d65342e'

        # Create the event parameters
        $eventParams = @{
            "Subject" = $eventName
            "Start"   = $startDateTime
            "End"     = $endDateTime
        }

        # Create the event
        # New-TeamChannelMessage -TeamId 59a783dd-9276-4fb0-b6d0-8d7a8d65342e  -MessageType "Chat" -Content $eventParams
# -ChannelId "YOUR_CHANNEL_ID"

        # สร้างเหตุการณ์ (event) ของห้องทีม
        New-Event -GroupId $GroupId -Subject $roomName -StartDateTime $startDateTime -EndDateTime $endDateTime
        

        Write-Host "Event created for $($startDate.ToString("yyyy-MM-dd"))"
    }

    # Move to the next day
    $startDate = $startDate.AddDays(1)
}
