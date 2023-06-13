# Authenticate to Microsoft Graph
Connect-Graph -Scopes "Group.ReadWrite.All"

# Set up meeting details
$teamId = "8c642491-6078-483f-a3d4-b09107359942"
$subject = "Team Meeting"
$startDateTime = "2023-06-14T10:00:00"
$endDateTime = "2023-06-14T12:00:00"

# Create the meeting request
$meetingRequest = @{
    subject = $subject
    startDateTime = $startDateTime
    endDateTime = $endDateTime
    attendees = @('mju6204101356@mju.ac.th')
    location = @{
        displayName = "Meeting Room"
    }
}

# Send the meeting request to the team
$endpoint = "/teams/$teamId/events"
$meeting = Invoke-GraphRequest -Method Post -Endpoint $endpoint -Body $meetingRequest

# Display the meeting details
$meeting

# Disconnect from Microsoft Graph
Disconnect-Graph
