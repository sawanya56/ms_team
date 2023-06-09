Connect-MicrosoftTeams

$csvPath = "D:\Ms Teams\room1.csv"
$rooms = Import-Csv $csvPath
foreach ($room in $rooms) {
    $teamName = $room."ชื่อห้อง"
    $teamDescription = "คำอธิบายห้อง $teamName"

    New-Team -DisplayName $teamName -Description $teamDescription
}
