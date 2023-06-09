$startDate = Get-Date -Year 2023 -Month 6 -Day 1
$endDate = Get-Date -Year 2023 -Month 7 -Day 10

$currentDate = $startDate
while ($currentDate -le $endDate) {
    # Perform your desired actions within the loop
    # Write-Host $currentDate.ToString("yyyy-MM-dd")
    $dayOfWeek = $currentDate.ToString("dddd", [System.Globalization.CultureInfo]::InvariantCulture)
    $shortDay = $dayOfWeek.Substring(0,2)
    $shortDay

    # Write-Output $datofweek

    # Increment the current date by one day
    $currentDate = $currentDate.AddDays(1)

}
