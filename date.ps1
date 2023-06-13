Add-Type -Path "C:\Program Files (x86)\MySQL\MySQL Connector NET 8.0.33\MySql.Data.dll"
$connection = [MySql.Data.MySqlClient.MySqlConnection]@{ConnectionString = 'server=127.0.0.1;uid=root;pwd="";database=msteam' }



$startDate = Get-Date -Year 2023 -Month 4 -Day 17
$endDate = Get-Date -Year 2023 -Month 6 -Day 4
$currentDate = $startDate

# section
$sectionId = 333145
$sectionId = $sectionId.ToString()
$query = "SELECT * FROM class WHERE section = '$sectionId';"
$connection.Open()
$command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
$adapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
$dataset = New-Object System.Data.DataSet
$adapter.Fill($dataset) | Out-Null
$table = $dataset.Tables[0]
$connection.Close()
# foreach ($row in $table.Rows) {
#     $row;
# }

while ($currentDate -le $endDate) {
    # Perform your desired actions within the loop
    # Write-Host $currentDate.ToString("yyyy-MM-dd")
    $dayOfWeek = $currentDate.ToString("dddd", [System.Globalization.CultureInfo]::InvariantCulture)
    $shortDay = $dayOfWeek.Substring(0, 2)
    $shortDay = $shortDay.ToUpper()

    foreach ($date_db in $table.Rows) {
        
        if ($shortDay -eq $date_db['week_of_day']) {
            <# Action to perform if the condition is true #>
            # $date_db
            # Write-Host $currentDate.ToString("yyyy-MM-dd")

            $time = $date_db['start_time']
            $date = Get-Date -Hour 8 -Minute 0 -Second 0 -Millisecond 0
            $dateTime = $date.Date.Add([TimeSpan]::Parse($time))
            
            $durationtime = New-TimeSpan -Minutes $date_db['duration_time']
            $newdatetime = $dateTime.Add($durationtime)

            $curent_date = $currentDate.ToString("yyyy-MM-dd")
            $start_time = $dateTime.ToString("H:mm")
            $end_time = $newdatetime.ToString("H:mm")
            Write-Host $curent_date $shortDay $start_time $end_time
        }
        # $shortDay = = $date_db['week_of_day']

    }
    # Write-Output $datofweek

    # Increment the current date by one day
    $currentDate = $currentDate.AddDays(1)

}
