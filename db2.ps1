Add-Type -Path "C:\Program Files (x86)\MySQL\MySQL Connector NET 8.0.33\MySql.Data.dll"
Connect-MicrosoftTeams

function AddStudents {
    param(
        [Parameter(Position = 0, Mandatory = $true)]
        [string] $sectionId,
        [Parameter(Position = 1, Mandatory = $true)]
        [string] $groupId
    )

    $connection.Open()
    $sectionId = $sectionId.ToString()
    $query = "SELECT * FROM view_students WHERE section = '$sectionId';"
    $command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
    $adapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    $table = $dataset.Tables[0]
    $connection.Close()
    foreach ($row in $table.Rows) {
        # $row['student_mail']
        $studentMail = $row['student_mail'];
        Write-Host "$sectionId + $studentMail "
        Add-TeamUser -GroupId $groupId -User $studentMail -Role "member"
    }
    return $sectionId
}

function AddInstructor {
    param(
        [Parameter(Position = 0, Mandatory = $true)]
        [string] $sectionId,
        [Parameter(Position = 1, Mandatory = $true)]
        [string] $groupId
    )

    $connection.Open()
    $sectionId = $sectionId.ToString()
    $query = "SELECT * FROM view_instructors WHERE section = '$sectionId';"
    $command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
    $adapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    $table = $dataset.Tables[0]
    $connection.Close()
    foreach ($row in $table.Rows) {
        $instructor_mail = $row['instructor_mail'];
        Write-Host "$sectionId + $instructor_mail "
        Add-TeamUser -GroupId $groupId -User $instructor_mail -Role "owner"
    }
    return $sectionId
}

function AddEvent {
    param(
        [Parameter(Position = 0, Mandatory = $true)]
        [string] $sectionId,
        [Parameter(Position = 1, Mandatory = $true)]
        [string] $groupId,
        [Parameter(Position = 2, Mandatory = $true)]
        [string] $startDate,
        [Parameter(Position = 3, Mandatory = $true)]
        [string] $endDate
    )
    $startDate = $startDate.Replace("00:00:00", "");
    $arrayStartDate = $startDate.Split("/");

    $endDate = $endDate.Replace("00:00:00", "");
    $arrayEndDate = $endDate.Split("/");


    $startDate = Get-Date -Year $arrayStartDate[2] -Month $arrayStartDate[0] -Day $arrayStartDate[1]
    $newStartDate = [DateTime]::ParseExact($startDate, "MM/dd/yyyy HH:mm:ss", $null)

    $endDate = Get-Date -Year $arrayEndDate[2] -Month $arrayEndDate[0] -Day $arrayEndDate[1]
    $newEndDate = [DateTime]::ParseExact($endDate, "MM/dd/yyyy HH:mm:ss", $null)
   
    
    $currentDate = $newStartDate

    # section
    $sectionId = $sectionId.ToString()
    $query = "SELECT * FROM class WHERE section = '$sectionId';"
    $connection.Open()
    $command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
    $adapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    $table = $dataset.Tables[0]
    $connection.Close()

    while ($currentDate -le $newEndDate) {
        # Perform your desired actions within the loop
        # Write-Host $currentDate.ToString("yyyy-MM-dd")
        # return $currentDate
        $dayOfWeek = $currentDate.ToString("dddd", [System.Globalization.CultureInfo]::InvariantCulture)
        $shortDay = $dayOfWeek.Substring(0, 2)
        $shortDay = $shortDay.ToUpper()
        
        foreach ($date_db in $table.Rows) {
        
            if ($shortDay -eq $date_db['week_of_day']) {

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
}

try {
    $connection = [MySql.Data.MySqlClient.MySqlConnection]@{ConnectionString = 'server=127.0.0.1;uid=root;pwd="";database=msteam' }
    $connection.Open()
    Write-Host "Connected to the database successfully."

    # ทำงานกับฐานข้อมูลที่เชื่อมต่อได้ที่นี่
    # $query = "SELECT * FROM sections WHERE id = '7' LIMIT 1;"
    $query = "SELECT * FROM view_sections;"
    $command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
    $adapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    $table = $dataset.Tables[0]
    $connection.Close()
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)"
}



foreach ($row in $table.Rows) {

    try {

        $section = $row["section"]
        # $team_name = $row["team_name"]
        $section = $section.ToString()

        # $full_name = $team_name + $section
        # $full_name
        #การสร้างห้อง

        $teamDescription = $row['description']

        # Check Empty String Or null
        if ([string]::IsNullOrEmpty($teamDescription)) {
            $teamDescription = "Description";
        }

        # Check Team exist
        # $team = New-TeamFromTemplate -DisplayName $section -Description $section -TemplateAppId "8ec74a39-ddf6-41e1-b0a2-ff0459ea8eb8"
        # $team = New-TeamFromTemplate -DisplayName "New Team" -Description "Description of the new team" -TemplateAppId "8ec74a39-ddf6-41e1-b0a2-ff0459ea8eb8"
 
        $team = New-Team -DisplayName $section -Description $section -Template "EDU_Class"

        if ($team) {
            # Write-Host "Team created successfully!"
            $groupId = $team.GroupId
            # $team
            Write-Output ("Team created successfully! KEY ID : " + $groupId)
        
            # UPDATE MS TEAM KEY
            $connection.Open()
            $groupId = $groupId.ToString()
            $query = "UPDATE sections SET ms_team_id = '$groupId' WHERE section = '$section';"
            $command = $connection.CreateCommand()
            $command.CommandText = $query
            $command.ExecuteNonQuery()
            $connection.Close()
            # UPDATE MS TEAM KEY

            # Call function for add student
            # AddStudents -sectionId $section -groupId $groupId
            # AddInstructor -sectionId $section -groupId $groupId
            # $startDate = $row['course_start_date']
            # $endDate = $row['course_end_date']
            # AddEvent -sectionId $section -groupId $groupId -startDate $startDate -endDate $endDate
        }
        else {
            Write-Output ("Crearte Fail {0}: {1}" -f $team, $row)
        }
        
        
       
        
    }
    catch {
        Write-Host "An error occurred: $($_.Exception.Message)"
    }
    

    # $term
    # $roomName = $year + "/" + $term + " " + $course_code + " " + $section + " " + $course_name
    # $roomName
    # $teamDescription = $subjectCode + "-" + $subjectName

    # $team = New-Team -DisplayName $roomName -Description $teamDescription
   
    # # แสดงค่า groupId
    # $groupId = $team.GroupId
    # Write-Output ("รหัสห้องทีมของห้อง {0}: {1}" -f $roomName, $groupId)


    # ทำสิ่งที่คุณต้องการกับข้อมูลแต่ละแถวที่นี่
    # Write-Host "Column 1: $column1Value, Column 2: $column2Value"
    # end สร้างห้อง

}
    
# $connection.Close()
# Write-Host "Connection closed."

