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
        [string] $sectionId
    )

    

    return $sectionId
}

try {
    $connection = [MySql.Data.MySqlClient.MySqlConnection]@{ConnectionString = 'server=127.0.0.1;uid=root;pwd="";database=msteam' }
    $connection.Open()
    Write-Host "Connected to the database successfully."

    # ทำงานกับฐานข้อมูลที่เชื่อมต่อได้ที่นี่
    # $query = "SELECT * FROM sections WHERE id = '7' LIMIT 1;"
    $query = "SELECT * FROM view_sections LIMIT 1;"
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
       
        $team = New-Team -DisplayName $section -Description $section

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
            AddStudents -sectionId $section -groupId $groupId
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

