Add-Type -Path "C:\Program Files (x86)\MySQL\MySQL Connector NET 8.0.33\MySql.Data.dll"
Connect-MicrosoftTeams

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
        # $year = $row["year"]
        # $course_code = $row["course_code"]
        # $term = $row["term"]
        $section = $row["section"]
        $team_name = $row["team_name"]
        # $course_name = $row["course_name"]
       
        # $year = $year.ToString()
        # $course_code = $course_code.ToString()
        # $term = $term.ToString()
        $section = $section.ToString()
        # $course_name = $course_name.ToString()
        # $team_name+$section
        $full_name = $team_name + $section
        $full_name
        #การสร้างห้อง
        # $row['team_name']
        # $teamName = "ห้องทีม $i"
        # $teamDescription = "คำอธิบายห้องทีม $i"
        $teamDescription = $row['description']
        $team = New-Team -DisplayName $row['team_name'] -Description $teamDescription
        # $sectionId = $row["id"]
        # # แสดงค่า groupId
        $groupId = $team.GroupId
        $team
        Write-Output ("รหัสห้องทีมของห้อง {0}: {1}" -f $team, $groupId)

        $connection.Open()
        # Define the SQL query for the insert statement
        $groupId = $groupId.ToString()
        $query = "UPDATE sections SET ms_team_id = '$groupId' WHERE section = '$section';"
        $query

        # # Execute the insert query
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteNonQuery()
        $connection.Close()
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

