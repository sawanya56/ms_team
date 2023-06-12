Add-Type -Path "C:\Program Files (x86)\MySQL\MySQL Connector NET 8.0.33\MySql.Data.dll"
Connect-MicrosoftTeams
# $connectionString = "server=<localhost>;port=<3306>;database=<msteam>;uid=<root>;pwd=<''>;"
# $connection = New-Object MySql.Data.MySqlClient.MySqlConnection($connectionString)
# $connection.Open()

$addMembersCsvPath = "D:\Ms Teams\UserinRoom.csv"

$connection = [MySql.Data.MySqlClient.MySqlConnection]@{ConnectionString = 'server=127.0.0.1;uid=root;pwd="";database=msteam' }


try {
    $connection.Open()
    Write-Host "Connected to the database successfully."

    # ทำงานกับฐานข้อมูลที่เชื่อมต่อได้ที่นี่
    # $query = "SELECT * FROM view_sections LIMIT 1;"
    $query = "SELECT
    sections.section,
    sections.team_name,
    courses.description
    FROM sections 
    LEFT JOIN courses 
    ON sections.course_code = courses.course_code 
    GROUP BY section LIMIT 1;"
    # $query = "SELECT * FROM courses INNER JOIN sections ON courses.course_code = sections.course_code LIMIT 1;"
    $command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
    $adapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    $table = $dataset.Tables[0]
    $connection.Close()

    foreach ($row in $table.Rows) {
        $year = $row["year"]
        $course_code = $row["course_code"]
        $term = $row["term"]
        $section = $row["section"]
        $course_name = $row["course_name"]
       
        $year = $year.ToString()
        $course_code = $course_code.ToString()
        $term = $term.ToString()
        $section = $section.ToString()
        $course_name = $course_name.ToString()

        #การสร้างห้อง
        # $row['team_name']
        # $teamName = "ห้องทีม $i"
        # $teamDescription = "คำอธิบายห้องทีม $i"
        $teamDescription = $row['description']
        $team = New-Team -DisplayName $row['team_name'] -Description $teamDescription
        $sectionId = $row["id"]
        # แสดงค่า groupId
        $groupId = $team.GroupId
        $team
        Write-Output ("รหัสห้องทีมของห้อง {0}: {1}" -f $roomName, $groupId)
        $connection.Open()
        # Define the SQL query for the insert statement
        $groupId = $groupId.ToString()
        $query = "UPDATE sections SET ms_team_id = '$groupId' WHERE id = $sectionId"
        $query

        # Execute the insert query
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteNonQuery()
        $connection.Close()

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


        # อ่านข้อมูลจากไฟล์เพิ่มสมาชิก
        $addMembersData = Import-Csv -Path $addMembersCsvPath
       
        foreach ($memberData in $addMembersData) {
            $email = $memberData.Email
            $role = $memberData.Role
    
            # เพิ่มสมาชิกลงในห้องทีม
            Add-TeamUser -GroupId $team.GroupId -User $email -Role $role
        }
    }
    
    $connection.Close()
    Write-Host "Connection closed."
}
catch {
    Write-Host "Failed to connect to the database: $_.Exception.Message"
}