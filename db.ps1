Add-Type -Path "C:\Program Files (x86)\MySQL\MySQL Connector NET 8.0.33\MySql.Data.dll"

# $connectionString = "server=<localhost>;port=<3306>;database=<msteam>;uid=<root>;pwd=<''>;"
# $connection = New-Object MySql.Data.MySqlClient.MySqlConnection($connectionString)
# $connection.Open()


$connection = [MySql.Data.MySqlClient.MySqlConnection]@{ConnectionString = 'server=127.0.0.1;uid=root;pwd=;database=msteam' }


try {
    $connection.Open()
    Write-Host "Connected to the database successfully."

    # ทำงานกับฐานข้อมูลที่เชื่อมต่อได้ที่นี่
    $query = "SELECT * FROM class INNER JOIN courses ON class.course_code = courses.course_code;"
    $command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
    $adapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    $table = $dataset.Tables[0]

    foreach ($row in $table.Rows) {
        # $year = $row["year"]
        # $course_code = $row["course_code"]
        # $term = $row["term"]
        # $section = $row["section"]
        # $course_name = $row["course_name"]
       
        # $course_name
         # $course_code

         # แปลงค่าวันเวลาของ MySQL เป็น System.DateTime
        #  $start_time = [DateTime]::ParseExact($row["start_time"], "yyyy-MM-dd HH:mm:ss", $null)
        #  $course_start_date = [DateTime]::ParseExact($row["course_start_date"], "yyyy-MM-dd HH:mm:ss", $null)
        #  $course_end_date = [DateTime]::ParseExact($row["course_end_date"], "yyyy-MM-dd HH:mm:ss", $null)

       

    #ยการสร้างห้อง
        #  $year = $year
        #  $term = $term
        #  $sec = $section
        #  $course_Code = $course_code
        # $subjectName = $data."ชื่อวิชา"

        # $roomName = $year + "/" + $term + " " + $subjectCode + " " + $sec + " " + $subjectName
        # $teamDescription = $subjectCode + "-" + $subjectName

        # $team = New-Team -DisplayName $roomName -Description $teamDescription
   
        # # แสดงค่า groupId
        # $groupId = $team.GroupId
        # Write-Output ("รหัสห้องทีมของห้อง {0}: {1}" -f $roomName, $groupId)


        # ทำสิ่งที่คุณต้องการกับข้อมูลแต่ละแถวที่นี่
        # Write-Host "Column 1: $column1Value, Column 2: $column2Value"
    }
    
    $connection.Close()
    Write-Host "Connection closed."
}
catch {
    Write-Host "Failed to connect to the database: $_.Exception.Message"
}