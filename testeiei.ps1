Add-Type -Path "C:\Program Files (x86)\MySQL\MySQL Connector NET 8.0.33\MySql.Data.dll"
Connect-MicrosoftTeams

$connection = [MySql.Data.MySqlClient.MySqlConnection]@{ConnectionString = 'server=127.0.0.1;uid=root;pwd="";database=msteam' }


$connection.Open()
# ทำงานกับฐานข้อมูลที่เชื่อมต่อได้ที่นี่
$query = "SELECT * FROM courses INNER JOIN sections ON courses.course_code = sections.course_code LIMIT 10;"
$command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
$adapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
$dataset = New-Object System.Data.DataSet
$adapter.Fill($dataset) | Out-Null
#  $table = $dataset.Tables[0]
$connection.Close()

# Execute the SQL query
$command = $connection.CreateCommand()
$command.CommandText = $sql
$result = $command.ExecuteReader()

# Loop through the result and process each record
while ($result.Read()) {
    $name = $result["name"]
    $subject = $result["subject_name"]
    $room = $result["room_name"]

    Write-Host "Name: $name, Subject: $subject, Room: $room"
}
