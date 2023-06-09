Add-Type -Path "C:\Program Files (x86)\MySQL\MySQL Connector NET 8.0.33\MySql.Data.dll"
Connect-MicrosoftTeams


$connection = [MySql.Data.MySqlClient.MySqlConnection]@{ConnectionString = 'server=127.0.0.1;uid=root;pwd=;database=msteam' }

try {
    $connection.Open()
    Write-Host "Connected to the database successfully."
    $query = "SELECT * FROM sections WHERE NOT ISNULL(ms_team_id)"
    $command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
    $adapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    $table = $dataset.Tables[0]
    $connection.Close()

}
catch {
    Write-Host "Failed to connect to the database: $_.Exception.Message"
}

foreach ($row in $table.Rows) {
    try {
        $groupId = $row["ms_team_id"]
        Remove-Team -GroupId $groupId

        $groupId = $groupId.ToString()
        $query = "UPDATE sections SET ms_team_id = null WHERE ms_team_id = '$groupId';"
        # # Execute the insert query
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteNonQuery()
        $connection.Close()
    }
    catch {
        Write-Host "Error Delete: $_.Exception.Message"
    }

}
