Connect-MicrosoftTeams

$csvPath = "D:\Ms Teams\test1.csv"

# อ่านข้อมูลจากไฟล์ CSV
$roomData = Import-Csv -Path $csvPath

# สร้างห้องและเพิ่มรหัสวิชาและชื่อวิชา
foreach ($data in $roomData) {
    $year = $data."ปี"
    $term = $data."เทอม"
    $sec = $data."กลุ่ม"
    $subjectCode = $data."รหัสวิชา"
    $subjectName = $data."ชื่อวิชา"

    $roomName = $year + "/" + $term + " " + $subjectCode + " " + $sec + " " + $subjectName
    $teamDescription = $subjectCode + "-" + $subjectName

    $team = New-Team -DisplayName $roomName -Description $teamDescription
   
    # แสดงค่า groupId
    $groupId = $team.GroupId
    Write-Output ("รหัสห้องทีมของห้อง {0}: {1}" -f $roomName, $groupId)

    New-Event -GroupId $groupId -Subject "meeting01" -StartDateTime "2023-06-10 08:00:00" -EndDateTime "2023-06-10 10:00:00"
    New-Event -GroupId $groupId -Subject "meeting02" -StartDateTime "2023-06-11 08:00:00" -EndDateTime "2023-06-11 10:00:00"

    # # แสดงค่า groupId โดยเฉพาะตัวเลขและจำกัดจำนวนหลักไม่เกิน 7 หลัก
    # $groupId = $team.GroupId -replace "[^0-9]", "" -replace "^(\d{7}).*", '$1'
    # Write-Output ("รหัสห้องทีมของห้อง {0}: {1}" -f $roomName, $groupId)
}

