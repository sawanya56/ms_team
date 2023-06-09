Connect-MicrosoftTeams

$createRoomCsvPath = "D:\Ms Teams\test1.csv"
$addMembersCsvPath = "D:\Ms Teams\UserinRoom.csv"

# อ่านข้อมูลจากไฟล์สร้างห้อง
$createRoomData = Import-Csv -Path $createRoomCsvPath

foreach ($roomData in $createRoomData) {
    $year = $roomData."ปี"
    $term = $roomData."เทอม"
    $sec = $roomData."กลุ่ม"
    $subjectCode = $roomData."รหัสวิชา"
    $subjectName = $roomData."ชื่อวิชา"

    $roomName = $year + "/" + $term + " " + $subjectCode + " " + $sec + " " + $subjectName
    $teamDescription = $subjectCode + "-" + $subjectName

    # สร้างห้องทีม
    $team = New-Team -DisplayName $roomName -Description $teamDescription

    # # แสดงค่า groupId โดยเฉพาะตัวเลขและจำกัดจำนวนหลักไม่เกิน 7 หลัก
    # $groupId = $team.GroupId -replace "[^0-9]", "" -replace "^(\d{7}).*", '$1'
    # Write-Output ("รหัสห้องทีมของห้อง {0}: {1}" -f $roomName, $groupId)
    
    # แสดงค่า groupId
    $groupId = $team.GroupId
    Write-Output ("รหัสห้องทีมของห้อง {0}: {1}" -f $roomName, $groupId)

    # อ่านข้อมูลจากไฟล์เพิ่มสมาชิก
    $addMembersData = Import-Csv -Path $addMembersCsvPath

    foreach ($memberData in $addMembersData) {
        $email = $memberData.Email
        $role = $memberData.Role

        # เพิ่มสมาชิกลงในห้องทีม
        Add-TeamUser -GroupId $team.GroupId -User $email -Role $role
    }
}
