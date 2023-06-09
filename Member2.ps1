Connect-MicrosoftTeams

$createRoomCsvPath = "D:\Ms Teams\test2.csv"
$addMembersCsvPath = "D:\Ms Teams\UserinRoom.csv"

# อ่านข้อมูลจากไฟล์สร้างห้อง
$createRoomData = Import-Csv -Path $createRoomCsvPath

foreach ($roomData in $createRoomData) {
    $year = $roomData."ปี"
    $term = $roomData."เทอม"
    $sec = $roomData."กลุ่ม"
    $subjectCode = $roomData."รหัสวิชา"
    $subjectName = $roomData."ชื่อวิชา"
    $startDate = $roomData."วันที่เริ่ม"
    $startTime = $roomData."เวลาเริ่ม"
    # $endDate = $roomData."วันที่สิ้นสุด"
    $endTime = $roomData."เวลาสิ้นสุด"

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

    # ตั้งค่าเวลาเริ่มและสิ้นสุดของห้องทีม
    $startDateTime = Get-Date -Date "$startDateTime"
    $endDateTime = Get-Date -Date "$endDate $endTime"

    # สร้างเหตุการณ์ (event) ของห้องทีม
    New-Event -GroupId $team.GroupId -Subject $roomName -StartDateTime $startDateTime -EndDateTime $endDateTime
    New-Event -GroupId $team.GroupId -Subject $roomName -StartDateTime $startDateTime -EndDateTime $endDateTime
}
