Connect-MicrosoftTeams

# กำหนดจำนวนห้องที่ต้องการสร้าง
$numberOfRooms = 3

# สร้างห้องโดยใช้ลูป
for ($i = 1; $i -le $numberOfRooms; $i++) {
    $teamName = "ห้องทีม $i"
    $teamDescription = "คำอธิบายห้องทีม $i"

    New-Team -DisplayName $teamName -Description $teamDescription
}
