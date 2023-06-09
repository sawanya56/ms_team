# เชื่อมต่อกับ Microsoft Teams
Connect-MicrosoftTeams

# # กำหนด Group ID ของห้องทีมที่ต้องการลบ
$groupId = "9fdd0781-8f6a-4ed0-b2d2-05a77feef16b"

# Remove-Team -GroupId "5130617"

# # ลบห้องทีม
Remove-Team -GroupId $groupId









# # กำหนดชื่อห้องทีมที่ต้องการลบ
# $roomName = "2566/1 cs101 sec.1 ภาษาไทย"

# # เชื่อมต่อกับ Microsoft Teams
# Connect-MicrosoftTeams

# # ค้นหา Group ID ของห้องทีม
# $team = Get-Team | Where-Object { $_.DisplayName -eq $roomName }

# # ตรวจสอบว่าพบห้องทีมหรือไม่ก่อนที่จะลบ
# if ($team) {
#     # ลบห้องทีม
#     Remove-Team -GroupId $team.GroupId
# }
# else {
#     Write-Output "ไม่พบห้องทีมที่ต้องการลบ"
# }
