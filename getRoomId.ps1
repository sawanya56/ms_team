Connect-MicrosoftTeams

# ค้นหาทีมจากชื่อห้อง
Get-Team | Where-Object { $_.DisplayName -eq "ห้องทีม 1" } | Select-Object GroupId