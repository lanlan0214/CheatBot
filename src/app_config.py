APP_VERSION = "1.0.2"

# 線上更新網址（支援兩種）
# 1) GitHub Releases latest API（推薦）
#    https://api.github.com/repos/<owner>/<repo>/releases/latest
#
# 2) 自訂 manifest JSON
#    https://raw.githubusercontent.com/<owner>/<repo>/main/latest.json
#    latest.json 內容需包含：
# {
#   "latest_version": "1.0.1",
#   "download_url": "https://.../桌面自動化工具.exe",
#   "notes": "修正左右切換與暫停邏輯"
# }
UPDATE_MANIFEST_URL = "https://api.github.com/repos/lanlan0214/CheatBot/releases/latest"
