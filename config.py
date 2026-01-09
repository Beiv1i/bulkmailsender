# 邮件发送配置
# 部署说明: 请在部署后修改此文件，或直接在网页前端填入账号信息

# SMTP服务器地址 (Gmail企业版/个人版默认通用)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# 你的发送账号
SENDER_EMAIL = "your_email@gmail.com"

# 应用专用密码 (不是登录密码!)
# 获取方式: Google账户 -> 安全 -> 两步验证 -> 应用专用密码
APP_PASSWORD = "xxxx xxxx xxxx xxxx"

# 发送设置
# 这里的名称会显示在收件人的"发件人"一栏
SENDER_NAME = "IT Support"

# 邮件标题
EMAIL_SUBJECT = "Notification"

# 单次任务发送数量限制
# 设为 0 表示不限制（一次性发完所有）
BATCH_LIMIT = 50
