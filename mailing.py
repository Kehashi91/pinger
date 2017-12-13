import smtplib

fromaddr = "xxx"
toaddr = "xxx"
msg = "\r\n".join([
  "From: a.xxx",
  "To: xxx",
  "Subject: Just a message",
  "",
  "Why, oh why : D"
  ])

username = "xxx"
passwd = ""
print("oj1")
mailserver = smtplib.SMTP_SSL('xxx', 465, timeout=10)
print("oj1")
mailserver.ehlo()
print("oj2")
#mailserver.starttls()
print("oj3")
mailserver.login(username, passwd)
print("oj4")
mailserver.sendmail(fromaddr, toaddr, msg)
print("oj5")
mailserver.quit()

print(msg)