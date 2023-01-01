import smtplib
userName = "ved.tiwari982@gmail.com"
passcode="********"
mssg= "hello"
to= "rams31824@gmail.com"
smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
smtpObj.starttls()
smtpObj.login(user=userName, password=passcode)
smtpObj.sendmail(userName, to, mssg)
print("Successfully sent email")
