import smtplib
import pandas as pd
from email.message import EmailMessage
from datetime import datetime

EMAIL='YOUR_EMAIL_ADDRESS'
EMAIL_PASS='YOUR_APP_PASSWORD'


data = pd.read_excel (r'DATA.xlsx')
df = pd.DataFrame(data, columns= ['EMAIL','DOB'])
mylist = df.values.tolist()
print("Logging In ...... \n")

smtp= smtplib.SMTP_SSL('smtp.gmail.com', 465)
smtp.login(EMAIL, EMAIL_PASS)

for da in mylist:
    if datetime.date(da[1]).day==datetime.date(datetime.now()).day and datetime.date(da[1]).month==datetime.date(datetime.now()).month :
        print('Birthday wishes to:\n')
        msg = EmailMessage()
        msg['Subject'] = 'Happy Birthday!'
        msg['From'] = EMAIL
        msg['To'] = da
        print(da[0])
        msg.add_alternative("""\
       <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Demystifying Email Design</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
</head>
<body style="margin: 0; padding: 0;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%"> 
        <tr>
            <td style="padding: 10px 0 30px 0;">
                <table align="center" border="0" cellpadding="0" cellspacing="0" width="600" style="border: 1px solid #cccccc; border-collapse: collapse;">
                    <tr>
                        <td align="center" bgcolor="#53BFFA" style="padding: 10px 0 10px 0; color: #153643; font-size: 28px; font-weight: bold; font-family: Arial, sans-serif;">
                            <img src="https://i.ibb.co/5kTBqbP/q2.jpg" alt="Creating Email Magic" width="550" height="330" style="display: block;" />
                        </td>
                    </tr>
                    <tr>
                        <td bgcolor="#ffffff" style="padding: 40px 30px 40px 30px;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td style="color: #153643; font-family: Arial, sans-serif; font-size: 24px;">
                                        <center><b>Happy Birthday!<br/></b></center>
                                    </td>
                                </tr>
                                <tr>
                                 </tr>
                                <tr>
                                   <tr>
                                    <td style="padding: 20px 0 30px 0; color: #153643; font-family: Arial, sans-serif; font-size: 16px; line-height: 20px;">
                                        Another year older, and you just keep getting stronger, wiser, funnier and more amazing!
                                        Warmest wishes and love on your birthday and always!
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td bgcolor="#ee4c50" style="padding: 30px 30px 30px 30px;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td style="color: #ffffff; font-family: Arial, sans-serif; font-size: 14px;" width="75%">
                                        &reg; Souvik , somewhere 2020<br/>
                                        <a href="#" style="color: #ffffff;"><font color="#ffffff">Unsubscribe</font></a> to this newsletter instantly Send me a email 
                                    </td>
                                    <td align="right" width="25%">
                                        <table border="0" cellpadding="0" cellspacing="0">
                                            <tr>
                                                <td style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold;">
                                                    
                                                    
                                                    <a href="https://www.instagram.com/Data.Pirates/" style="color: #ffffff;">
                                                        <img src="https://i.ibb.co/6BhKGyq/insta.jpg" alt="Twitter" width="38" height="38" style="display: block;" border="0" />
                                                    </a>
                                                </td>
                                                <td style="font-size: 0; line-height: 0;" width="20">&nbsp;</td>
                                                <td style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold;">
                                                    <a href="https://www.facebook.com/profile.php?id=100037514408308" style="color: #ffffff;">
                                                        <img src="https://i.ibb.co/4ZgZNQd/images.png" alt="Facebook" width="38" height="38" style="display: block;" border="0" />
                                                    </a>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
             """, subtype='html')
        smtp.send_message(msg)

    if 25==datetime.date(datetime.now()).day and 12==datetime.date(datetime.now()).month :
        print('Christmas wishes to:\n')
        msg = EmailMessage()
        msg['Subject'] = 'Merry Christmas'
        msg['From'] = EMAIL
        print(da[0])
        msg['To'] = da
        msg.add_alternative("""\
            <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Demystifying Email Design</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
</head>
<body style="margin: 0; padding: 0;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%"> 
        <tr>
            <td style="padding: 10px 0 30px 0;">
                <table align="center" border="0" cellpadding="0" cellspacing="0" width="600" style="border: 1px solid #cccccc; border-collapse: collapse;">
                    <tr>
                        <td align="center" bgcolor="#53BFFA" style="padding: 10px 0 10px 0; color: #153643; font-size: 28px; font-weight: bold; font-family: Arial, sans-serif;">
                            <img src="https://i.ibb.co/xM6Y474/merry-christmas-calligraphy-with-baubles-1262-7024.jpg" alt="Creating Email Magic" width="550" height="330" style="display: block;" />
                        </td>
                    </tr>
                    <tr>
                        <td bgcolor="#ffffff" style="padding: 40px 30px 40px 30px;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td style="color: #153643; font-family: Arial, sans-serif; font-size: 24px;">
                                        <center><b>Merry Christmas!<br/></b></center>
                                    </td>
                                </tr>
                                <center>Jingle bells, jingle bells<br/>
Jingle all the way!<br/>
Oh what fun it is to ride in a one-horse open sleigh</center>
                                <tr>
                                 </tr>
                                <tr>
                                   <tr>
                                    <td style="padding: 20px 0 30px 0; color: #153643; font-family: Arial, sans-serif; font-size: 16px; line-height: 20px;">
                                      
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td bgcolor="#ee4c50" style="padding: 30px 30px 30px 30px;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td style="color: #ffffff; font-family: Arial, sans-serif; font-size: 14px;" width="75%">
                                        &reg; Souvik , somewhere 2020<br/>
                                        <a href="#" style="color: #ffffff;"><font color="#ffffff">Unsubscribe</font></a> to this newsletter instantly Send me a email 
                                    </td>
                                    <td align="right" width="25%">
                                        <table border="0" cellpadding="0" cellspacing="0">
                                            <tr>
                                                <td style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold;">
                                                    
                                                    
                                                    <a href="https://www.instagram.com/Data.Pirates/" style="color: #ffffff;">
                                                        <img src="https://i.ibb.co/6BhKGyq/insta.jpg" alt="Twitter" width="38" height="38" style="display: block;" border="0" />
                                                    </a>
                                                </td>
                                                <td style="font-size: 0; line-height: 0;" width="20">&nbsp;</td>
                                                <td style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold;">
                                                    <a href="https://www.facebook.com/profile.php?id=100037514408308" style="color: #ffffff;">
                                                        <img src="https://i.ibb.co/4ZgZNQd/images.png" alt="Facebook" width="38" height="38" style="display: block;" border="0" />
                                                    </a>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
             """, subtype='html')
        smtp.send_message(msg)

    if 1==datetime.date(datetime.now()).day and 1==datetime.date(datetime.now()).month :
        print('New Year wishes to:\n')
        msg = EmailMessage()
        msg['Subject'] = 'Happy New Year'
        msg['From'] = EMAIL
        print(da[0])
        msg['To'] = da
           
          

        msg.add_alternative("""\
            <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Demystifying Email Design</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
</head>
<body style="margin: 0; padding: 0;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%"> 
        <tr>
            <td style="padding: 10px 0 30px 0;">
                <table align="center" border="0" cellpadding="0" cellspacing="0" width="600" style="border: 1px solid #cccccc; border-collapse: collapse;">
                    <tr>
                        <td align="center" bgcolor="#53BFFA" style="padding: 10px 0 10px 0; color: #153643; font-size: 28px; font-weight: bold; font-family: Arial, sans-serif;">
                            <img src=" https://i.ibb.co/CzXmRZG/fa3602d610f99d364a5e0d06acf9d58f.jpg" alt="Creating Email Magic" width="550" height="330" style="display: block;" />
                        </td>
                    </tr>
                    <tr>
                        <td bgcolor="#ffffff" style="padding: 40px 30px 40px 30px;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td style="color: #153643; font-family: Arial, sans-serif; font-size: 24px;">
                                        <center><b>Happy New Year!<br/></b></center>
                                    </td>
                                </tr>
                                <tr>
                                Wishing you a Happy New Year with the hope that you will have many blessings in the year to come.
                                 </tr>
                                <tr>
                                   <tr>
                                    <td style="padding: 20px 0 30px 0; color: #153643; font-family: Arial, sans-serif; font-size: 16px; line-height: 20px;">
                                      
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td bgcolor="#ee4c50" style="padding: 30px 30px 30px 30px;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td style="color: #ffffff; font-family: Arial, sans-serif; font-size: 14px;" width="75%">
                                        &reg; Souvik , somewhere 2020<br/>
                                        <a href="#" style="color: #ffffff;"><font color="#ffffff">Unsubscribe</font></a> to this newsletter instantly Send me a email 
                                    </td>
                                    <td align="right" width="25%">
                                        <table border="0" cellpadding="0" cellspacing="0">
                                            <tr>
                                                <td style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold;">
                                                    
                                                    
                                                    <a href="https://www.instagram.com/Data.Pirates/" style="color: #ffffff;">
                                                        <img src="https://i.ibb.co/6BhKGyq/insta.jpg" alt="Twitter" width="38" height="38" style="display: block;" border="0" />
                                                    </a>
                                                </td>
                                                <td style="font-size: 0; line-height: 0;" width="20">&nbsp;</td>
                                                <td style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold;">
                                                    <a href="https://www.facebook.com/profile.php?id=100037514408308" style="color: #ffffff;">
                                                        <img src="https://i.ibb.co/4ZgZNQd/images.png" alt="Facebook" width="38" height="38" style="display: block;" border="0" />
                                                    </a>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
             """, subtype='html')
        smtp.send_message(msg)
