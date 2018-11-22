# Recurse AD looking for passwords which will expire in the next 7 days. If found email the user a reminder every day until they change their PW.
# v1.0
#
# Next steps: Add logging, Capture and log any errors

MAXPASSWORDAGEINDAYS = 45

import smtplib
from email.message import EmailMessage
from datetime import date, datetime, timedelta
from pyad import *
import logging

# Prep for running the program
filename = datetime.today().strftime("%Y%m%d") + ".log"
logging.basicConfig(filename=filename, filemode='a', format='%(asctime)s - %(funcName)s - %(levelname)s - %(message)s', level=logging.INFO)
logging.info(__file__ + " is starting.\n")

# Make a fancy message
def makeMessage(humanoid, xDays, xDate):
    msg = """    
<html>
    <head>
    <style><!--
    /* Font Definitions */
    @font-face
        {font-family:Wingdings;
        panose-1:5 0 0 0 0 0 0 0 0 0;}
    @font-face
        {font-family:"Cambria Math";
        panose-1:2 4 5 3 5 4 6 3 2 4;}
    @font-face
        {font-family:Calibri;
        panose-1:2 15 5 2 2 2 4 3 2 4;}
    @font-face
        {font-family:Tahoma;
        panose-1:2 11 6 4 3 5 4 4 2 4;}
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
        {margin:0in;
        margin-bottom:.0001pt;
        font-size:11.0pt;
        font-family:"Calibri","sans-serif";}
    a:link, span.MsoHyperlink
        {mso-style-priority:99;
        color:blue;
        text-decoration:underline;}
    a:visited, span.MsoHyperlinkFollowed
        {mso-style-priority:99;
        color:purple;
        text-decoration:underline;}
    span.EmailStyle17
        {mso-style-type:personal-compose;
        font-family:"Calibri","sans-serif";
        color:windowtext;
        font-weight:normal;
        font-style:normal;}
    .MsoChpDefault
        {mso-style-type:export-only;
        font-family:"Calibri","sans-serif";}
    @page WordSection1
        {size:8.5in 11.0in;
        margin:1.0in 1.0in 1.0in 1.0in;}
    div.WordSection1
        {page:WordSection1;}
    --></style>
    </head>
	<body lang=EN-US link=blue vlink=purple>
		<div class=WordSection1>
			<p class=MsoNormal>
				<span style='color:#1F497D'>Greetings {humanoid} </span>
				<o:p></o:p>
			</p>
			<p class=MsoNormal><span style='color:#1F497D'>
				<o:p>&nbsp;</o:p></span></p>
			<p class=MsoNormal>
				Your password is expiring in the next {xDays} days, please change it before the expiration date of {xDate} <o:p></o:p>
			</p>
			<p class=MsoNormal>If your password expires, your supervisor will have to request I.T. to reset your password.<o:p></o:p>
			</p>
			<p class=MsoNormal>
				<o:p>&nbsp;</o:p>
			</p>
			<p class=MsoNormal style='margin-left:.5in'>
				<b><span style='color:#0000CC'>To change your password, press CTLR + ALT + DEL -></span></b>
				<span style='color:#0000CC'> Change a password <o:p></o:p></span>
				</b>
			</p>
			<p class=MsoNormal style='margin-left:.5in'>
				<b><span style='color:#0000CC'>Enter your current password and a new one in the two boxes and click OK<o:p></o:p></span></b>
			</p>
			<p class=MsoNormal style='margin-left:.5in'><o:p>&nbsp;</o:p>
			</p>
			<p class=MsoNormal style='margin-left:.5in'>
				<span style='color:#3333CC'>The notice of your password expiring is seen when you are logging onto your PC, <o:p></o:p></span>
			</p>
			<p class=MsoNormal style='margin-left:.5in'>
				<span style='color:#3333CC'>If you have not logged out in some time that would be why you are not seeing the notice.<o:p></o:p></span>
			</p>
			<p class=MsoNormal style='margin-left:.5in'>
				<span style='color:#3333CC'><o:p>&nbsp;</o:p></span>
			</p>
			<p class=MsoNormal style='margin-left:.5in'>
				<span style='color:#3333CC'>It is recommended to reboot your PC at least once a week to refresh your system, <o:p></o:p></span>
			</p>
			<p class=MsoNormal style='text-indent:.5in'>
				<span style='color:#3333CC'>write your profile back to the server and see these type of notices.<o:p></o:p></span>
			</p>
			<p class=MsoNormal>
				<o:p>&nbsp;</o:p>
			</p>
			<p class=MsoNormal style='text-autospace:none'>
				<b><span style='font-size:10.0pt;font-family:"Tahoma","sans-serif";color:#000099'>Thank you,<o:p></o:p></span></b>
			</p>
			<p class=MsoNormal style='text-autospace:none'>
				<b><span style='font-size:10.0pt;font-family:"Tahoma","sans-serif";color:#000099'>The I.T. Department</span></b><span style='font-size:8.0pt;font-family:"Tahoma","sans-serif";color:#0000CC'><o:p></o:p></span>
			</p>
			<p class=MsoNormal>
				<o:p>&nbsp;</o:p>
			</p>
		</div>
	</body>
</html>
"""
    logging.info("Human:" + humanoid + " Password will expire in " + str(xDays) + " days on " + str(xDate))
    msg = msg.replace("{humanoid}", humanoid)
    msg = msg.replace("{xDays}", str(xDays))
    msg = msg.replace("{xDate}", str(xDate))
    return msg;
    
# Send the message via our own SMTP server.
def mailMsg(address, msg, days, humanoid):
    logging.info("eMailing PW expiration warning message.\n")
    msg = "Content-Type: text/html;\nMime-Version: 1.0;\nTo: klacrosse@pkwillis.com\nSubject: "+ humanoid + ", your password will expire in " + str(days) + " days\n\n" + msg
    s = smtplib.SMTP('10.0.0.69')
    s.set_debuglevel(0)
    s.sendmail("pkw-reminders@pkwillis.com", "klacrosse@pkwillis.com", msg)
    s.quit()
#    logging.info(msg)
    
# MAIN: Query AD for delinquent passwords
def main():
    q = adquery.ADQuery()

    q.execute_query(
        attributes = ["distinguishedName", "displayName", "memberOf", "mail"],
        where_clause = "objectClass = 'User'",
        base_dn = "OU=PKW Corp,DC=pkwillis,DC=local"
    )
    
    for row in q.get_results():
        logging.info ("Inspecting " + row["distinguishedName"])
        user1 = aduser.ADUser.from_dn(row["distinguishedName"])
        address = row["mail"] 
        pwLastSet = user1.get_password_last_set()
        pwAge = (datetime.today() - pwLastSet).days
        uctl = user1.get_user_account_control_settings()
        pwDoesntExpire = uctl["DONT_EXPIRE_PASSWD"]
        badOU = "Computers" in row["distinguishedName"] or "SecuredProfiles" in row["distinguishedName"] or "Security Groups" in row["distinguishedName"] or "Training" in row["distinguishedName"]

        if badOU == False and pwDoesntExpire == False and pwAge > MAXPASSWORDAGEINDAYS-7:
           pwExpirationDate = (pwLastSet + timedelta(MAXPASSWORDAGEINDAYS)).strftime("%m/%d/%Y")
           logging.info ("Preparing to email warning to " + row["displayName"] + ": PW is " + str(pwAge) + " days old. PW_Doesn't_Expire flag is " + str(pwDoesntExpire) + " and Expires on " + str(pwExpirationDate))
           msg = makeMessage(row["displayName"], MAXPASSWORDAGEINDAYS-pwAge, pwExpirationDate)
           mailMsg(address, msg, MAXPASSWORDAGEINDAYS-pwAge, row["displayName"])
    
main()

logging.info(__file__ + " completed.\n")
