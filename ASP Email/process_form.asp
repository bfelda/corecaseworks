<%@Language="VBScript" %>


<%

// CONFIGURATION SETTINGS FOR EMAIL AAAAAAAAAAHAHAHAHAHAA!!!!!!!
Const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
Const cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Const cdoSMTPConnectionTimeout="http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
Const cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Const cdoURLGetLatestVersion = "http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion"
Const cdoSendUserName = "http://schemas.microsoft.com/cdo/configuration/sendusername"
Const cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
Const cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
Const cdoSMTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/smtpusessl"

Set cdoConfig = CreateObject("CDO.Configuration")
cdoConfig.Fields.Item(cdoSendUsingMethod) = 2
cdoConfig.Fields.Item(cdoSMTPServer) = "smtp.gmail.com"
cdoConfig.Fields.Item(cdoSMTPServerPort) = 465 '465 or 587 
cdoConfig.Fields.Item(cdoSMTPConnectionTimeout) = 10
cdoConfig.Fields.Item(cdoURLGetLatestVersion) = True
cdoConfig.Fields.Item(cdoSMTPAuthenticate) = 1
cdoConfig.Fields.Item(cdoSMTPUseSSL) = True
cdoConfig.Fields.Item(cdoSendUserName) = "<INSERT EMAIL ADDRESS>"
cdoConfig.Fields.Item(cdoSendPassword) = "<INSERT EMAIL PASSWORD>"
cdoConfig.Fields.Update
Set myMail=CreateObject("CDO.Message")
Set myMail.Configuration = cdoConfig


dim phone_num,oncall_email
		set fs=server.CreateObject("Scripting.FileSystemObject")
		set f=fs.OpenTextFile("c:\phone\phone.txt",1)
		phone_num = f.readline
		oncall_email = f.readline
		f.close
		set f=nothing
		set fs=nothing
from_address="RFT.Email@gmail.com"
'to_address="technicians-internal@rft.com"
to_address="cstcustomerescalations@rft.com"
'to_address="dyatzeck@rft.com"
	today_=weekday(Date())
	now_=hour(Time())
	if (today_="1" OR today_="7") OR (now_<7 OR now_>18) AND (len(phone_num) > 4) then
cc_address=Right(phone_num,10)&"@vtext.com"	
to_address=to_address&"; "&oncall_email
	end if
'cc_address=Right(phone_num,10)&"@vtext.com"
'to_address="sberman@rft.com"
'cc_address=Right(phone_num,10)&"@vtext.com"
'to_address=to_address&"; "&oncall_email
bcc_address="dyatzeck@rft.com"



'''''''''''''''''''''''''EMAIL'''''''''''''''''''''''''''''''

body="Offline OnContact Submission" & vbcrlf & vbcrlf

body=body&"Request Submitted by " & Request.QueryString("submitted_name") & vbcrlf
body=body&vbcrlf
body=body&vbcrlf

body=body&"--Customer Info--"&vbcrlf
body=body&"Customer Name: " & request.querystring("cust_name") & vbcrlf
body=body&"City: " & request.querystring("cust_city") & vbcrlf
body=body&"State: " & request.querystring("cust_state") & vbcrlf
body=body&"Contact Name: " & request.querystring("first_name") & " " & request.querystring("last_name") & vbcrlf
body=body&"Department: " & request.querystring("department_") & vbcrlf
body=body&"Phone: " & request.querystring("phone_") & vbcrlf
body=body&"Fax: " & request.querystring("fax_") & vbcrlf
body=body&vbcrlf

body=body&"--System Type--"&vbcrlf
if (request.querystring("sys_qr-nc")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Quick Response/Nurse Call"&vbcrlf
if (request.querystring("sys_sp")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Safe Place"&vbcrlf
if (request.querystring("sys_wander")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Wanderer Monitoring"&vbcrlf
body=body&"Software Version: " & request.querystring("sys_ver") & vbcrlf
body=body&vbcrlf

body=body&"--Component--"&vbcrlf
if (request.querystring("c_doorcontroller")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Door Controller"&vbcrlf
if (request.querystring("c_doorreceiver")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Door Receiver"&vbcrlf
if (request.querystring("c_doorlock")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Door Lock"&vbcrlf
if (request.querystring("c_computer")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Computer"&vbcrlf
if (request.querystring("c_transmitter")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Transmitter"&vbcrlf
if (request.querystring("c_pullcord")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Pull Cord"&vbcrlf
if (request.querystring("c_pendant")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Pendant"&vbcrlf
if (request.querystring("c_powersupply")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Power Supply"&vbcrlf
if (request.querystring("c_backupinterface")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Backup Interface"&vbcrlf
if (request.querystring("c_pagingbase")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Paging Base/Pager"&vbcrlf
body=body&"Other: " & request.querystring("c_other") & vbcrlf
body=body&vbcrlf

body=body&"--Description--"&vbcrlf
body=body&request.querystring("description_")&vbcrlf
body=body&vbcrlf

body=body&"--RFT Actions--"&vbcrlf
if (request.querystring("a_logincident")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Log Incident"&vbcrlf
if (request.querystring("a_contactcustomer")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Contact Customer"&vbcrlf
if (request.querystring("a_oncallcontacted")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"On Call Contacted"&vbcrlf
if (request.querystring("a_issueresolved")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Issue Resolved"&vbcrlf
if (request.querystring("a_partsrequired")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Parts Required"&vbcrlf
if (request.querystring("a_servicerequired")="on") then
body=body&"[x] "
else
body=body&"[ ] "
end if
body=body&"Service Required"&vbcrlf
body=body&"Other: " & request.querystring("a_other") & vbcrlf
body=body&vbcrlf


myMail.From=from_address
myMail.To=to_address
'myMail.CC=cc_address 'NO LONG EMAIL TO TXT
myMail.BCC=bcc_address
myMail.Subject="OnContact Log Sheet"
myMail.TextBody = body
'myMail.TextBody="This is a test.  Please disregard."

myMail.Send




'''''''''''''''''''''''''TXT MESSAGE'''''''''''''''''''''''''''''''

body=""
body=body&"CUST-"& request.querystring("cust_name") & vbcrlf
body=body&"NAM-" & request.querystring("first_name") & " " & request.querystring("last_name") & vbcrlf
body=body&"PH-" & request.querystring("phone_") & vbcrlf
body=body&"SYS-"
if (request.querystring("sys_qr-nc")="on") then
body=body&"QR "
end if
if (request.querystring("sys_sp")="on") then
body=body&"SP "
end if
if (request.querystring("sys_wander")="on") then
body=body&"Wnder "
end if
body=body&"(" & request.querystring("sys_ver") & ")" & vbcrlf

body=body&"PART-"
if (request.querystring("c_doorcontroller")="on") then
body=body&"DorCtrl "
end if
if (request.querystring("c_doorreceiver")="on") then
body=body&"DorRcvr "
end if
if (request.querystring("c_doorlock")="on") then
body=body&"DorLck "
end if
if (request.querystring("c_computer")="on") then
body=body&"Cmptr "
end if
if (request.querystring("c_transmitter")="on") then
body=body&"Tx "
end if
if (request.querystring("c_pullcord")="on") then
body=body&"PulCrd "
end if
if (request.querystring("c_pendant")="on") then
body=body&"Pndnt "
end if
if (request.querystring("c_powersupply")="on") then
body=body&"PwrSply "
end if
if (request.querystring("c_backupinterface")="on") then
body=body&"BkpInt "
end if
if (request.querystring("c_pagingbase")="on") then
body=body&"PgngBse "
end if
if (Request.QueryString("c_other")>"") then
body=body& request.querystring("c_other")
end if
body=body&vbcrlf

body=body&"ACT-"
if (request.querystring("a_logincident")="on") then
body=body&"LogIncdnt "
end if
if (request.querystring("a_contactcustomer")="on") then
body=body&"CallCust "
end if
if (request.querystring("a_oncallcontacted")="on") then
body=body&"Tech "
end if
if (request.querystring("a_issueresolved")="on") then
body=body&"None "
end if
if (request.querystring("a_partsrequired")="on") then
body=body&"PartsReqd "
end if
if (request.querystring("a_servicerequired")="on") then
body=body&"SrviceReqd "
end if
if (Request.QueryString("c_other")>"") then
body=body& request.querystring("c_other")
end if
body=body&vbcrlf

body=body&"DESC-"
body=body&request.querystring("description_")&vbcrlf
body=body&vbcrlf


myMail.From=from_address
myMail.To=""
myMail.CC=cc_address
myMail.BCC=bcc_address
myMail.Subject="ONC Log"
myMail.TextBody = body

myMail.Send


set myMail=nothing
set cdoConfig=nothing


%>
<script language="javascript">
	location.href="end.htm";
</script>
<!--<h2>Submission Complete!</h2><br><%=from_address%>
Click <a href='index.htm'>here</a> to enter another submission, otherwise you may close your browser.
-->

