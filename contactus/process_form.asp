<%@Language="VBScript" %>


<%

// CONFIGURATION SETTINGS FOR EMAIL AAAAAAAAAAHAHAHAHAHAA!!!!!!!
Const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
Const cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Const cdoSMTPConnectionTimeout="http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
Const cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Const cdoURLGetLatestVersion = "http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion"
Set cdoConfig = CreateObject("CDO.Configuration")
cdoConfig.Fields.Item(cdoSendUsingMethod) = 2
cdoConfig.Fields.Item(cdoSMTPServer) = "smtp-server.wi.rr.com"
cdoConfig.Fields.Item(cdoSMTPServerPort) = 25 '465 or 587 
cdoConfig.Fields.Item(cdoSMTPConnectionTimeout) = 20
cdoConfig.Fields.Item(cdoURLGetLatestVersion) = True
cdoConfig.Fields.Update
Set myMail=CreateObject("CDO.Message")
Set myMail.Configuration = cdoConfig


'''''''''''''''''''''''''EMAIL'''''''''''''''''''''''''''''''

body="Core Caseworks" & vbcrlf & vbcrlf

body=body&"--Customer Info--"&vbcrlf
body=body&"Name: " & request.querystring("customername") & vbcrlf
body=body&"Email: " & request.querystring("email") & vbcrlf
body=body&vbcrlf
body=body&"Profession: " & request.querystring("profession") & vbcrlf
body=body&vbcrlf
body=body&"--Comments--"&vbcrlf
body=body&request.querystring("comment")&vbcrlf
body=body&vbcrlf



myMail.From="bfelda@gmail.com"
myMail.To="ben@benfelda.com"

myMail.Subject="CoreCaseworks_Comment"
myMail.TextBody = body

myMail.Send


set myMail=nothing
set cdoConfig=nothing


%>
<script language="javascript">
	location.href="end.htm";
</script>

