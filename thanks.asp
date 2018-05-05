<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>MOSERS - Thanks for Signing up!</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<link href="stylesheet.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {
	color: #800302;
	font-weight: bold;
}
.style2 {
	font-size: 0.8em;
	font-style: italic;
}
.style5 {font-weight: bold}
.style6 {font-size: 1.2em}
.style7 {font-size: 0.8em}
-->
</style>
</head>

<body ><table align="center" cellpadding="12" id="main">
  <tr>
    <td width="278"><a href="http://www.dot.state.tx.us/" target="_blank"><img src="images/txdot.sm.gif" width="278" height="30" border="0"></a></td>
    <td width="426"><div align="right"><a href="http://tti.tamu.edu/" target="_blank"><img src="images/ttilogo.tranparent.gif" width="163" height="36" border="0"></a></div></td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td>
    <blockquote><h1>MOSERS</h1></blockquote></td>
    <td><h3 align="right">The Texas Guide to Accepted<br>
    Mobile Source Emission Reduction Strategies</h3></td>
  </tr>
  <tr>
    <td colspan="2"><br>
      <table width="600" border="0" align="center" cellpadding="7" cellspacing="0" id="sched" summary="MOSERS Training Schedule">
        <tr>
          <td width="558" class="head"><div align="center">Thanks for Signing up! </div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td colspan="2"><blockquote><a href="index.html">Back to Home Page</a></blockquote>
      <blockquote class="style2"><br>
    For assistance or questions, please call 817-277-5503 or contact us by <span class="style1" id="alt"><a href="mailto:mosers@tamu.edu">email</a></span>. </blockquote></td>
  </tr>
</table>
<%
'Send Out Email
Set objMailer = Server.CreateObject ("Dundas.Mailer")

from = "mosers@ttimail.tamu.edu"
email_to = "mosers@tamu.edu"
subject = "MOSERS Email News Sign-Up"
body = "The following person would like to sign up! " & request.form("name") & vbCrLf & vbCrLf & "Name: " & request.form("name")  & vbCrLf & "Email Address: " & request.form("address") & vbCrLf & "Agency: " & request.form("agency") & vbCrLf & "City: " & request.form("city") & vbCrLf & "State: " & request.form("state") & vbCrLf

sReturn = objMailer.QuickSend("" & from & "","" & email_to & "",""& subject & "","" & body & "")

Set objMailer = Nothing
%>
</body>
</html>
