<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Server.ScriptTimeOut=5000 %>
<!--#include FILE="upload.inc"--> 
<%
myemail = "contact@nlgservices.co.nz"
myname = "New Line Group"
smtpserver = "mail.weborigin.co.nz"
smtpuser = "nlg@weborigin.co.nz"
smtppwd = "qwertyuiop"
smtpemail = "nlg@weborigin.co.nz"

dim upload,file,formName,formPath 
set upload=new upload_5xSoft

CName = upload.form("Name")
CMail = upload.form("Email")
CCompany = upload.form("Company")
CPhone = upload.form("Phone")
CEnquiry = upload.form("Enquiry")
CServiceCategory = upload.form("ServiceCategory")
CNewsletter = upload.form("Newsletter")

set jmail=server.CreateObject ("jmail.message")

jmail.From = "nlg@weborigin.co.nz"
jmail.FromName = CName
jmail.ReplyTo = CMail
jmail.Subject = "Web Customer: " & CName 

jmail.Body = "Name: " & CName & vbcrlf & vbcrlf & "Company: " & CCompany & vbcrlf & vbcrlf & "E-mail: " & CMail & vbcrlf & vbcrlf & "Phone: " & CPhone & vbcrlf & vbcrlf & "Service Category: " & CServiceCategory & vbcrlf & vbcrlf & "Newsletter: " & CNewsletter & vbcrlf & vbcrlf & "Enquiry:" & vbcrlf & CEnquiry

jmail.AddRecipient myemail,myname
jmail.MailServerUserName = smtpuser
jmail.MailServerPassWord = smtppwd

isgo=jmail.Send(smtpserver)

Response.Redirect("http://www.nlgservices.co.nz/Respond/")

jmail.Close
set jmail=nothing
%>