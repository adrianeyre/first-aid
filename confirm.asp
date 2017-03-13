<%
' ****************************************************
' *                    confirm.asp                   *
' *                                                  *
' *            Coded by : Adrian Eyre                *
' *                Date : 05/11/2012                 *
' *             Version : 1.0.0                      *
' *                                                  *
' ****************************************************
%>
<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/FirstAidConnector.asp" -->

<style type="text/css">
<!--
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
}
body {
	margin-left: 0px;
	margin-top: 5px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {
	font-size: x-large;
	color: #FFFFFF;
}
-->
</style>
<%
Dim Info(13)
Info(1) = Request.Form("DateField")
Info(2) = Request.Form("TimeField")
Info(3) = Request.Form("StudentNameField")
Info(4) = Request.Form("FormField")
Info(5) = Request.Form("FirstAiderField")
Info(6) = Request.Form("InjuryField")
Info(7) = Request.Form("DescriptionField")
Info(8) = Request.Form("DepartmentField")
Info(9) = Request.Form("TreatmentField")
Info(10) = Request.Form("PostField")
Info(11) = Request.Form("LetterField")
Info(12) = Request.Form("OnlineField")
Info(13) = Request.Form("OtherField")

if Info(1) = "" then Info(1) = Date()
If Info(2) = "" then Info(2) = left(Time(),5)
If Info(3) = "" then Info(3) = "Student"
If Info(4) = "" then Info(4) = "None"
If Info(5) = 0 then Info(5) = 1
If Info(6) = 0 then Info(6) = 1
If Info(8) = 0 then Info(8) = 1
If Info(10) = 0 then Info(10) = 1
If Info(12) = 0 then Info(12) = 1

set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = MM_FirstAidConnector_STRING
Command1.CommandText = "INSERT INTO dbo.ResultsTable (Date, Time, StudentName, Form, FirstAider, InjuryType, InjuryDescription, Department, TreatmentGiven, PostTreatment, HeadBump, OnlineReport, Info)  VALUES ('" & Info(1) & "' , '" & info(2) & "' , '" & info(3) & "' , '" & info(4) & "' , '" & info(5) & "' , '" & info(6) & "' , '" & info(7) & "' , '" & info(8) & "' , '" & info(9) & "' , '" & info(10) & "' , '" & info(11) & "' , '" & info(12) & "' , '" & info(13) & "' ) "
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()

%>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="../images/backdefault.png" bgcolor="#192F68"><div align="center"><span class="style1">First Aid Reporting </span></div></td>
  </tr>
</table>
<table width="715" height="112" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><div align="center">Record Added </div></td>
  </tr>
  <tr>
    <td><div align="center">
      <form action="../main.asp?menu=staff/staff.asp" method="post" name="form1" target="_parent">
        <label>
        <input type="submit" name="Submit" value="Done">
        </label>
            </form>
      <label></label>
    </div></td>
  </tr>
</table>
