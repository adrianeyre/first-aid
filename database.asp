<%
' ****************************************************
' *                   database.asp                   *
' *                                                  *
' *            Coded by : Adrian Eyre                *
' *                Date : 05/11/2012                 *
' *             Version : 1.0.0                      *
' *                                                  *
' ****************************************************
%>
<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/FirstAidConnector.asp" -->
<%
Dim FirstAidersNames
Dim FirstAidersNames_numRows
Dim RecordNumber

Set FirstAidersNames = Server.CreateObject("ADODB.Recordset")
FirstAidersNames.ActiveConnection = MM_FirstAidConnector_STRING
FirstAidersNames.Source = "SELECT * FROM dbo.FirstAiders"
FirstAidersNames.CursorType = 0
FirstAidersNames.CursorLocation = 2
FirstAidersNames.LockType = 1
FirstAidersNames.Open()

FirstAidersNames_numRows = 0
%>
<%
Dim InjuryType
Dim InjuryType_numRows

Set InjuryType = Server.CreateObject("ADODB.Recordset")
InjuryType.ActiveConnection = MM_FirstAidConnector_STRING
InjuryType.Source = "SELECT * FROM dbo.InjuryType"
InjuryType.CursorType = 0
InjuryType.CursorLocation = 2
InjuryType.LockType = 1
InjuryType.Open()

InjuryType_numRows = 0
%>
<%
Dim Department
Dim Department_numRows

Set Department = Server.CreateObject("ADODB.Recordset")
Department.ActiveConnection = MM_FirstAidConnector_STRING
Department.Source = "SELECT * FROM dbo.Departments"
Department.CursorType = 0
Department.CursorLocation = 2
Department.LockType = 1
Department.Open()

Department_numRows = 0
%>
<%
Dim PostTreatment
Dim PostTreatment_numRows

Set PostTreatment = Server.CreateObject("ADODB.Recordset")
PostTreatment.ActiveConnection = MM_FirstAidConnector_STRING
PostTreatment.Source = "SELECT * FROM dbo.PostTreatment"
PostTreatment.CursorType = 0
PostTreatment.CursorLocation = 2
PostTreatment.LockType = 1
PostTreatment.Open()

PostTreatment_numRows = 0
%>
<%
Dim OnlineReport
Dim OnlineReport_numRows

Set OnlineReport = Server.CreateObject("ADODB.Recordset")
OnlineReport.ActiveConnection = MM_FirstAidConnector_STRING
OnlineReport.Source = "SELECT * FROM dbo.OnlineReport"
OnlineReport.CursorType = 0
OnlineReport.CursorLocation = 2
OnlineReport.LockType = 1
OnlineReport.Open()

OnlineReport_numRows = 0
%>
<%
Dim VisableUsers(100)
Dim VisableUsersData
Dim VisableUsers_numRows

Set VisableUsersData = Server.CreateObject("ADODB.Recordset")
VisableUsersData.ActiveConnection = MM_FirstAidConnector_STRING
VisableUsersData.Source = "SELECT * FROM dbo.VisableUsers"
VisableUsersData.CursorType = 0
VisableUsersData.CursorLocation = 2
VisableUsersData.LockType = 1
VisableUsersData.Open()

VisableUsers_numRows = 0
While (NOT VisableUsersData.EOF)
	VisableUsers_numRows = VisableUsers_numRows + 1
	VisableUsers(VisableUsers_numRows) = int(VisableUsersData.Fields.Item("VisableID").Value)
	VisableUsersData.MoveNext()
	' response.write(VisableUsers(VisableUsers_numRows)&" ")
Wend
%>
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
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="../images/backdefault.png" bgcolor="#192F68"><div align="center"><span class="style1">First Aid Reporting </span></div></td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><form name="form1" method="post" action="confirm.asp">
	  <table width="715" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="40">&nbsp;</td>
          <td width="215">&nbsp;</td>
          <td width="460">&nbsp;</td>
        </tr>
        <tr>
          <td width="40" height="20" bgcolor="#B0B0B0">&nbsp;</td>
          <td width="215" height="20" bgcolor="#B0B0B0">ID</td>
          <td width="460" height="20" bgcolor="#B0B0B0"><table width="100" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
            <tr>
              <td>New Record</td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#CCCCCC">&nbsp;</td>
          <td height="20" bgcolor="#CCCCCC">Date</td>
          <td height="20" bgcolor="#CCCCCC"><input name="DateField" type="text" id="DateField" value="<%response.write(date())%>" size="10" maxlength="10" /></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#B0B0B0">&nbsp;</td>
          <td height="20" bgcolor="#B0B0B0">Time</td>
          <td height="20" bgcolor="#B0B0B0"><input name="TimeField" type="text" id="TimeField" value="<%response.write(left(time(),5))%>" size="5" maxlength="5" /></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#CCCCCC">&nbsp;</td>
          <td height="20" bgcolor="#CCCCCC">Student Name </td>
          <td height="20" bgcolor="#CCCCCC"><label>
            <input name="StudentNameField" type="text" id="StudentNameField" size="70" maxlength="70">
          </label></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#B0B0B0">&nbsp;</td>
          <td height="20" bgcolor="#B0B0B0">Form</td>
          <td height="20" bgcolor="#B0B0B0"><input name="FormField" type="text" id="FormField" size="6" maxlength="6"></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#CCCCCC">&nbsp;</td>
          <td height="20" bgcolor="#CCCCCC">First Aider </td>
          <td height="20" bgcolor="#CCCCCC"><label>
            <select name="FirstAiderField" id="FirstAiderField">
              <option value="0">< Select ></option>
              <%
			  'response.write("IM HERE")
While (NOT FirstAidersNames.EOF)
	'response.write("IM HERE")
	for a = 1 to VisableUsers_numRows
		response.write("IM HERE")
		if VisableUsers(a) = int(FirstAidersNames.Fields.Item("ID").Value) then
			%><option value="<%=(FirstAidersNames.Fields.Item("ID").Value)%>"><%=(FirstAidersNames.Fields.Item("Name").Value)%></option><%
		end if
	next
	FirstAidersNames.MoveNext()
Wend
If (FirstAidersNames.CursorType > 0) Then
  FirstAidersNames.MoveFirst
Else
  FirstAidersNames.Requery
End If
%>
            </select>
          </label></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#B0B0B0">&nbsp;</td>
          <td height="20" bgcolor="#B0B0B0">Injury Type </td>
          <td height="20" bgcolor="#B0B0B0"><label>
            <select name="InjuryField" id="InjuryField">
              <option value="0">< Select ></option>
              <%
While (NOT InjuryType.EOF)
%><option value="<%=(InjuryType.Fields.Item("ID").Value)%>"><%=(InjuryType.Fields.Item("Name").Value)%></option>
              <%
  InjuryType.MoveNext()
Wend
If (InjuryType.CursorType > 0) Then
  InjuryType.MoveFirst
Else
  InjuryType.Requery
End If
%>
            </select>
          </label></td>
        </tr>
        <tr>
          <td height="100" bgcolor="#CCCCCC">&nbsp;</td>
          <td height="100" valign="top" bgcolor="#CCCCCC">Injury Description </td>
          <td height="100" bgcolor="#CCCCCC"><label>
            <textarea name="DescriptionField" cols="54" rows="5" id="DescriptionField"></textarea>
          </label></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#B0B0B0">&nbsp;</td>
          <td height="20" bgcolor="#B0B0B0">Department</td>
          <td height="20" bgcolor="#B0B0B0"><label>
            <select name="DepartmentField" id="DepartmentField">
              <option value="0">< Select ></option>
              <%
While (NOT Department.EOF)
%><option value="<%=(Department.Fields.Item("ID").Value)%>"><%=(Department.Fields.Item("Name").Value)%></option>
              <%
  Department.MoveNext()
Wend
If (Department.CursorType > 0) Then
  Department.MoveFirst
Else
  Department.Requery
End If
%>
            </select>
          </label></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#CCCCCC">&nbsp;</td>
          <td height="20" bgcolor="#CCCCCC">Treatment Given </td>
          <td height="20" bgcolor="#CCCCCC"><label>
            <input name="TreatmentField" type="text" id="TreatmentField" size="70" maxlength="70">
          </label></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#B0B0B0">&nbsp;</td>
          <td height="20" bgcolor="#B0B0B0">Post Treatment </td>
          <td height="20" bgcolor="#B0B0B0"><label>
            <select name="PostField" id="PostField">
              <option value="0">< Select ></option>
              <%
While (NOT PostTreatment.EOF)
%><option value="<%=(PostTreatment.Fields.Item("ID").Value)%>"><%=(PostTreatment.Fields.Item("Name").Value)%></option>
              <%
  PostTreatment.MoveNext()
Wend
If (PostTreatment.CursorType > 0) Then
  PostTreatment.MoveFirst
Else
  PostTreatment.Requery
End If
%>
            </select>
          </label></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#CCCCCC">&nbsp;</td>
          <td height="20" bgcolor="#CCCCCC">Head Bump Letter Given </td>
          <td height="20" bgcolor="#CCCCCC"><label>
            <input name="LetterField" type="checkbox" id="LetterField" value="checkbox">
          </label></td>
        </tr>
        <tr>
          <td height="20" bgcolor="#B0B0B0">&nbsp;</td>
          <td height="20" bgcolor="#B0B0B0">Online Report </td>
          <td height="20" bgcolor="#B0B0B0"><label>
            <select name="OnlineField" id="OnlineField">
              <option value="0">< Select ></option>
              <%
While (NOT OnlineReport.EOF)
%><option value="<%=(OnlineReport.Fields.Item("ID").Value)%>"><%=(OnlineReport.Fields.Item("Name").Value)%></option>
              <%
  OnlineReport.MoveNext()
Wend
If (OnlineReport.CursorType > 0) Then
  OnlineReport.MoveFirst
Else
  OnlineReport.Requery
End If
%>
            </select>
          </label></td>
        </tr>
        <tr>
          <td height="100" bgcolor="#CCCCCC">&nbsp;</td>
          <td height="100" valign="top" bgcolor="#CCCCCC">Other Information </td>
          <td height="100" bgcolor="#CCCCCC"><textarea name="OtherField" cols="54" rows="5" id="OtherField"></textarea></td>
        </tr>
        <tr>
          <td height="50" colspan="3"><div align="center">
            <label>
            <input type="submit" name="Submit" value="Add Record" />
            </label>
          </div></td>
          </tr>
      </table>

    </form>
    </td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="../images/backdefault.png" bgcolor="#192F68"><div align="center"><span class="style1">Search Database  </span></div></td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#CCCCCC"><form id="form2" name="form2" method="post" action="showdatabase.asp?recordnumber=1">
      <table width="715" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50" height="30" valign="middle">&nbsp;</td>
          <td width="123" height="30" valign="middle">Student Name </td>
          <td width="542" height="30" valign="middle"><label>
            <input name="SearchName" type="text" id="SearchName" size="60" maxlength="60" />
            <input type="submit" name="Submit" value="Submit" />
            </label></td>
        </tr>
        </table>
        </form>
    </td>
  </tr>
</table>
<br>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><form id="form3" name="form3" method="post" action="showdatabasereport.asp" target="_blank">
      <table width="715" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50" height="30" bgcolor="#CCCCCC">&nbsp;</td>
          <td width="123" height="30" bgcolor="#CCCCCC">Report from</td>
          <td width="100" height="30" bgcolor="#CCCCCC"><label for="FromDate"></label>
            <input name="FromDate" type="text" id="FromDate" value="<%=(date-30)%>" size="10" maxlength="10" /></td>
          <td width="25" height="30" align="center" bgcolor="#CCCCCC">to</td>
          <td width="100" height="30" bgcolor="#CCCCCC"><input name="ToDate" type="text" id="ToDate" value="<%=date%>" size="10" maxlength="10" /></td>
          <td width="317" height="30" bgcolor="#CCCCCC"><input type="submit" name="Submit2" id="Submit" value="Submit" /></td>
        </tr>
        </table>
      <table width="715" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td colspan="5" align="center" bgcolor="#CCCCCC">Show following Data</td>
          </tr>
        <tr>
          <td width="143" bgcolor="#CCCCCC"><input name="Option1" type="checkbox" id="Option1" checked="checked" />
            <label for="Option1">Date</label></td>
          <td width="143" bgcolor="#CCCCCC"><input type="checkbox" name="Option2" id="Option2" />
            <label for="Option2">Time</label></td>
          <td width="143" bgcolor="#CCCCCC"><input name="Option3" type="checkbox" id="Option3" checked="checked" />
            <label for="Option3">Student Name</label></td>
          <td width="143" bgcolor="#CCCCCC"><input type="checkbox" name="Option4" id="Option4" />
            Form</td>
          <td width="143" bgcolor="#CCCCCC"><input type="checkbox" name="Option5" id="Option5" />
            First Aider
</td>
        </tr>
        <tr>
          <td width="143" bgcolor="#CCCCCC"><input name="Option6" type="checkbox" id="Option6" checked="checked" />
Injury Type</td>
          <td width="143" bgcolor="#CCCCCC"><input type="checkbox" name="Option7" id="Option7" />
Description</td>
          <td width="143" bgcolor="#CCCCCC"><input name="Option8" type="checkbox" id="Option8" checked="checked" />
            Department
</td>
          <td width="143" bgcolor="#CCCCCC"><input name="Option9" type="checkbox" id="Option9" checked="checked" />
            Treatment
</td>
          <td width="143" bgcolor="#CCCCCC"><input type="checkbox" name="Option10" id="Option10" />
            Letter
</td>
        </tr>
        <tr>
          <td width="143" bgcolor="#CCCCCC"><input type="checkbox" name="Option11" id="Option11" />
Online Report</td>
          <td width="143" bgcolor="#CCCCCC"><input type="checkbox" name="Option12" id="Option12" />
Information </td>
          <td width="143" bgcolor="#CCCCCC">&nbsp;</td>
          <td width="143" bgcolor="#CCCCCC">&nbsp;</td>
          <td width="143" bgcolor="#CCCCCC">&nbsp;</td>
        </tr>
        <tr>
          <td width="143" bgcolor="#CCCCCC">&nbsp;</td>
          <td width="143" bgcolor="#CCCCCC">&nbsp;</td>
          <td width="143" bgcolor="#CCCCCC">&nbsp;</td>
          <td width="143" bgcolor="#CCCCCC">&nbsp;</td>
          <td width="143" bgcolor="#CCCCCC">&nbsp;</td>
        </tr>
      </table>
    </form></td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="40" bgcolor="#CCCCCC"><div align="center"><a href="../main.asp?page=55" target="_parent"><img src="../images/showbutton.png" alt="Show all records" width="140" height="29" border="0" /></a></div></td>
  </tr>
</table>
<p><a href="../main.asp?menu=/staff/showdatabase.asp" target="_parent"></a><a href="../main.asp?menu=/staff/showdatabase.asp" target="_parent"></a></p>
<p>&nbsp;</p>

<%
FirstAidersNames.Close()
Set FirstAidersNames = Nothing
%>
<%
InjuryType.Close()
Set InjuryType = Nothing
%>
<%
Department.Close()
Set Department = Nothing
%>
<%
PostTreatment.Close()
Set PostTreatment = Nothing
%>
<%
OnlineReport.Close()
Set OnlineReport = Nothing
%>
<%
VisableUsersData.Close()
Set VisableUsersData = Nothing
%>
