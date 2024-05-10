<%@ Language=VBScript %>
<%Response.Expires = 0%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<Title>Edit Guest Task Report Templates</title>
<Script ID=clientEventHandlersVBS Language=vbscript>
<!--
  Sub lstCompany_onclick
      window.form1.cmdEdit.disabled = false
  End sub
  
  Sub lstCompany_onDblClick
      cmdEdit_onclick
  End sub
  
  sub cmdEdit_onclick
      window.parent.location.href = "EditCompanyTemp.asp?ID=" & window.form1.lstCompany.value
  End Sub
  
  sub cmdCancel_onclick
      window.parent.location.href = "Administration.asp"
  end sub
-->
</Script>

</HEAD>
<BODY bgcolor=silver topmargin="0" leftmargin="4" marginwidth="0" marginheight="0" link="black" vlink="black" alink="black">
<!--#include file = "Header.inc" -->
<form id=form1 name=form1>
<div align="left">
  <table border="0" width="750">
    <tr>
      <td>
		<SELECT size="20" id=lstCompany name=lstCompany style="width: 750; position: relative; float: left">
	     <%
	       Set cnSQL = Server.CreateObject("ADODB.Connection")
		   Set rsCompanies = Server.CreateObject("ADODB.Recordset")
		   
		   cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")
	       Set rsCompanies = cnSQL.Execute("Select tblCompany.CompanyID, tblCompany.CompanyName from tblCompany Order by tblCompany.CompanyName")
	       
	       Do While Not rsCompanies.EOF
				Response.Write "<OPTION value=" & rsCompanies.Fields("CompanyID") & ">" & rsCompanies.Fields("CompanyName") & "</Option>"
				rsCompanies.MoveNext
			Loop
			
			Response.Write "<BR><BR>"
	     
	     %>
</SELECT>
	  </td>
    </tr>
    <tr>
      <td>
		<INPUT type="button" value="Edit Template" id=cmdEdit name=cmdEdit disabled>
		<input type="button" value="Cancel" id=cmdCancel name=cmdCancel>
      </td>
    </tr>
  </table>
</div>

</form>
</BODY>
</HTML>
