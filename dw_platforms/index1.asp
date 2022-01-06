<%@LANGUAGE="VBSCRIPT" CODEPAGE="1251"%>
<!--#include file="Connections/conn_platforms.asp" -->

<%
Dim rs_0101
set rs_0101 = server.CreateObject("ADODB.RECORDSET")
	rs_0101.ActiveConnection = conn_platforms_STRING
	rs_0101.Source = "SELECT t_0101.p01, t_0101.p02, t_0101.p03 FROM t_0101"
	rs_0101.CursorType = 0
	rs_0101.CursorLocation = 2
	rs_0101.LockType = 1
	rs_0101.Open()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Робоча область</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251" />
</head>
<body>
		<table  border="1px solid black" cellpadding="0" cellspacing="0">
        <tr style=" width:auto"><td><%=rs_0101("p01")%></td><td><%=rs_0101("p02")%></td><td><%=rs_0101("p03")%></td></tr>
        </table>
					<%		
				rs_0101.Close()		
			Set rs_0101 = Nothing
		 %>   
	</body>
</html>