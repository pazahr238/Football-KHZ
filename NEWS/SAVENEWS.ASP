<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<%
dim cn
dim rs
dim strsql
dim str
set cn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
str="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("happen.mdb") & ";Persist Security Info=False" 
cn.open str
strsql="select * from news"
rs.open strsql,cn , 3, 3
if rs.recordcount > 19 then
rs.movelast
rs.delete
rs.update
end if
set rs=cn.execute("INSERT INTO news VALUES('" & request("title") & "','" & date & "','" & request("userdate") & "','" & request("new") & "')")
cn.close
set cn=nothing
%>
	<script language="javascript">
		window.location='endnews.html';
	</script> 
</body>
</html>