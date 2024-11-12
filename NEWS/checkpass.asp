<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Checking ...</title>
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
 strsql="select * from check"
rs.open strsql,cn , 3, 3
if request("adminpass") = "a2m429cf6" then 
rs.movefirst
rs("user") = "admin"
rs.update
%>
	<script language="javascript">
		window.location='newsform.asp';
	</script> 
<%
else 
%>
	<script language="javascript">
		window.location='news.asp';
	</script> 
<%
end if
cn.close
set cn=nothing
%>
</body>
</html>