<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<%
dim iloop
dim strdelitems
dim adelitems
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
strdelitems = request("checkbox")
adelitems = split(strdelitems,",")
		for iloop = lbound(adelitems) to ubound(adelitems)
		rs.movefirst
				do while not rs.eof
					if trim(adelitems(iloop)) = rs("title") then
					rs.delete
					rs.update
					end if
				rs.movenext
				loop
		next
cn.close
set cn=nothing
%>
	<script language="javascript">
		window.location='news.asp';
	</script> 
</body>
</html>