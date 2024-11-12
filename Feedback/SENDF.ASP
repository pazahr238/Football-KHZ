<html>
<body>
<%
dim cn
dim rs
dim strsql
dim str
set cn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
str="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("f_am.mdb") & ";Persist Security Info=False" 
cn.open str
strsql="INSERT INTO feedback VALUES('" & request("comment") & "','" & request("name") & "','" & request("last") & "','" & request("degree") & "','" & request("tel") & "','" & request("address") & "')"
 rs.open strsql,cn , 3, 3
cn.close
set cn=nothing
%>
<script language="javascript">
		window.location='thanks.html';
	</script> 
</body>
</html>