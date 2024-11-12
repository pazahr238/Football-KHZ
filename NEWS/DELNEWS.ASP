<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>&#1581;&#1584;&#1601; &#1582;&#1576;&#1585; &#1583;&#1604;&#1582;&#1608;&#1575;&#1607;</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#669966">
<form name="form1" method="post" action="delaction.asp">
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
rs.movefirst
do while not rs.eof
 %>
  <div align="right"></div>
  <div align="right"> <b> 
    <% response.write rs("title") %>
    <input type="checkbox" name="checkbox" value=<%=rs("title")%>>
    </b> <br>
    <%
response.write rs("userdate")
%>
    <br>
    <%
response.write rs("context")
%>
    <br>
  </div>
  <hr>
  <div align="right"><br>
    <%
rs.movenext
loop

cn.close
set cn=nothing
%>
    <br>
    <input type="submit" name="Submit" value="&#1578;&#1575;&#1610;&#1610;&#1583;" style="font-family: Tahoma">
  </div>
</form>
</body>
</html>