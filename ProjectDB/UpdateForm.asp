<html>
<head>
	<title> Update Records </title>
<style>
		th,tr,td{
			border: 1px dotted #7d7357;
			width: 150px;
		}

		table{
			margin-left: 100px;
		}
</style>
</head>
<body>
<%
	set conn = Server.CreateObject("ADODB.Connection")
		conn.Mode = adModeReadWrite
		conn.ConnectionString = "Provider=MSOLEDBSQL; Server=.\SQLEXPRESS; Database= ; User Id= ; Password= ;"
		conn.open
set rs=Server.CreateObject("ADODB.Recordset")
rs.open "SELECT * FROM Users",conn
%>

<h2>Update Database Record</h2>
<table>
<tr>
<%
for each x in rs.Fields
  response.write("<th>" & ucase(x.name) & "</th>")
next
%>
</tr>
<% do until rs.EOF %>
<tr>
<form method="post" action="update.asp">
<%
for each x in rs.Fields
  if lcase(x.name)="id" then%>
    <td>
    <input type="submit" name="ID" value="<%=x.value%>">
    </td>
  <%else%>
    <td><%Response.Write(x.value)%></td>
  <%end if
next
%>
</form>
<%rs.MoveNext%>
</tr>
<%
loop
conn.close
%>
</table>

</body>
</html>