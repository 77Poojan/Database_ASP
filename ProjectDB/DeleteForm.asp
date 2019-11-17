<html>
<head>
	<title> Delete Records </title>
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
		dim conn,rs
		set conn = Server.CreateObject("ADODB.Connection")
		conn.Mode = adModeReadWrite
		conn.ConnectionString = "Provider=MSOLEDBSQL; Server=.\SQLEXPRESS; Database= ; User Id= ; Password= ;"
		conn.open
		set rs = Server.CreateObject("ADODB.recordset")
		rs.Open "SELECT * FROM Users", conn
	%>

	<h2>Delete Records</h2>
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
	<form method="post" action="delete.asp">
	<%
	for each x in rs.Fields
	  if x.name="ID" then%>
	    <td>
	    <input type="submit" name="ie" value="<%=x.value%>">
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