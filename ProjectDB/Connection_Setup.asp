<%@ Language = VBScript %>
<!--#include file="adovbs.inc"-->

<!DOCTYPE html>
<html>
<head>
	<title> Database Records </title>
	<style>
		th,tr,td{
			border: 1px dotted #7d7357;
			width: 150px;
		}	

		table{
			margin-left: 100px;
		}

		hr{
			margin-top: -10px;
		}

		input{
			border:2px;
			border-radius: 5px;
			margin-left: 280px;
			margin-top: 5px;
		}
	</style>
</head>
<body>
	
	<!--Database Connection Setup-->
	<%
		dim conn,rs
	
		set conn = Server.CreateObject("ADODB.Connection")
		conn.Mode = adModeReadWrite
		conn.ConnectionString = "Provider=MSOLEDBSQL; Server=.\SQLEXPRESS; Database= ; User Id= ; Password= ;"
		conn.open
		response.write("=> Database is connected sucessfully...")	

		set rs = Server.CreateObject("ADODB.recordset")
		rs.Open "SELECT * FROM Users", conn
		
		response.Write("<br> <br>") 
	%>

	<!--Display Records Title-->
	<h3 align="center"> Records </h3> 
	<hr size=1 width=55 color=black align="center">

	<table>
			<tr>
		<% for each x in rs.Fields 
		    	response.write("<th>" & ucase(x.name) & "</th>")
			next %>
		 	</tr>
	</table>

	<!--Display Records-->
	<%do until rs.EOF %>
		<table>
			<tr>
		<% for each x in rs.Fields %>
		    	 <td> <%=x.value%> </td>
		 <% next %>
		 	</tr>
		</table>
		<%rs.MoveNext%>
	<%loop%>

	<br> <br>
	<form>
		<a href="http://localhost:3000/ProjectDB/InsertForm.html"><input type="button" value="Insert"></a>
		<a href="http://localhost:3000/ProjectDB/DeleteForm.asp"><input type="button" value="Delete"></a>
		<a href="http://localhost:3000/ProjectDB/UpdateForm.asp"><input type="button" value="Update"></a>
	</form>
</body>
</html>