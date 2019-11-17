<html>
<head>
	<title>Insert Conformation Page</title>
	<style type="text/css">
	    input{
	      border:2px;
	      border-radius: 5px;
	      margin-left: 280px;
	      margin-top: 5px;
	    }
  	</style>
</head>
<body>

<%
	dim conn,sql
	
	set conn = Server.CreateObject("ADODB.Connection")
	conn.Mode = adModeReadWrite
	conn.ConnectionString = "Provider=MSOLEDBSQL; Server=.\SQLEXPRESS; Database=MBMDB2; User Id=PujanDB; Password=PJN;"
	conn.open

	set sql = Server.CreateObject("ADODB.recordset")
	sql.Open "SELECT * FROM Users", conn
		
	sql="INSERT INTO Users VALUES "
	sql=sql & "('" & Request.Form("ie") & "',"
	sql=sql & "'" & Request.Form("fname") & "',"
	sql=sql & "'" & Request.Form("lname") & "',"
	sql=sql & "'" & Request.Form("address") & "',",
	sql=sql & "'" & Request.Form("phone") & "',"
	sql=sql & "'" & Request.Form("dob") & "',"
	sql=sql & "'" & Request.Form("fee") & "')"
	on error resume next
	conn.Execute sql,recaffected
%>
<h2> Record Added Conformation </h2>
<%	
	if err<>0 then
	  Response.Write("No update permissions!")
	else
	  Response.Write("<h3>" & recaffected & " record added</h3>")
	end if
	conn.close
%>

<form>
    <a href="http://localhost:3000/NCCClass/AssignmentSQL.asp"><input type="button" value="Home"></a>
</form>
</body>
</html>