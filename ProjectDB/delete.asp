<html>
<head>
    <title> Delete Conformation Page</title>
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

<h2>Delete Record Conformation</h2>
<%
dim conn,rs,id
  
    set conn = Server.CreateObject("ADODB.Connection")
    conn.Mode = adModeReadWrite
    conn.ConnectionString = "Provider=MSOLEDBSQL; Server=.\SQLEXPRESS; Database= ; User Id= ; Password= ;"
    conn.open

    id=Request.Form("ie")
    
    sql="DELETE FROM Users"
    sql=sql & " WHERE ID='" & id & "'"
    on error resume next
    conn.Execute sql
    if err<>0 then
      response.write("No update permissions!")
    else
      response.write("Record " & id & " was deleted!")
    end if

conn.close
%>

<form>
    <a href="http://localhost:3000/Connection_Setup.asp"><input type="button" value="Home"></a>
</form>
</body>
</html>