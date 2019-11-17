<html>
<head>
  <title> Update Database Records</title>
  <style type="text/css">
    .sd{
      border:2px;
      border-radius: 5px;
      margin-left: 280px;
      margin-top: 5px;
    }
  </style>
</head>
<body>
<%
dim conn,rs,cid
set conn = Server.CreateObject("ADODB.Connection")
    conn.Mode = adModeUpdate
    conn.ConnectionString = "Provider=MSOLEDBSQL; Server=.\SQLEXPRESS; Database= ; User Id= ; Password= ;"
    conn.open

cid=Request.Form("ID")

if Request.form("FirstName")="" then
  set rs=Server.CreateObject("ADODB.Recordset")
  rs.open "SELECT * FROM Users where ID='"&cid&"'",conn
  %>
<h2>Update Record</h2>  
  <form method="post" action="update.asp">
  <table>
  <%for each x in rs.Fields%>
  <tr>
  <td><%=x.name%></td>
  <td><input name="<%=x.name%>" value="<%=x.value%>"></td>
  <%next%>
  </tr>
  </table>
  <br><br>
  <input type="submit" value="Update record">
  </form>
<%else%>
<%
  sql="UPDATE Users SET "
  sql=sql & "ID='" & Request.Form("ID") & "',"
  sql=sql & "FirstName='" & Request.Form("FirstName") & "',"
  sql=sql & "LastName='" & Request.Form("LastName") & "',"
  sql=sql & "Address='" & Request.Form("Address") & "',"
  sql=sql & "Phone='" & Request.Form("Phone") & "',"
  sql=sql & "DOB='" & Request.Form("DOB") & "',"
  sql=sql & "TotalFee='" & Request.Form("TotalFee") & "'"
  sql=sql & " WHERE ID='" & cid & "'"
%>
<%  
  on error resume next
  conn.Execute sql
%>

<h2>Update Record Conformation </h2> 
  <%if err<>0 then
    response.write("No update permissions!") 
  else
    response.write("Record " & cid & " was updated!") %>

    <form>
      <a href="http://localhost:3000/Connection_Setup.asp"><input type="button" value="Home" class="sd"></a>
    </form>

<%    
  end if
end if
conn.close
%>


</body>
</html>