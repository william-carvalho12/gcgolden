<%
'//////////////////////////////////////////////////CONEXAO DO BANCO DE DADOS 1////////////////////////////////////////////////////////////////////
		' ConnectString = "driver={SQLServer}; server=198.27.115.205\SQLEXPRESS; database=gc; uid=sa; pwd=@Senhasqlgc;"
                ConnectString = "Provider=SQLNCLI11;Server=198.27.115.205;Database=gc;Uid=sa;Pwd=@Senhasqlgc;"
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open ConnectString
		Set rs = Server.CreateObject("ADODB.Recordset") 
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

   

sql = "SELECT * FROM dbo.users INNER JOIN dbo.CashUsers ON dbo.users.Login = dbo.CashUsers.Login"
rs.Open sql, Conn
      
%>
<head>
    <link href="css/metro.css" rel="stylesheet">
    <link href="css/metro-icons.css" rel="stylesheet">
    <link href="css/docs.css" rel="stylesheet">

    <script src="js/jquery-2.1.3.min.js"></script>
    <script src="js/jquery.dataTables.min.js"></script>
    <script src="js/metro.js"></script>
    <script src="js/docs.js"></script>
    <script src="js/prettify/run_prettify.js"></script>
    <script src="js/ga.js"></script>

    <script>
        $(function(){
            //$('#example_table').dataTable();
        });
    </script>
</head>
<table id="example_table" class="dataTable striped border bordered" data-role="datatable" data-searching="true">
	<Caption>Relação de Usuarios</caption>     
	<thead>
		<tr>
			<th>ON</th>
			<th>UID</th>
			<th>Login</th>
			<th>Sexo</th>
			<th>GP</th>
			<th>Cash</th>
			<th>PlayTime</th>
			<th>E-mail</th>
			<th>Ult.Conexão</th>
		</tr>
	</thead>

	<tfoot>
		<tr>
			<th>ON</th>
			<th>UID</th>
			<th>Login</th>
			<th>Sexo</th>
			<th>GP</th>
			<th>Cash</th>
			<th>PlayTime</th>
			<th>E-mail</th>
			<th>Ult.Conexão</th>
		</tr>
	</tfoot>
	
	<tbody>
<%
       
         do while not rs.eof
            Response.Write "<tr>"
			 Response.Write "<td> " & rs.fields("Connecting") & "</td>"& vbcrlf
            Response.Write "<td> " & rs.fields("LoginUID") & "</td>"& vbcrlf
            Response.Write "<td> " & rs.fields("Login") & "</td>"& vbcrlf
            Response.Write "<td> " & rs.fields("sex") & "</td>"& vbcrlf
            Response.Write "<td> "  & rs.fields("gamePoint") & "</td>"& vbcrlf
			Response.Write "<td> "  & rs.fields("Cash") & "</td>"& vbcrlf
            Response.Write "<td> "  & rs.fields("playTime") & "</td>"& vbcrlf
            Response.Write "<td> "  & rs.fields("email")& "</td>"& vbcrlf
			Response.Write "<td> "  & FormatDateTime(rs.fields("lastConnect"),2)& "</td>"& vbcrlf
            Response.Write "</tr>"& vbcrlf
            rs.movenext
         loop
         Response.Write "</tbody></table>"& vbcrlf
     rs.Close
%>