<%@ Page Language="vb" Codebehind="Qbr.aspx.vb" Inherits="QBRProject.QBR"%>
<%@ import namespace = "System" %>
<%@ import namespace = "System.Data" %>
<%@ import namespace = "System.Data.OleDb " %>
<%@ import namespace = "System.Web.UI" %>
<%@ import namespace = "System.Web.UI.WebControls" %>

<%@ import namespace = "System.drawing" %>

<script runat="server">
    '|------------------------------------------------------------------------------------------------------------------|
    '| Main : Call on page load																							|
    '|------------------------------------------------------------------------------------------------------------------|
	
    sub Page_Load
		Dim dbConn as OleDbConnection
		
		EstablishConnection(dbConn)
		dbConn.open
		
    	CreateReport(dbConn)
    	
    	dbConn.close
    
	end sub
	
</script>

<html>
	<head>
		<title>Report</title>
		<link rel="stylesheet" type="text/css" href="Styles.css" />
	</head>
	<body>
		<form runat="server">
			<table width="500%">
				<tr>
					<td>
						<asp:PlaceHolder id="ph1" runat="server" />
					</td>
				</tr>
				<tr height="20px"><td></td></tr>
				<tr>
					<td>
						<asp:dataGrid id="dgQbr" runat="server" 
							BorderColor="black"
							BorderWidth="1"
							GridLines="Both"
							CellPadding="3"
							CellSpacing="0"
							Font-Name="Verdana"
							Font-Size="8pt"
							HeaderStyle-BackColor="#99CCFF"
							HeaderStyle-Font-Bold="True"
							AutoGenerateColumns="False" 
							/>
						<td>
					</tr>
					<tr height="50px" valign="bottom">
						<td>
							<asp:button class="button" onclick="ExportToExcel" text="Export to Excel" runat="server" />
						</td>
					</tr>
				</table>
							
		</form>
	</body>

</html>