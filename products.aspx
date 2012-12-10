<%@ import namespace = "System.Web.UI.WebControls" %>
<%@ import namespace = "System.Web.UI" %>
<%@ import namespace = "System.Data.OleDb " %>
<%@ import namespace = "System.Data" %>
<%@ import namespace = "System" %>
<%@ Page Language="vb" Codebehind="Qbr.aspx.vb" Inherits="QBRProject.QBR"%>
<HTML>
	<HEAD>
		<title>Products</title>
		<script runat="server">
    '|------------------------------------------------------------------------------------------------------------------|
    '| Main : Call on page load																							|
    '|------------------------------------------------------------------------------------------------------------------|
	
    sub Page_Load
		Dim dbConn as OleDbConnection
		EstablishConnection(dbConn)
		dbConn.open
         
		if request.queryString("Nu") <> Nothing Then			
			if isNumeric(request.queryString("Nu")) Then
				if not page.isPostBack then
					'defines how many lines to print
					Session("NbProducts") = 3
					AddObjects(dbConn)
				else if request.form("More") = "More" then
					Session("NbProducts") += 1
					AddObjects(dbConn)
				else
					AddObjects(dbConn)
				end if
    		end if
    	end if
    	
    	dbConn.close
	end sub
	
		</script>
		<link rel="stylesheet" type="text/css" href="Styles.css">
			<script>
			//places current window in the center
			var xpos = (screen.width - 900) / 2
			var ypos = (screen.height - 450) / 2
			moveTo(xpos, ypos);
			</script>
	</HEAD>
	<body>
		<form runat="server" method="post">
			<table>
				<tr height="50" valign="top">
					<td class="DarkBlue MidText underline">
						Adding Products to the QBR
					</td>
				</tr>
				<tr height="40" valign="top">
					<td class="DarkBlue SmallText bold">
						Please select the main product, and then check all the specific products needed 
						for the QBR.
					</td>
				</tr>
				<tr>
					<td>
						<asp:PlaceHolder id="ph1" runat="server" />
					</td>
				</tr>
				<tr>
					<td>
						<% if request.queryString("Mode") = "1" or request.queryString("Mode") = "2" then %>
						<asp:Button class="button" text="More" id="More" runat="server" />&nbsp;&nbsp;&nbsp;&nbsp;
						<% end if %>
						<asp:Button class="button" text="OK" onClick="ConfirmProducts" runat="server" id="Button1" />
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
