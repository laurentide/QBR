<%@ Page Language="vb" Codebehind="Qbr.aspx.vb" Inherits="QBRProject.QBR" %>
<%@ import namespace = "System.Data.OleDb " %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
  <HEAD>
		<title>Email</title>
		<script runat="server">
    '|------------------------------------------------------------------------------------------------------------------|
    '| Main : Called on page load																						|
    '|------------------------------------------------------------------------------------------------------------------|
	
    sub Page_Load
		Dim dbConn as oleDbConnection
        EstablishConnection(dbConn)
        dbConn.open
        
        if not page.ispostback then
			Dim intNoEmp as string   = ReturnField(dbConn, "Qbr", "EmpNo", "QbrNo", Request.queryString("Nu"))
			Dim intNocontact as string= ReturnField(dbConn, "Qbr", "ContactNo", "QbrNo", Request.queryString("Nu"))
						
			Dim strTemp as string = ReturnField(dbConn, "Employee", "Name", "EmpNo", intNoEmp)
	        
			txtFrom.text = Replace(strTemp, " ", ".")
			txtFrom.text += "@laurentidecontrols.com"
			txtTo.text = ReturnField(dbConn, "Contact", "EMail", "ContactNo", intNocontact)
			
			NoQbr.text = Request.queryString("Nu")
			
			dbConn.close
		end if
    end sub
		</script>
		<script type="text/javascript" src="QBR.js"></script>
		<link rel="stylesheet" type="text/css" href="Styles.css">
  </HEAD>
	<body leftMargin="0" bottommargin="0" topmargin="0" RightMargin="0" marginheight="0" marginwidth="0">
		<form method="post" runat="server" id="QBR">
			<center>
				<table width="100%" border="3">
					<tr>
						<td class="bg1" width="6%"></td>
						<td align="center">
							<table>
								<tr height="20">
									<td></td>
								</tr>
							</table>
							<asp:Label class="MidText DarkBlue underLine" text="Sending Email" runat="server" id="Label1" />
							<table>
								<tr height="30">
									<td></td>
								</tr>
							</table>
							<table class="Bordure">
								<tr height="40">
									<td class="bordure DarkBlue">
										From:
									</td>
									<td class="bordure1">
										<asp:TextBox id="txtFrom" width="400" runat="server"></asp:TextBox>
									</td>
								</tr>
								<tr height="40">
									<td class="bordure DarkBlue">
										To:
									</td>
									<td class="bordure1">
										<asp:TextBox id="txtTo" width="400" runat="server"></asp:TextBox>
									</td>
								</tr>
								<tr height="40">
									<td class="bordure DarkBlue">
										CC:
									</td>
									<td class="bordure1">
										<asp:TextBox id="TxtCC" width="400" runat="server"></asp:TextBox>
									</td>
								</tr>
								<tr height="40">
									<td class="bordure DarkBlue">
										BCC:
									</td>
									<td class="bordure1">
										<asp:TextBox id="TxtBCC" width="400" runat="server"></asp:TextBox>
									</td>
								</tr>
								<tr height="40">
									<td class="bordure DarkBlue">
										Subject:
									</td>
									<td class="bordure1">
										<asp:TextBox id="txtSubject" width="400" runat="server">Quantified Business Result</asp:TextBox>
									</td>
								</tr>
								<tr height="40">
									<td colspan="2" class="bordure DarkBlue">
										Text to be inserted before Qbr
									</td>
								</tr>
								<tr>
									<td colspan="2" class="bordure1">
										<asp:TextBox id="txtText" width="470" runat="server" TextMode="MultiLine" Rows="5"></asp:TextBox>
									</td>
								</tr>
								<tr height="40">
									<td colspan="2" class="bordure DarkBlue">
										Signature (will be inserted at the end of the message)
									</td>
								</tr>
								<tr>
									<td colspan="2" class="bordure1">
										<asp:TextBox id="txtSignature" width="470" runat="server" TextMode="MultiLine" Rows="5"></asp:TextBox>
									</td>
								</tr>
								<tr>
									<td colspan="2" align="center">
										<asp:button class="button" Text="Preview" OnClick="PreviewMail" runat="server" id="Button1" />
									</td>
								</tr>
							</table>
						</td>
						<td class="bg1" width="6%"></td>
					</tr>
				</table>
			</center>
			<asp:textbox ID="NoQbr" visible="False" runat="server" />
		</form>
	</body>
</HTML>
