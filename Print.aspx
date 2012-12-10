<%@ Page Language="vb" Codebehind="Qbr.aspx.vb" Inherits="QBRProject.QBR" %>
<%@ import namespace = "System" %>
<%@ import namespace = "System.Data" %>
<%@ import namespace = "System.Data.OleDb " %>
<%@ import namespace = "System.Web.UI" %>
<%@ import namespace = "System.Web.UI.WebControls" %>
<HTML>
	<HEAD>
		<title>QBR</title>
		<script runat="server">
    '|------------------------------------------------------------------------------------------------------------------|
    '| Main : Called on page load																						|
    '|------------------------------------------------------------------------------------------------------------------|
	
    sub Page_Load
		Dim dbConn as oleDbConnection
        EstablishConnection(dbConn)
        dbConn.open
        ShowPrintableQbr(dbConn)
      
		if not page.ispostback then
			if Session("Print") = "Email" then
				txtText.text = Request.form("txtText")
				txtFrom.text =  Request.form("txtFrom")
				txtTo.text =  Request.form("txtTo")
				txtCC.text = Request.form("txtCC") 
				txtBCC.text =  Request.form("txtBCC")
				txtSubject.text = Request.form("txtSubject")
				txtSignature.text = Request.form("txtSignature")
				NoQbr.text = Request.QueryString("Nu")
			else
				lblTextMail.text = ""
				lblSignature.text = ""
				txtText.text = ""
				txtFrom.text =  ""
				txtTo.text =  ""
				txtCC.text = ""
				txtBCC.text =  ""
				txtSignature.text = ""
				NoQbr.text = ""
				txtSubject.text = ""
			end if
		end if
		
		lblTextMail.text = Replace(txtText.text, vbNewLine, "<br />")
		lblSignature.text = Replace(txtSignature.text, vbNewLine, "<br />")
        
        dbConn.close
    end sub
		</script>
		<link rel="stylesheet" type="text/css" href="Styles.css">
			<style>
			.MiniText { FONT-SIZE: 80% }
			.SmallText { FONT-SIZE: 90% }
			.MidText { FONT-SIZE: 150% }
			.NormalText { FONT-SIZE: 110% }
			.BigText { FONT-SIZE: 250% }
			.Bold { FONT-WEIGHT: bold }
			.Erreur { COLOR: red }
			.Bordure { BORDER-RIGHT: #330099 thin solid; PADDING-RIGHT: 4px; BORDER-TOP: #330099 thin solid; PADDING-LEFT: 4px; BACKGROUND: #99ccff; PADDING-BOTTOM: 4px; BORDER-LEFT: #330099 thin solid; PADDING-TOP: 4px; BORDER-BOTTOM: #330099 thin solid }
			.Bordure1 { BORDER-TOP-WIDTH: thin; BORDER-LEFT-WIDTH: thin; BORDER-LEFT-COLOR: #330099; BACKGROUND: aliceblue; BORDER-BOTTOM-WIDTH: thin; BORDER-BOTTOM-COLOR: #330099; BORDER-TOP-COLOR: #330099; BORDER-RIGHT-WIDTH: thin; BORDER-RIGHT-COLOR: #330099 }
			.BordureTablePrint { BORDER-RIGHT: 1pt solid; BORDER-TOP: 1pt solid; BORDER-LEFT: 1pt solid; BORDER-BOTTOM: 1pt solid; BORDER-COLLAPSE: collapse }
			.BordureTDPrint { BORDER-RIGHT: 1pt solid; BORDER-TOP: 1pt solid; BORDER-LEFT: 1pt solid; BORDER-BOTTOM: 1pt solid; BORDER-COLLAPSE: collapse }
			.BordureTablePrint { WIDTH: 100% }
			.CellLength1 { WIDTH: 120px }
			.padding { PADDING-RIGHT: 20px; PADDING-LEFT: 3px }
			.spaces { PADDING-RIGHT: 6px }
			</style>
			<script type="text/javascript" src="QBR.js"></script>
	</HEAD>
	<body>
		<br>
		<form runat="server">
			<% if Session("Print") = "Email" then %>
			<asp:Label id="lblTextMail" class="NormalText" runat="server" />
			<% End If%>
			<center>
				<table>
					<tr>
						<td>
							<table width="100%">
								<tr valign="top">
									<td><img src="LAURENTIDE_LOGO.bmp" width="150" height="86"></td>
									<td align="center" class="BigText">QBR FORM</td>
									<td><img src="emerson-small.bmp" width="150" height="75"></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr height="20">
						<td></td>
					</tr>
					<tr>
						<td>
							<table class="BordureTablePrint">
								<tr>
									<td class="bold BordureTDPrint CellLength1">Date:</td>
									<td colspan="3"><asp:Label id="lblDateQBR" size="30" runat="server" />&nbsp;</td>
								</tr>
								<tr>
									<td class="bold BordureTDPrint CellLength1">Employee:</B></td>
									<td class="BordureTDPrint"><asp:Label id="lblEmploye" runat="server" />&nbsp;</td>
									<td class="bold BordureTDPrint CellLength1">Outside Sale:</td>
									<td class="BordureTDPrint"><asp:Label id="lblOS" runat="server" />&nbsp;</td>
								</tr>
								<tr>
									<td class="bold BordureTDPrint CellLength1">Enterprise:</B></td>
									<td class="BordureTDPrint" colspan="3"><asp:Label id="lblCorporation" runat="server" />&nbsp;</td>
								</tr>
								<tr>
									<td class="bold BordureTDPrint CellLength1">Customer Name:</td>
									<td class="BordureTDPrint"><asp:Label id="lblCustomerName" runat="server" />&nbsp;</td>
									<td class="bold BordureTDPrint CellLength1">Contact Name:</td>
									<td class="BordureTDPrint"><asp:Label id="lblContactName" runat="server" />&nbsp;</td>
								</tr>
								<tr>
									<td class="bold BordureTDPrint CellLength1">Address:</td>
									<td class="BordureTDPrint"><asp:Label id="lblAddress" runat="server" />&nbsp;</td>
									<td class="bold BordureTDPrint CellLength1">Contact title:</td>
									<td class="BordureTDPrint"><asp:Label id="lblContactTitle" runat="server" />&nbsp;</td>
								</tr>
								<tr>
									<td class="bold BordureTDPrint CellLength1">Location:</td>
									<td class="BordureTDPrint"><asp:Label id="lblLocation" runat="server" />&nbsp;</td>
									<td class="bold BordureTDPrint CellLength1">Contact eMail:</td>
									<td class="BordureTDPrint"><asp:Label id="lblContactEMail" runat="server" />&nbsp;</td>
								</tr>
								<tr>
									<td class="bold BordureTDPrint CellLength1">Industry:</td>
									<td class="BordureTDPrint"><asp:Label id="lblIndustry" runat="server" />&nbsp;</td>
									<td class="bold BordureTDPrint CellLength1">Contact tel #:</td>
									<td class="BordureTDPrint"><asp:Label id="lblContactTel" runat="server" />&nbsp;</td>
								</tr>
								<tr>
									<td class="bold BordureTDPrint" colspan="2">Where can we use the QBR?</td>
									<td class="BordureTDPrint" colspan="2"><asp:Label id="lblUse" runat="server" /></td>
								</tr>
								<tr>
									<td class="bold BordureTDPrint" colspan="2">Input Type</td>
									<td class="BordureTDPrint" colspan="2"><asp:Label id="lblInputType" runat="server" /></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr height="20">
						<td></td>
					</tr>
					<tr>
						<td>
							<table class="BordureTablePrint">
								<tr>
									<td class="bold BordureTDPrint CellLength1">Project Name:</td>
									<td class="BordureTDPrint"><asp:Label id="lblProjectName" runat="server" />&nbsp;</td>
								</tr>
								<tr>
									<td class="bold BordureTDPrint CellLength1">Application:</td>
									<td class="BordureTDPrint"><asp:Label id="lblApplication" runat="server" />&nbsp;</td>
								</tr>
								<tr valign="top">
									<td class="bold BordureTDPrint CellLength1">Products:</td>
									<td class="BordureTDPrint"><asp:Label id="lblProduct" runat="server" /></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr height="20">
						<td></td>
					</tr>
					<tr>
						<td>
							<table class="BordureTablePrint">
								<tr>
									<td class="bold BordureTDPrint" colspan="4">Original Situation / Challenge (include 
										sketch, photos):</td>
								</tr>
								<tr>
									<td class="BordureTDPrint" colspan="4">
										<asp:Label id="lblSituation" runat="server" />&nbsp;<br>
										<br>
										<asp:Label id="lblLinksSituation" runat="server" />
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr height="20">
						<td></td>
					</tr>
					<tr>
						<td>
							<table class="BordureTablePrint">
								<tr>
									<td class="bold BordureTDPrint" colspan="4">Solution (include sketch, photos):</td>
								</tr>
								<tr>
									<td class="BordureTDPrint" colspan="4"><asp:Label id="lblSolution" runat="server" />&nbsp;<br>
										<br>
										<asp:Label id="lblLinksSolution" runat="server" />
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr height="20">
						<td></td>
					</tr>
					<tr>
						<td>
							<table class="BordureTablePrint">
								<tr>
									<td class="bold BordureTDPrint" colspan="4">Improved Business Results Calculation:</td>
								</tr>
								<tr>
									<td class="BordureTDPrint" colspan="4"><asp:Label id="lblResults" runat="server" />&nbsp;<br>
										<br>
										<asp:Label id="lblLinksResult" runat="server" />
									</td>
								</tr>
								<tr>
									<td colspan="4" class="BordureTDPrint">
										<table>
											<tr>
												<td class="bold">Costs:
												</td>
												<td><asp:Label id="lblCost" runat="server" />&nbsp;&nbsp;&nbsp;</td>
												<td class="bold">One time savings:
												</td>
												<td><asp:Label id="lblOnceSavings" runat="server" />&nbsp;&nbsp;&nbsp;</td>
												<td class="bold">ROI:</td>
												<td><asp:Label id="lblROI" runat="server" />&nbsp;&nbsp;&nbsp;</td>
												<td class="bold">Approved by Customer:</td>
												<td><asp:Label id="lblApproved" runat="server" /></td>
											</tr>
											<tr>
												<td></td>
												<td></td>
												<td class="bold">Annual savings:
												</td>
												<td><asp:Label id="lblAnnualSavings" runat="server" />&nbsp;&nbsp;&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr height="20">
						<td></td>
					</tr>
					<tr>
						<td>
							<table class="BordureTablePrint">
								<tr>
									<td class="bold BordureTDPrint" colspan="4">Customer Testimonial / Quote:</td>
								</tr>
								<tr>
									<td class="BordureTDPrint" colspan="4"><asp:Label id="lblTestimonial" runat="server" />&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr height="10">
						<td></td>
					</tr>
					<tr>
						<td>
							<table width="100%">
								<tr>
									<td colspan="4" align="center">
										<% if Session("Print") <> "Email" then %>
										<input type="button" class="Button" id="btnPrint" value="Print" onClick="javascript:PrintQBR();">
										<% End If %>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</center>
			<% if Session("Print") = "Email" then %>
			<asp:Label id="lblSignature" class="NormalText" Runat="server" /><br>
			<br>
			<asp:TextBox ID="txtFrom" Visible="False" Runat="server" />
			<asp:TextBox ID="txtTo" Visible="False" Runat="server" />
			<asp:TextBox ID="txtCC" Visible="False" Runat="server" />
			<asp:TextBox ID="txtBCC" Visible="False" Runat="server" />
			<asp:TextBox ID="txtText" Visible="False" Runat="server" />
			<asp:TextBox ID="txtSignature" Visible="False" Runat="server" />
			<asp:TextBox ID="txtSubject" Visible="False" Runat="server" />
			<asp:TextBox ID="NoQbr" Visible="False" Runat="server" />
			<center>
				<asp:button class="Button" id="SendEmail" text="Send" OnClick="SendMail" Runat="server" />
			</center>
			<% End If %>
		</form>
	</body>
</HTML>
