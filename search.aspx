<%@ Page Language="vb" Codebehind="Qbr.aspx.vb" Inherits="QBRProject.QBR"%>
<%@ import namespace = "System" %>
<%@ import namespace = "System.Data" %>
<%@ import namespace = "System.Data.OleDb " %>
<%@ import namespace = "System.Web.UI" %>
<%@ import namespace = "System.Web.UI.WebControls" %>

<HTML>
  <HEAD>
		<title>Search</title>
		<script runat="server">
    '|----------------------------------------------------------------------------------------------------------------|
    '| Main : Call on page load																						  |
    '|----------------------------------------------------------------------------------------------------------------|
	
    sub Page_Load
    	dim dbConn as oleDbConnection
		EstablishConnection(dbConn)
		dbConn.open()

		Dim strWhere as string = ""

		'Active une variable session pour définir si on est en ajout, modification ou consultation
		if request.queryString("Type") <> Nothing then
			Session("Type") = request.queryString("Type")
		end if

		'if first time entered or in "order mode"
		if request.form("btnSearch") = "" and request.form("ImgStart.x") = "" _
				and request.form("ImgEnd.x") = "" and Request.Form("__EVENTTARGET") = "" then	
			txtCorporation.text = Session("Corporation")
			txtCustomer.text = Session("Customer")
			txtIndustry.text = Session("Industry")
			txtApplication.text = Session("Application")
			txtEmploye.text = Session("Employee") 
			txtStartDate.text = Session("StartDate")
			txtEndDate.text = Session("EndDate")
			txtProduct.text = Session("Product")
			txtProduct.text = Session("SubProduct")
			txtProduct.text = Session("ModelNo")
			
			ShowQbr(dbConn, request.queryString("Where"))
		end if
		
		dbConn.close
    end sub
    
		</script>
		<link rel="stylesheet" type="text/css" href="Styles.css">
			<script LANGUAGE="JavaScript" SRC="CalendarPopup.js"></script>
			<link rel="stylesheet" type="text/css" href="Calendar.css">
  </HEAD>
	<body leftMargin="0" bottommargin="0" topmargin="0" RightMargin="0" marginheight="0" marginwidth="0">
		<form runat="server" method="post">
			<center>
				<table width="100%" height="100%" border="3">
					<tr valign="top">
						<td class="bg1" width="6%"></td>
						<td align="center">
							<table>
								<tr height="30">
									<td></td>
								</tr>
							</table>
							<table>
								<tr>
									<td class='DarkBlue Bold MiniText'>Enterprise</td>
									<td class='DarkBlue Bold MiniText'>Customer</td>
									<td class='DarkBlue Bold MiniText'>Industry</td>
									<td class='DarkBlue Bold MiniText'>Application</td>
									<td class='DarkBlue Bold MiniText'>Employee</td>
								</tr>
								<tr>
									<td class='spaces'><asp:TextBox class="BGBabyBlue DarkBlue" id="txtCorporation" size="18" runat="server" /></td>
									<td class='spaces'><asp:TextBox class="BGBabyBlue DarkBlue" id="txtCustomer" size="18" runat="server" /></td>
									<td class='spaces'><asp:TextBox class="BGBabyBlue DarkBlue" id="txtIndustry" size="18" runat="server" /></td>
									<td class='spaces'><asp:TextBox class="BGBabyBlue DarkBlue" id="txtApplication" size="18" runat="server" /></td>
									<td class='spaces'><asp:TextBox class="BGBabyBlue DarkBlue" id="txtEmploye" size="18" runat="server" /></td>
								</tr>
								<tr>
									<td class='DarkBlue Bold MiniText'>Product</td>
									<td class='DarkBlue Bold MiniText'>Sub Product</td>
									<td class='DarkBlue Bold MiniText'>Model #</td>
									<td class='DarkBlue Bold MiniText'>Start Date</td>
									<td class='DarkBlue Bold MiniText'>End Date</td>
								</tr>
								<tr>
									<td class='spaces'><asp:TextBox class="BGBabyBlue DarkBlue" id="txtProduct" size="18" runat="server" /></td>
									<td class='spaces'><asp:TextBox class="BGBabyBlue DarkBlue" id="txtSubProduct" size="18" runat="server" /></td>
									<td class='spaces'><asp:TextBox class="BGBabyBlue DarkBlue" id="txtModelNo" size="18" runat="server" /></td>
									<td class='spaces'>
										<asp:TextBox class="BGBabyBlue DarkBlue" id="txtStartDate" size="13" maxlength="10" runat="server" />
										<a href="#" id="anchor2" onClick="cal1xx.select(document.forms[0].txtStartDate,'anchor2','MM/dd/yyyy'); return false;">
											<img src="images/calendar.gif" border="0" width="21" height="17"> </a>
									</td>
									<td class='spaces'>
										<asp:TextBox class="BGBabyBlue DarkBlue" id="txtEndDate" size="13" maxlength="10" runat="server" />
										<a href="#" id="anchor1" onClick="cal1xx.select(document.forms[0].txtEndDate,'anchor1','MM/dd/yyyy'); return false;">
											<img src="images/calendar.gif" border="0" width="21" height="17"> </a>
									</td>
								</tr>
								<tr>
									<td>
										<asp:RegularExpressionValidator ControlToValidate="txtStartDate" display="dynamic" ValidationExpression="[0-9][0-9/-]+"
											ErrorMessage="Invalid date" runat="server" id="RegularExpressionValidator1" />
										<asp:CustomValidator ControlToValidate="txtStartDate" OnServerValidate="ValidDate" Text="Invalid date"
											runat="server" id="CustomValidator1" />
									</td>
									<td>
										<asp:RegularExpressionValidator ControlToValidate="txtEndDate" display="dynamic" ValidationExpression="[0-9][0-9/-]+"
											ErrorMessage="Invalid date" runat="server" id="RegularExpressionValidator2" />
										<asp:CustomValidator ControlToValidate="txtEndDate" OnServerValidate="ValidDate" Text="Invalid date"
											runat="server" id="CustomValidator2" />
									</td>
								</tr>
								<tr>
									<td></td>
									<td></td>
									<td></td>
								</tr>
								<tr>
									<td colspan="6" class="BoutonEmplacement"><asp:Button Text="Search" class="Button" runat="server" onClick="Search" id="btnSearch" /></td>
								</tr>
							</table>
							<br>
							<br>
							<asp:Label id="lblQbr" runat="server" />
							<br>
						</td>
						<td class="bg1" width="6%"></td>
					</tr>
				</table>
			</center>
		</form>
		<DIV ID="Calendrier" STYLE="VISIBILITY:hidden;POSITION:absolute;BACKGROUND-COLOR:white;layer-background-color:white"></DIV>
	</body>
</HTML>
