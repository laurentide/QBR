<%@ import namespace = "System" %>
<%@ import namespace = "System.Data" %>
<%@ import namespace = "System.Data.OleDb " %>
<%@ import namespace = "System.Web.UI" %>
<%@ import namespace = "System.Web.UI.WebControls" %>
<%@ import namespace = "System.Security.Principal" %>
<%@ Page Language="vb" Codebehind="Qbr.aspx.vb" Inherits="QBRProject.QBR" %>
<HTML>
	<HEAD>
		<title>QBR</title>
		<script runat="server">
    '|------------------------------------------------------------------------------------------------------------------|
    '| Main : Called on page load																						|
    '|------------------------------------------------------------------------------------------------------------------|
    sub Page_Load

		Dim strWhere as String
		Dim strDate() as String
    	Dim dbConn as oleDbConnection
    	
		EstablishConnection(dbConn)
		dbConn.open()
		ClearSearchSessionVariables()
		
		'defines a session variable to know if we are in write, read, or adding mode
		' 1=add, 2=Modify, 3=Consult
		if request.queryString("Type") <> Nothing then
			Session("Type") = request.queryString("Type").toString()
		end if
		
		Product.Text = "Select Products"		
		
		if Session("Type") <> 1 then
			'If we are in Consult or Modify mode
			QBRcache.text = request.queryString("Nu")
			if not Page.isPostBack then
				CustomerName.Items.clear()
				ContactName.Items.clear()
			end if
			
			if Session("Type") = "2" then
				lblTitle.text = "Modify an existing QBR"
				if not Page.isPostBack then
					'If the page has not post back, show informations from db 
					ShowInfosQbr(dbConn,"W")
					Session("Products") = Nothing
					ResetSessionGraphs()
				else if Session("postCustomer") <> Nothing then
					ContactName.Items.clear()
					CreateListBox(dbConn, "Contact", "Name", "ContactNo", "Name", " where ClientNo=" & CustomerName.SelectedItem.Value, true, ContactName)
					ContactName.SelectedValue = Session("postCustomer")
					ShowContactInfos(dbConn) 
					Session("postCustomer") = Nothing
				end if
			else if Session("Type") = "3" then
				lblTitle.text = "Consult a QBR"
				'Always show informations from the db
				If not Page.isPostBack then
					Session("Products") = Nothing
					ResetSessionGraphs()
					ShowInfosQbr(dbConn,"R")
					Product.Text = "Show Products"
				End If
			end if
		end if
		
		if Session("Type") = "1" then
			lblTitle.text = "Enter a new QBR"
			If not Page.isPostBack then
			'Fill in the Lists
				Session("Products") = Nothing
				ResetSessionGraphs()
				CreateListBox(dbConn, "Employee", "Name", "EMPNo", "Name", "", true, Employe)
				CreateListBox(dbConn, "Employee", "Name", "EMPNo", "Name", " where title='OS'", true, OS)
				CreateListBox(dbConn, "Client", "(Name + ', ' + Cast(ClientNo as Varchar))", "ClientNo", "Name", "", true, CustomerName)  
				CreateListBox(dbConn, "Contact", "Name", "ContactNo", "Name", " where 2=1", true, ContactName)
				CreateListBox(dbConn, "UseQbr", "Name", "UseNo", "Name", "", true, Use)
				Use.selectedindex = 1
				Session("Customer") = ""
				
				' Afficher la date du jour 
				strDate = Split(Format(Now, "d"),"/")
				if strdate(0) < 10 then
                    txtQBRDate.text += "0"
                end if
                txtQBRDate.text += strdate(0) & "/"
                if strdate(1) < 10 then
                    txtQBRDate.text += "0"
                end if
                txtQBRDate.text += strdate(1) & "/"
                txtQBRDate.text += strdate(2)
				
				QBRcache.text = "-1"
			else if Session("postCustomer") <> Nothing then
				ContactName.Items.clear()
				CreateListBox(dbConn, "Contact", "Name", "ContactNo", "Name", " where ClientNo=" & CustomerName.SelectedItem.Value, true, ContactName)
				ContactName.SelectedValue = Session("postCustomer")
				ShowContactInfos(dbConn) 
				Session("postCustomer") = Nothing
			end if			
		end if
		
		if ContactName.selectedItem.Value = "" then
			ContactTitle.text = ""
			ContactEMail.text = ""
			ContactTel.text = ""		
		end if
		
		If not page.ispostBack then
			'Consultation
			if Session("Type") = 3 then
				ShowGraphs(dbConn, QBRcache.text, "R")
			else
				ShowGraphs(dbConn, QBRcache.text, "W")
			end if
		end if
		
		dbConn.close()
	
    end sub

		</script>
		<LINK href="Styles.css" type="text/css" rel="stylesheet">
			<script src="QBR.js" type="text/javascript"></script>
			<script language="JavaScript" src="CalendarPopup.js"></script>
			<LINK href="Calendar.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0" marginwidth="0" marginheight="0">
		<form id="QBR" method="post" runat="server">
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
							<asp:label class="MidText DarkBlue underLine" id="lblTitle" runat="server"></asp:label>
							<table>
								<tr height="30">
									<td></td>
								</tr>
							</table>
							<table class="Bordure" width="100%">
								<tr>
									<td class="bordure DarkBlue" width="19%">Employee</td>
									<td class="bordure1" width="31%"><asp:dropdownlist class="DarkBlue BGBabyBlue" id="Employe" runat="server"></asp:dropdownlist></td>
									<td class="bordure DarkBlue" width="19%">Outside Salesman</td>
									<td class="bordure1" width="31%"><asp:dropdownlist class="DarkBlue BGBabyBlue" id="OS" runat="server"></asp:dropdownlist></td>
								</tr>
								<% if Session("Type")=3 then %>
								<tr>
									<td class="bordure DarkBlue" height="30">Enterprise</td>
									<td class="bordure1" colSpan="3" height="30"><asp:dropdownlist class="DarkBlue BGBabyBlue" id="Corporation" runat="server" autopostback="true"></asp:dropdownlist></td>
								</tr>
								<% end if %>
								<tr>
									<td class="bordure DarkBlue">Customer Name</td>
									<td class="bordure1"><asp:dropdownlist class="DarkBlue BGBabyBlue" id="CustomerName" runat="server" autopostback="true"
											OnSelectedIndexChanged="ShowContacts"></asp:dropdownlist></td>
									<td class="bordure DarkBlue">Contact Name</td>
									<td class="bordure1"><asp:dropdownlist class="DarkBlue BGBabyBlue" id="ContactName" runat="server" AutoPostBack="true"
											onSelectedIndexChanged="CallsShowContactInfos" readOnly="true"></asp:dropdownlist>
										<% if Session("type") <> 3 then %>
										&nbsp;&nbsp;
										<asp:button class="SmallButton" id="btnNewContact" onclick="openNewContact" runat="server" text="New"></asp:button><asp:button class="SmallButton" id="btnEditContact" onclick="EditContact" runat="server" text="Edit"></asp:button>
										<%end if%>
									</td>
								</tr>
								<tr>
									<td class="bordure DarkBlue">Address</td>
									<td class="bordure1"><asp:textbox class="DarkBlue BGBabyBlue" id="Address" runat="server" readOnly="true" size="30"></asp:textbox></td>
									<td class="bordure DarkBlue">Contact title</td>
									<td class="bordure1"><asp:textbox class="DarkBlue BGBabyBlue" id="ContactTitle" runat="server" readOnly="true" size="30"></asp:textbox></td>
								</tr>
								<tr>
									<td class="bordure DarkBlue"></td>
									<td class="bordure1"><asp:textbox class="DarkBlue BGBabyBlue" id="Location" runat="server" readOnly="true" size="30"></asp:textbox></td>
									<td class="bordure DarkBlue">Contact eMail</td>
									<td class="bordure1"><asp:textbox class="DarkBlue BGBabyBlue" id="ContactEMail" runat="server" readOnly="true" size="30"></asp:textbox></td>
								</tr>
								<tr>
									<td class="bordure DarkBlue">Industry</td>
									<td class="bordure1"><asp:textbox class="DarkBlue BGBabyBlue" id="Industry" runat="server" readOnly="true" size="30"></asp:textbox></td>
									<td class="bordure DarkBlue">Contact tel #</td>
									<td class="bordure1"><asp:textbox class="DarkBlue BGBabyBlue" id="ContactTel" runat="server" readOnly="true" size="30"></asp:textbox></td>
								</tr>
								<tr>
									<td class="bordure DarkBlue" colSpan="2" height="23">Where can we use the QBR?</td>
									<td class="bordure1" colSpan="2" height="23"><asp:dropdownlist class="DarkBlue BGBabyBlue" id="Use" runat="server"></asp:dropdownlist></td>
								</tr>
								<tr>
									<td class="bordure DarkBlue" colSpan="2">Input Type</td>
									<td class="bordure1" colSpan="2"><asp:dropdownlist class="DarkBlue BGBabyBlue" id="Input" runat="server">
											<asp:ListItem></asp:ListItem>
											<asp:ListItem>QBR</asp:ListItem>
											<asp:ListItem>Application Note</asp:ListItem>
											<asp:ListItem>Customer Reference / Testimonial</asp:ListItem>
										</asp:dropdownlist></td>
								</tr>
								<tr>
									<td class="BGWhite" colSpan="4">&nbsp;</td>
								</tr>
								<tr>
									<td class="bordure DarkBlue">Date</td>
									<td class="bordure1 spaces" colSpan="3"><asp:textbox class="DarkBlue BGBabyBlue" id="txtQbrDate" runat="server" size="29" maxlength="10"></asp:textbox>
										<% if Session("Type") <> "3" then %>
										&nbsp;&nbsp;&nbsp;&nbsp; <A id="anchor1" onclick="cal1xx.select(document.forms[0].txtQbrDate,'anchor1','MM/dd/yyyy'); return false;"
											href="#"><IMG height="17" src="images/calendar.gif" width="21" border="0"></A>
										<% end if %>
										&nbsp;&nbsp;&nbsp;&nbsp;
										<asp:regularexpressionvalidator id="ValDate" runat="server" ErrorMessage="Invalid date" ValidationExpression="[0-9][0-9/-]+"
											display="dynamic" ControlToValidate="txtQbrDate"></asp:regularexpressionvalidator><asp:customvalidator id="Customvalidator1" runat="server" ControlToValidate="txtQbrDate" Text="Invalid date"
											OnServerValidate="ValidDate"></asp:customvalidator></td>
								</tr>
								<tr>
									<td class="bordure DarkBlue">Project Name</td>
									<td class="bordure1" colSpan="3"><asp:textbox class="DarkBlue BGBabyBlue" id="ProjectName" runat="server" size="84" maxlength="100"></asp:textbox></td>
								</tr>
								<tr>
									<td class="bordure DarkBlue">Application</td>
									<td class="bordure1" colSpan="3"><asp:textbox class="DarkBlue BGBabyBlue" id="QBRApplication" runat="server" size="84" maxlength="100"></asp:textbox></td>
								</tr>
								<tr>
									<td class="bordure DarkBlue">Products</td>
									<td class="bordure" colSpan="4"><asp:button class="SmallButton" id="Product" onclick="openProducts" runat="server"></asp:button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<span class="DarkBlue MiniText">* Changes will 
            apply when the QBR is saved </span></td>
								</tr>
								<tr>
									<td class="BGWhite" colSpan="4">&nbsp;</td>
								</tr>
								<tr>
									<td class="bordure DarkBlue" colSpan="4">Original Situation / Challenge (include 
										sketch, photos)</td>
								</tr>
								<tr>
									<td class="bordure1" colSpan="4"><asp:textbox class="DarkBlue BGBabyBlue" id="Situation" runat="server" cols="115" rows="5" TextMode="MultiLine"></asp:textbox><br>
										<br>
										<% if Session("Type") <> "3" then %>
										<span class="DarkBlue">Add a File: </span><input id="MyFileSituation" type="file" size="87" name="MyFileSituation" runat="server">&nbsp;&nbsp;&nbsp;
										<asp:button id="LinksSituation" onclick="AddFile" runat="server" text="Add"></asp:button><br>
										<span class="DarkBlue MiniText">* File size must be smaller than 50 Mo<span>
												<br>
												<asp:label class="erreur" id="ErrorLinkSituation" runat="server"></asp:label>
												<br>
												<% End If %>
												<asp:label id="lblLinksSituation" runat="server"></asp:label>
												<% If lblLinksSituation.text <> "" and Session("Type") <> "3" then %>
												<asp:button class="SmallButton" id="DeleteGraphLinksSituation0" onclick="DeleteGraphs" runat="server"
													text="Delete"></asp:button>
												<% End If %>
											</span></span></td>
								</tr>
								<tr>
									<td class="BGWhite" colSpan="4">&nbsp;</td>
								</tr>
								<tr>
									<td class="bordure DarkBlue" colSpan="4">Solution (include sketch, photos)</td>
								</tr>
								<tr>
									<td class="bordure1" colSpan="4"><asp:textbox class="DarkBlue BGBabyBlue" id="Solution" runat="server" cols="115" rows="5" TextMode="MultiLine"></asp:textbox><br>
										<br>
										<% if Session("Type") <> "3" then %>
										<span class="DarkBlue">Add a File: </span><input id="MyFileSolution" type="file" size="87" name="MyFileSolution" runat="server">&nbsp;&nbsp;&nbsp;
										<asp:button id="LinksSolution" onclick="AddFile" runat="server" text="Add"></asp:button><br>
										<span class="DarkBlue MiniText">* File size must be smaller than 50 Mo<span>
												<br>
												<asp:label class="erreur" id="ErrorLinkSolution" runat="server"></asp:label>
												<br>
												<% End If %>
												<asp:label id="lblLinksSolution" runat="server"></asp:label>
												<% If lblLinksSolution.text <> "" and Session("Type") <> "3" then %>
												<asp:button class="SmallButton" id="DeleteGraphLinksSolution1" onclick="DeleteGraphs" runat="server"
													text="Delete"></asp:button>
												<% End If %>
											</span></span></td>
								</tr>
								<tr>
									<td class="BGWhite" colSpan="4">&nbsp;</td>
								</tr>
								<tr>
									<td class="bordure DarkBlue" colSpan="4">Improved Business Results Calculation</td>
								</tr>
								<tr>
									<td class="bordure1" colSpan="4"><asp:textbox class="DarkBlue BGBabyBlue" id="Results" runat="server" cols="115" rows="5" TextMode="MultiLine"></asp:textbox><br>
										<br>
										<% if Session("Type") <> "3" then %>
										<span class="DarkBlue">Add a File: </span><input id="MyFileResult" type="file" size="87" name="MyFileResult" runat="server">&nbsp;&nbsp;&nbsp;
										<asp:button id="LinksResult" onclick="AddFile" runat="server" text="Add"></asp:button><br>
										<span class="DarkBlue MiniText">* File size must be smaller than 50 Mo<span>
												<br>
												<asp:label class="erreur" id="ErrorLinkResult" runat="server"></asp:label>
												<br>
												<% End If %>
												<asp:label id="lblLinksResult" runat="server"></asp:label>
												<% If lblLinksResult.text <> "" and Session("Type") <> "3" then %>
												<asp:button class="SmallButton" id="DeleteGraphLinksResult2" onclick="DeleteGraphs" runat="server"
													text="Delete"></asp:button>
												<% End If %>
											</span></span></td>
								</tr>
								<tr>
									<td class="DarkBlue" colSpan="4">
										<table>
											<tr>
												<td class="SmallText Darkblue">Costs:
												</td>
												<td class="SmallText Darkblue">$<asp:textbox class="DarkBlue BGBabyBlue" id="Cost" onblur="CalculROI();" runat="server" size="8"
														maxlength="14"></asp:textbox>
												</td>
												<td class="SmallText Darkblue">&nbsp;&nbsp;&nbsp;One time savings:
												</td>
												<td class="SmallText Darkblue">$<asp:textbox class="DarkBlue BGBabyBlue" id="OnceSavings" onblur="CalculROI();" runat="server"
														size="8" maxlength="14"></asp:textbox>
												</td>
												<td class="SmallText Darkblue">&nbsp;&nbsp;&nbsp;First year ROI:
												</td>
												<td class="SmallText Darkblue"><asp:textbox class="DarkBlue BGBabyBlue" id="ROI" runat="server" readOnly="true" size="4"></asp:textbox>%
												</td>
												<td class="SmallText Darkblue">&nbsp;&nbsp;&nbsp;Approved by Customer:
												</td>
												<td><asp:checkbox class="DarkBlue BGBabyBlue" id="Approved" runat="server"></asp:checkbox></td>
											</tr>
											<tr>
												<td></td>
												<td></td>
												<td class="SmallText Darkblue">&nbsp;&nbsp;&nbsp;Annual savings:</td>
												<td class="SmallText Darkblue">$<asp:textbox class="DarkBlue BGBabyBlue" id="AnnualSavings" onblur="CalculROI();" runat="server"
														size="8" maxlength="14"></asp:textbox>
												</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td class="BGWhite" colSpan="4">&nbsp;</td>
								</tr>
								<tr>
									<td class="bordure DarkBlue" colSpan="4">Customer Testimonial / Quote</td>
								</tr>
								<tr>
									<td class="bordure1" colSpan="4"><asp:textbox class="DarkBlue BGBabyBlue" id="Testimonial" runat="server" cols="115" rows="5"
											TextMode="MultiLine"></asp:textbox></td>
								</tr>
							</table>
							<br>
							<center><asp:textbox id="Qbrcache" runat="server" visible="false"></asp:textbox><asp:button class="button" id="btnEmail" onclick="SendMailOutlook" runat="server" text="Email"></asp:button>
								<% if Session("Type") <> "3" then %>
								<asp:button class="button" id="Button1" onclick="SaveQbr" runat="server" text="Save"></asp:button>
								<%else%>
								<asp:button class="button" id="Button2" onclick="openPrint" runat="server" text="Printable version"></asp:button>
								<%end if%>
							</center>
						</td>
						<td class="bg1" width="6%"></td>
					</tr>
				</table>
			</center>
		</form>
		<DIV id="Calendrier" style="VISIBILITY: hidden; POSITION: absolute; BACKGROUND-COLOR: white; layer-background-color: white"></DIV>
	</body>
</HTML>
