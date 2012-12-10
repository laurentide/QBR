<%@ import namespace = "System.Web.UI.WebControls" %>
<%@ import namespace = "System.Web.UI" %>
<%@ import namespace = "System.Data.OleDb " %>
<%@ import namespace = "System.Data" %>
<%@ import namespace = "System" %>
<%@ Page Language="vb" Codebehind="Qbr.aspx.vb" Inherits="QBRProject.QBR"%>
<HTML>
	<HEAD>
		<title>Contact</title>
		<script runat="server">
    '|------------------------------------------------------------------------------------------------------------------|
    '| Main : Called on page load																						|
    '|------------------------------------------------------------------------------------------------------------------|
	
    sub Page_Load
		Dim strReq   As string
		Dim dbConn as OleDbConnection
			
		Dim cmdTable As oleDbDataAdapter
		Dim dTable   As dataTable
		
		EstablishConnection(dbConn)
		dbConn.open
			
		if request.queryString("Con") <> Nothing then
			If not Page.IsPostBack then
				'Show informations on the contact
				strReq = "Select * from contact where ContactNo = " & request.queryString("Con")
				
				cmdTable = New oleDbDataAdapter(strReq, dbConn)
				dTable   = New dataTable()

				'Execute request
				cmdTable.fill(dTable)

				With dTable.Rows.Item(0)
					If Not isDbNull(.Item("Name")) Then
						txtContactName.text = .Item("Name")
					Else
						txtContactName.text = ""
					End If
					
					If Not isDbNull(.Item("Title")) Then
						ContactTitle.text = .Item("Title")
					Else
						ContactTitle.text = ""
					End If
				
					If Not isDbNull(.Item("Phone")) Then
						ContactTel.text = "(" & Mid(.Item("Phone"),1,3) & ") " &  Mid(.Item("Phone"),4,3) & "-" & _
											Mid(.Item("Phone"),7,4)
					Else
						ContactTel.text = ""
					End If
					
					If Not isDbNull(.Item("Email")) Then
						ContactEmail.text = .Item("Email")
					Else
						ContactEmail.text = ""
					End If
				End With
			End if
			lblContact.text = "Modify an existing contact"
		else
			lblContact.text = "Add a new Contact"	
		end if
		
		dbConn.close
	end sub
	
		</script>
		<link rel="stylesheet" type="text/css" href="Styles.css">
			<script>
			//places current window in the center
			var xpos = (screen.width - 600) / 2
			var ypos = (screen.height - 305) / 2
			moveTo(xpos, ypos);
			</script>
	</HEAD>
	<body>
		<asp:Label class="midText underline darkblue" id="lblContact" runat="server" />
		<br>
		<br>
		<span class="smalltext darkblue bold">
				Please enter the Contact's name and all the information you know about him. 
			</span>
		<br>
		<form runat="server" method="post">
			<table class="Bordure" width="100%">
				<tr valign="top">
					<td class="bordure DarkBlue" width="22%">Contact's name</td>
					<td class="bordure1" width="28%">
						<asp:textBox class="DarkBlue BGBabyBlue" id="txtcontactName" size="20" maxlength="50" runat="server" />
						<asp:RequiredFieldValidator ControlToValidate="txtcontactName" Text="Name required" runat="server" id="RequiredFieldValidator1" />
					</td>
					<td class="bordure DarkBlue" width="22%">Title</td>
					<td class="bordure1" width="28%"><asp:textBox class="DarkBlue BGBabyBlue" id="contactTitle" size="20" maxlength="50" runat="server" /></td>
				</tr>
				<tr valign="top">
					<td class="bordure DarkBlue">eMail</td>
					<td class="bordure1">
						<asp:textBox class="DarkBlue BGBabyBlue" id="contactEMail" size="20" maxlength="50" runat="server" />
						<asp:customValidator maxlength="256" controltoValidate="contactEMail" onServerValidate="ValidEMail" errormessage="Invalid eMail"
							runat="server" id="CustomValidator1" />
					</td>
					<td class="bordure DarkBlue">*Phone #
					</td>
					<td class="bordure1">
						<asp:textBox class="DarkBlue BGBabyBlue" id="contactTel" size="20" maxlength="14" runat="server" />
						<asp:RegularExpressionValidator ControlToValidate="contactTel" ValidationExpression="^\(\d{3}\)\s\d{3}\-\d{4}$"
							ErrorMessage="Invalid phone #" runat="server" id="RegularExpressionValidator1" />
					</td>
				</tr>
			</table>
			<br>
			<br>
			<center>
				<% If request.queryString("Con") = Nothing Then %>
				<asp:Button class="button" text="Add contact" id="btnOk" onClick="AddContact" size="20" maxlength="50"
					runat="server" />
				<% Else %>
				<asp:Button class="button" text="Update contact" id="btnOk1" onClick="UpdateContact" size="20"
					maxlength="50" runat="server" />
				<% End If %>
				&nbsp;&nbsp;&nbsp; <input class="button" type="button" value="Cancel" onclick="javascript:opener.document.forms[0].submit();self.close();"
					size="20" maxlength="50">
			</center>
			<br>
			<br>
			<span class="darkblue minitext">* Phone number must be in format: (111) 111-1111</span>
		</form>
	</body>
</HTML>
