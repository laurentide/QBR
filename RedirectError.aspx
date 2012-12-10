<%@ Page Language="vb" Codebehind="Qbr.aspx.vb" Inherits="QBRProject.QBR" %>

<HTML>
  <HEAD>
		<title>QBR</title>
		<script runat="server">
			sub Page_Load
				lblLink.text = "You will be redirected to the <a href=Email.aspx?Nu=" & Request.QueryString("Nu") & ">Email form</a> in 5 seconds"
			
				response.write("<")
				response.write("script>setTimeout(""Redirige(" & Request.QueryString("Nu") & ")"", 5000);</script")
				response.write(">")
				
			end sub
		</SCRIPT>
		<script language="javascript">
		
			function Redirige(qbrNo) {
				self.location = "Email.aspx?Nu=" + qbrNo;	
			}
			
		</script>
</HEAD>
		
	<body MS_POSITIONING="GridLayout">
<TABLE height=132 cellSpacing=0 cellPadding=0 width=951 border=0 
ms_2d_layout="TRUE">
  <TR vAlign=top>
    <TD width=10 height=15></TD>
    <TD width=941></TD></TR>
  <TR vAlign=top>
    <TD height=117></TD>
    <TD>
		<center>
		<br > <br > 
			<h2>
				An Error Occured while sending your email. <br > <br > 
				<asp:Label ID="lblLink" Runat="server" />
			</h2>
		</center></TD></TR></TABLE>
	</body>
</HTML>
