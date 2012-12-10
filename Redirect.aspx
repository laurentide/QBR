<%@ Page Language="vb" Codebehind="Qbr.aspx.vb" Inherits="QBRProject.QBR" %>
<HTML>
	<HEAD>
		<title>QBR</title>
		<script runat="server">
			sub Page_Load
				lblLink.text = "You will be redirected to the <a href=Qbr.aspx?Nu=" & Request.QueryString("Nu") & ">Qbr</a> in 5 seconds"
			
				response.write("<")
				response.write("script>setTimeout(""Redirige(" & Request.QueryString("Nu") & ")"", 5000);</script")
				response.write(">")
				
			end sub
		</script>
		<script language="javascript">
		
			function Redirige(qbrNo) {
				self.location = "Qbr.aspx?Nu=" + qbrNo;	
			}
			
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<center>
			<br />
			<br />
			<h2>
				Your Email has been sent to the customer.
				<br />
				<br />
				<asp:Label ID="lblLink" Runat="server" />
			</h2>
		</center>
	</body>
</HTML>
