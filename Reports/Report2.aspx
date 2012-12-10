<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Report2.aspx.vb" Inherits="Reports.WebForm1"%>
<%@ Register TagPrefix="cr" Namespace="CrystalDecisions.Web" Assembly="CrystalDecisions.Web, Version=9.1.5000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>WebForm1</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<cr:CrystalReportViewer id="CRViewReport2" style="Z-INDEX: 101; LEFT: -96px; POSITION: absolute; TOP: 48px"
				runat="server" Width="350px" Height="50px"></cr:CrystalReportViewer>
		</form>
	</body>
</HTML>
