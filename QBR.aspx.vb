Imports System
Imports System.IO
Imports System.Net
Imports System.Math
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports Microsoft.VisualBasic
Imports System.Web.Mail
Imports System.Collections
Imports Microsoft.Office.Interop
Imports System.Web.SessionState


Public Class QBR
    Inherits System.Web.UI.Page

    ' List of needed variables or objects
    Protected lblQbr As Label

    Protected txtCorporation As TextBox
    Protected txtCustomer As TextBox
    Protected txtIndustry As TextBox
    Protected txtApplication As TextBox
    Protected txtProduct As TextBox
    Protected txtSubProduct As TextBox
    Protected txtModelNo As TextBox
    Protected txtEmploye As TextBox
    Protected txtStartDate As TextBox
    Protected txtEndDate As TextBox
    Protected txtOrder As TextBox

    Protected CustomerName As DropDownList
    Protected Corporation As DropDownList
    Protected ContactName As DropDownList
    Protected Employe As DropDownList
    Protected OS As DropDownList
    Protected Use As DropDownList
    Protected Input As DropDownList

    Protected SubProduct As RadioButtonList

    Protected Address As TextBox
    Protected Location As TextBox
    Protected Industry As TextBox

    Protected ProjectName As TextBox
    Protected txtQbrDate As TextBox
    Protected QBRApplication As TextBox
    Protected Situation As TextBox
    Protected Solution As TextBox
    Protected Testimonial As TextBox

    Protected Results As TextBox
    Protected Cost As TextBox
    Protected OnceSavings As TextBox
    Protected AnnualSavings As TextBox
    Protected ROI As TextBox
    Protected Approved As CheckBox
    Protected WithEvents ContactTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents ContactEMail As System.Web.UI.WebControls.TextBox
    Protected WithEvents ContactTel As System.Web.UI.WebControls.TextBox


    Protected txtcontactName As TextBox
    Protected WithEvents Qbrcache As System.Web.UI.WebControls.TextBox


    Protected lblEmploye As Label
    Protected lblOS As Label
    Protected lblCustomerName As Label
    Protected lblAddress As Label
    Protected lblContactName As Label
    Protected lblContactTitle As Label
    Protected lblLocation As Label
    Protected WithEvents lblContactEMail As System.Web.UI.WebControls.Label
    Protected lblIndustry As Label
    Protected lblContactTel As Label
    Protected WithEvents lblDateQBR As System.Web.UI.WebControls.Label
    Protected lblProjectName As Label
    Protected lblApplication As Label
    Protected lblProduct As Label
    Protected lblDescProd As Label
    Protected lblSituation As Label
    Protected lblSolution As Label
    Protected lblResults As Label
    Protected lblCost As Label
    Protected lblOnceSavings As Label
    Protected lblAnnualSavings As Label
    Protected lblROI As Label
    Protected lblApproved As Label
    Protected lblTestimonial As Label
    Protected lblCorporation As Label
    Protected lblUse As Label
    Protected lblInputType As Label

    Protected ErrorLinkSituation As Label
    Protected ErrorLinkSolution As Label
    Protected ErrorLinkResult As Label
    Protected lblLinksSituation As Label
    Protected lblLinksSolution As Label
    Protected lblLinksResult As Label

    Protected MyFileSituation As HtmlInputFile
    Protected MyFileSolution As HtmlInputFile
    Protected MyFileResult As HtmlInputFile

    Protected btnPrint As Button
    Protected Product As Button

    Protected ph1 As PlaceHolder

    Protected ValDate As RegularExpressionValidator
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents lblTitle As System.Web.UI.WebControls.Label
    Protected WithEvents btnNewContact As System.Web.UI.WebControls.Button
    Protected WithEvents btnEditContact As System.Web.UI.WebControls.Button
    Protected WithEvents Customvalidator1 As System.Web.UI.WebControls.CustomValidator
    Protected WithEvents LinksSituation As System.Web.UI.WebControls.Button
    Protected WithEvents DeleteGraphLinksSituation0 As System.Web.UI.WebControls.Button
    Protected WithEvents LinksSolution As System.Web.UI.WebControls.Button
    Protected WithEvents DeleteGraphLinksSolution1 As System.Web.UI.WebControls.Button
    Protected WithEvents LinksResult As System.Web.UI.WebControls.Button
    Protected WithEvents DeleteGraphLinksResult2 As System.Web.UI.WebControls.Button
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents Button3 As System.Web.UI.WebControls.Button

    Protected dgQbr As DataGrid
    Protected btnEmail As Button
    Protected WithEvents txtFrom As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtTo As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCC As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtBCC As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtText As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtSubject As System.Web.UI.WebControls.TextBox

    Protected WithEvents NoQbr As System.Web.UI.WebControls.TextBox
    Protected WithEvents SendEmail As System.Web.UI.WebControls.Button
    Protected WithEvents More As System.Web.UI.WebControls.Button
    Protected WithEvents lblTextMail As System.Web.UI.WebControls.Label
    Protected WithEvents lblSignature As System.Web.UI.WebControls.Label
    Protected WithEvents RegularExpressionValidator1 As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents RegularExpressionValidator2 As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents CustomValidator2 As System.Web.UI.WebControls.CustomValidator
    Protected WithEvents btnSearch As System.Web.UI.WebControls.Button
    Protected WithEvents lblContact As System.Web.UI.WebControls.Label
    Protected WithEvents RequiredFieldValidator1 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents btnOk As System.Web.UI.WebControls.Button
    Protected WithEvents btnOk1 As System.Web.UI.WebControls.Button
    Protected WithEvents lblLink As System.Web.UI.WebControls.Label


    Protected txtSignature As TextBox

    '|------------------------------------------------------------------------------------------------------------------|
    '| EstablishConnection: Establishes the connection to the database		                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub EstablishConnection(ByRef dbConn As OleDbConnection)
        Dim strConn As String = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=QBR;User ID=test;Password=test"

        dbConn = New OleDbConnection(strConn)
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowQbr: Show the list of Qbrs in the database dependending on the search                                        | 
    '|		fields entered or the order asked		                                                                    |
    '|      Parameters:                                                                                                 |
    '|      StrWhere: If there are any constraints we would like to add                                                 |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowQbr(ByVal dbConn As OleDbConnection, ByVal strWhere As String)
        Dim strReq As String
        Dim strOrder As String
        Dim strWay As String
        Dim strImg As String
        Dim strColumnNames() As String = {"&nbsp;&nbsp;&nbsp;#&nbsp;&nbsp;&nbsp;", "Enterprise", "Customer", "Contact", "Industry", "Application", "Employee", "Date"}

        Dim index As Integer = 0

        Dim cmdTable As OleDbDataAdapter
        Dim dsTable As DataSet
        Dim myRow As DataRow
        Dim myCol As DataColumn

        If Session("Type") = 2 Then
            strWhere += " and Approved=0"
        End If

        strOrder = request.QueryString("Order")
        strWay = request.QueryString("Way")

        strReq = "SELECT QBR.QBRNo AS [QBR.QBRNo], Corporation.Name AS [Corporation.Name], Client.Name AS [Client.Name], " & _
                    "Contact.Name AS [Contact.Name], Industry.Industry, QBR.Application, Employee.Name AS [Employee.Name], QBR.QbrDate " & _
                    "FROM ((Industry RIGHT JOIN (Employee INNER JOIN (Corporation INNER JOIN ((Client INNER JOIN Contact " & _
                    "ON Client.ClientNo = Contact.ClientNo) INNER JOIN QBR ON Contact.ContactNo = QBR.ContactNo) ON " & _
                    "Corporation.CorporationNo = Client.CorporationNo) ON Employee.EmpNo = QBR.EmpNo) ON Industry.industryNo " & _
                    "= Client.IndustryNo) LEFT JOIN (Product_Service RIGHT JOIN Product ON Product_Service.ProductNo = Product.ProductNo) " & _
                    "ON QBR.QBRNo = Product.QBRNo) LEFT JOIN Result ON QBR.QBRNo = Result.QBRNo " & _
                    "WHERE 1=1 " & PutsCaracters(strWhere) & _
                    " GROUP BY QBR.QBRNo, Corporation.Name, Client.Name, Contact.Name, Industry.Industry, " & _
                    "QBR.Application, Employee.Name, QBR.QbrDate " & _
                    "ORDER BY " & StrOrder & " " & StrWay

        cmdTable = New OleDbDataAdapter(strReq, dbConn)
        dsTable = New DataSet
        'Execute request
        cmdTable.Fill(dsTable, "QBR")

        'Filling the html table
        If dsTable.Tables("QBR").Rows.Count <> 0 Then
            lblQbr.Text = "<table class='Bordure'>"
            lblQbr.Text += "<tr>"
            For Each myCol In dsTable.Tables("QBR").Columns
                If strOrder = myCol.ToString() Then
                    If request.QueryString("Way") = "" Then
                        StrWay = "DESC"
                        strImg = "up"
                    Else
                        StrWay = ""
                        strImg = "down"
                    End If
                    strImg = "&nbsp; <img src='images\" & strImg & ".gIf' border=0 />"
                Else
                    strWay = ""
                    strImg = ""
                End If
                lblQbr.Text += "<td class='Bordure padding Bold DarkBlue SmallText'><a href='search.aspx?Type=" & _
                                Trim(Session("Type")) & "&Where=" & Replace(strWhere, "#", "@*@") & "&Order=" & _
                                mycol.ToString() & "&Way=" & strWay & "'>" & strColumnNames(index) & "</a>"
                lblQbr.Text += strImg
                lblQbr.Text += "</td>"
                index += 1
            Next
            lblQbr.Text += "</tr>"

            For Each myRow In dsTable.Tables("QBR").Rows
                lblQbr.Text += "<tr>"
                For Each myCol In dsTable.Tables("QBR").Columns
                    If myCol.ToString() = "QBR.QBRNo" Then
                        lblQbr.Text += "<td align='center' class='Bordure1 Bold padding Small' align='left'><div class='CorporateLink'>" & _
                          "<a href='Qbr.aspx?Nu=" & myRow(myCol) & "&Type=" & Session("Type") & "'>&nbsp;&nbsp;&nbsp;" & myRow(myCol) & _
                                               "&nbsp;&nbsp;&nbsp;</a></div></td>"
                    Else
                        lblQbr.Text += "<td class='Bordure1 padding DarkBlue MiniText' align='left'>" & myRow(myCol) & "</td>"
                    End If
                Next
                lblQbr.Text += "</tr>"
            Next
            lblQbr.Text += "</table>"
        Else
            lblQbr.Text = "<span class='DarkBlue SmallText'>Your search returned no results.</span>"
        End If

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| CreateListBox: Creates a DropDownList with values comming from a field in a table                                |
    '| Parameters: strTable: the table where the searched field is                                                      |
    '|             strField: the name of the searched field                                                             |
    '|             strIndex: The value we want to give to the field to use As an index                                  |
    '|             strWhere: If there are any contraints we would like to add                                           |
    '|             Blank:    If we put a blank field at the beginning                                                   |
    '|             list:     The name of the DropDownList we want items to be added                                     |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub CreateListBox(ByVal dbConn As OleDbConnection, ByVal strTable As String, ByVal strField As String, ByVal strIndex As String, _
    ByVal strAs As String, ByVal strWhere As String, ByVal blank As Boolean, _
                                                    ByRef list As DropDownList)

        Dim strReq As String = "Select " & strIndex & ", " & strField & " AS [" & strAs & "] from " & strTable & _
                                      strWhere & " group by " & strAs & ", " & strIndex & " order by " & strAs

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet
        Dim myRow As DataRow

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        'Filling the list
        If blank Then
            list.Items.Add(New ListItem("", ""))
        End If

        For Each myRow In dsTable.Tables(strTable).Rows
            If Not isDbNull(myRow(strAs)) Then
                list.Items.Add(New ListItem(myRow(strAs), myRow(strIndex)))
            End If
        Next

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| CreateLabel: Creates a Label with values comming from a field in a table                                         |
    '| Parameters: strTable: the table where the searched field is                                                      |
    '|             strField: the name of the searched field                                                             |
    '|             strWhere: If there are any contraints we would like to add                                           |
    '|             field:     The name of the DropDownList we want items to be added                                    |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub CreateLabel(ByVal dbConn As OleDbConnection, ByVal strTable As String, ByVal strField As String, ByVal strWhere As String, _
                        ByRef field As Label)

        Dim strReq As String = "Select " & strField & " from " & strTable & strWhere & " group by " & strField & _
                                          " order by " & strField

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        With dsTable.Tables(strTable).Rows.Item(0)
            If Not isDbNull(.Item(strField)) Then
                field.Text = .Item(strField)
            End If
        End With

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| Search: Makes the search depending on what the user asked				                                        |
    '|------------------------------------------------------------------------------------------------------------------|	
    Sub Search(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dbConn As OleDbConnection
        Dim strTab() As String
        Dim strDate As String
        Dim strWhere As String = ""

        EstablishConnection(dbConn)
        dbConn.Open()


        Session("Customer") = txtCustomer.Text
        If txtCustomer.Text <> "" Then
            strWhere += " AND Client.Name like @*@%" & Replace(txtCustomer.Text, """", """""") & "%@*@"
        End If

        Session("Industry") = txtIndustry.Text
        If txtIndustry.Text <> "" Then
            strWhere += " AND Industry.Industry like @*@%" & Replace(txtIndustry.Text, """", """""") & "%@*@"
        End If

        Session("Application") = txtApplication.Text
        If txtApplication.Text <> "" Then
            strWhere += " AND QBR.Application like @*@%" & Replace(txtApplication.Text, """", """""") & "%@*@"
        End If

        Session("Corporation") = txtCorporation.Text
        If txtCorporation.Text <> "" Then
            strWhere += " AND Corporation.Name like @*@%" & Replace(txtCorporation.Text, """", """""") & "%@*@"
        End If

        Session("Product") = txtProduct.Text
        If txtProduct.Text <> "" Then
            strWhere += " AND Product_service.PrimaryP like @*@%" & Replace(txtProduct.Text, """", """""") & "%@*@"
        End If

        Session("subProduct") = txtSubProduct.Text
        If txtSubProduct.Text <> "" Then
            strWhere += " AND Product_service.SecondaryP like @*@%" & Replace(txtSubProduct.Text, """", """""") & "%@*@"
        End If

        Session("ModelNo") = txtModelNo.Text
        If txtModelNo.Text <> "" Then
            strWhere += " AND Product.ModelNo like @*@%" & Replace(txtModelNo.Text, """", """""") & "%@*@ AND Product.ModelNo <> @*@NULL@*@"
        End If

        Session("Employee") = txtEmploye.Text
        If txtEmploye.Text <> "" Then
            strWhere += " AND Employee.Name like @*@%" & Replace(txtEmploye.Text, """", """""") & "%@*@"
        End If

        Session("StartDate") = txtStartDate.Text
        Session("EndDate") = txtEndDate.Text
        If txtStartDate.Text <> "" And txtEndDate.Text <> "" Then               'between
            If IsDate(txtStartDate.Text) And IsDate(txtEndDate.Text) Then
                strWhere += " AND QBR.QBRDate Between @*@" & txtStartDate.Text & "@*@ And @*@" & txtEndDate.Text & "@*@"
            End If
        ElseIf txtStartDate.Text <> "" Then                                     '>
            If IsDate(txtStartDate.Text) Then
                strWhere += " AND QBR.QBRDate >= @*@" & txtStartDate.Text & "@*@"
            End If
        ElseIf txtEndDate.Text <> "" Then                                       '< 
            If IsDate(txtEndDate.Text) Then
                strWhere += " AND QBR.QBRDate <= @*@" & txtEndDate.Text & "@*@"
            End If
        End If

        If page.IsValid Then
            ShowQbr(dbConn, strWhere)
        End If

        dbConn.Close()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowCustomerInfos: Shows customer infos on Qbr               		                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowCustomerInfos(ByVal dbConn As OleDbConnection)
        Dim strReq As String = "Select Address, City, Province, Industry from Client, Industry where " & _
                                   " Client.IndustryNo = Industry.IndustryNo AND ClientNo=" & CustomerName.SelectedItem.Value

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet
        Dim strTable = "Client"

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        With dsTable.Tables(strTable).Rows.Item(0)
            If Not IsDBNull(.Item("Address")) Then
                Address.Text = .Item("Address")
            Else
                Address.Text = ""
            End If

            Location.Text = ""
            If Not IsDBNull(.Item("City")) Then
                Location.Text += .Item("City")
            End If

            If Not IsDBNull((.Item("Province"))) Then
                If Location.Text <> "" Then
                    Location.Text += ", "
                End If
                Location.Text += .Item("Province")
            End If

            If Not IsDBNull(.Item("Industry")) Then
                Industry.Text = .Item("Industry")
            Else
                Industry.Text = ""
            End If
        End With
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowContactInfos: Shows contact infos on Qbr                                                                     |
    '|                      isDbNull : VB functions that returns whether or not the field from the db is NULL           |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowContactInfos(ByVal dbConn As OleDbConnection)
        Dim strReq As String = "Select Title, EMail, Phone from Contact where ContactNo=" & _
                                    ContactName.SelectedItem.Value

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet
        Dim strTable = "Contact"

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        With dsTable.Tables(strTable).Rows.Item(0)
            If Not isDbNull(.Item("Title")) Then
                ContactTitle.Text = .Item("Title")
            Else
                ContactTitle.Text = ""
            End If

            If Not isDbNull(.Item("EMail")) Then
                ContactEMail.Text = .Item("EMail")
            Else
                ContactEMail.Text = ""
            End If

            If Not isDbNull(.Item("Phone")) Then
                ContactTel.Text = "(" & Mid(.Item("Phone"), 1, 3) & ") " & Mid(.Item("Phone"), 4, 3) & "-" & _
                                    Mid(.Item("Phone"), 7, 4)
            Else
                ContactTel.Text = ""
            End If
        End With
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowProductInfos: Shows product infos on Qbr                 		                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowProductInfos(ByVal dbConn As OleDbConnection, ByRef rbList As RadioButtonList, ByVal ProductName As String, ByVal strtoSelect As String)
        Dim strReq As String = "Select ProductNo, SecondaryP from Product_Service where PrimaryP='" & ProductName & "'"
        Dim strReqExist As String
        Dim i = 0

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet
        Dim dsTableExist As DataSet
        Dim strTable As String = "Product"
        Dim strTableExist As String = "Exist"
        Dim myRow As DataRow

        Dim strTProducts() As String              'Table of products

        Dim index As Integer = 0

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        rbList.Items.Clear()

        For Each myRow In dsTable.Tables(strTable).Rows
            If Not isDbNull(myRow("SecondaryP")) Then
                rbList.Items.Add(New ListItem(trim(myRow("SecondaryP")), trim(myRow("ProductNo"))))
                rbList.Items(i).Selected = False

                If trim(myRow("ProductNo")) = strtoSelect Then
                    rbList.Items(i).Selected = True
                End If
            Else
                rbList.Items.Clear()
            End If
            i += 1
        Next

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowQbrRelatedInfos: Shows infos on Qbr                                                                          |
    '|          Parameters:                                                                                             |
    '|              intQbr: the number related to the Qbr in the database	                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowQbrRelatedInfos(ByVal dbConn As OleDbConnection, ByVal intQbr As Integer)
        Dim strReq As String = "Select Name, QBRDate, Application, Situation, Solution, Testimonial from Qbr " & _
                                 "where QBRNo=" & intQbr
        Dim strTable As String = "Qbr"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet
        Dim strdate() As String

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        With dsTable.Tables(strTable).Rows.Item(0)
            If Not isDbNull(.Item("Name")) Then
                ProjectName.Text = .Item("Name")
            Else
                ProjectName.Text = ""
            End If

            If Not isDbNull(.Item("QBRDate")) Then
                txtQBRDate.Text = ""
                strdate = split(.Item("QBRDate"), "/")
                If strdate(0) < 10 Then
                    txtQBRDate.Text += "0"
                End If
                txtQBRDate.Text += strdate(0) & "/"
                If strdate(1) < 10 Then
                    txtQBRDate.Text += "0"
                End If
                txtQBRDate.Text += strdate(1) & "/"
                txtQBRDate.Text += strdate(2)
            Else
                txtQBRDate.Text = ""
            End If

            If Not isDbNull(.Item("Application")) Then
                QBRApplication.Text = .Item("Application")
            Else
                QBRApplication.Text = ""
            End If

            If Not IsDBNull(.Item("Situation")) Then
                Situation.Text = .Item("Situation")
            Else
                Situation.Text = ""
            End If

            If Not IsDBNull(.Item("Solution")) Then
                Solution.Text = .Item("Solution")
            Else
                Solution.Text = ""
            End If

            If Not IsDBNull(.Item("Testimonial")) Then
                Testimonial.Text = .Item("Testimonial")
            Else
                Testimonial.Text = ""
            End If
        End With
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowResultInfos: Shows the results on Qbr                                                                        |
    '|          Parameters:                                                                                             |
    '|              intQbr: the number related to the Qbr in the database	                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowResultInfos(ByVal dbConn As OleDbConnection, ByVal intQbr As Integer)
        Dim strReq As String = "Select Summary, Costs, OnceSavings, AnnualSavings, Approved from Result where QBRNo=" & intQbr
        Dim strTable As String = "Qbr"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet
        Dim cout as integer = 1

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        With dsTable.Tables(strTable).Rows.Item(0)
            If Not IsDBNull(.Item("Summary")) Then
                Results.Text = .Item("Summary")
            Else
                Results.Text = ""
            End If

            If Not IsDBNull(.Item("Costs")) Then
                Cost.Text = ShowMoney(.Item("Costs"))
            Else
                Cost.Text = ""
            End If

            If Not IsDBNull(.Item("OnceSavings")) Then
                OnceSavings.Text = ShowMoney(.Item("OnceSavings"))
            Else
                OnceSavings.Text = ""
            End If

            If Not IsDBNull(.Item("AnnualSavings")) Then
                AnnualSavings.Text = ShowMoney(.Item("AnnualSavings"))
            Else
                AnnualSavings.Text = ""
            End If


            If Not IsDBNull(.Item("Costs")) And (Not IsDBNull(.Item("OnceSavings")) Or Not IsDBNull(.Item("AnnualSavings"))) Then
                If Not IsDBNull(.Item("OnceSavings")) And Not IsDBNull(.Item("AnnualSavings")) Then
                    ROI.Text = Round((.Item("OnceSavings") + .Item("AnnualSavings")) / .Item("Costs") * 100, 2)
                ElseIf Not IsDBNull(.Item("OnceSavings")) Then
                    ROI.Text = Round(.Item("OnceSavings") / .Item("Costs") * 100, 2)
                ElseIf Not IsDBNull(.Item("AnnualSavings")) Then
                    ROI.Text = Round(.Item("AnnualSavings") / .Item("Costs") * 100, 2)
                End If
            ElseIf IIf(IsDBNull(.Item("Costs")), 0, .Item("cOSTS")) = 0 Then
                ROI.Text = 100
            Else
                ROI.Text = ""
            End If

            Approved.Checked = .Item("Approved")
        End With
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowInfosQbr: Shows the inFormations we have on the demanded QBR                                                 |
    '|              If mode = "Read"  -> read only                                                                      |
    '|              If mode = "Write" -> possibility to change certain fields                                           |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowInfosQbr(ByVal dbConn As OleDbConnection, ByVal strMode As String)
        Dim strReqQBR As String = "Select QbrNo, EmpNo, Qbr.OSNo, Qbr.UseNo, Qbr.ContactNo, Contact.ClientNo, " & _
                                         "CorporationNo, InputType from Qbr, Contact, Client where Qbr.ContactNo = Contact.ContactNo " & _
                                         "AND Contact.ClientNo = Client.ClientNo AND QbrNo=" & Request.QueryString("Nu")
        Dim strTableQBR As String = "QBR"
        Dim strTableProduct As String = "Product"
        Dim StrWhereCorp As String
        Dim StrWhereEmp As String
        Dim StrWhereOs As String
        Dim StrWhereClient As String
        Dim strWhereUse As String

        Dim cmdTable As New OleDbDataAdapter(strReqQBR, dbConn)
        Dim dsTable As New DataSet

        Dim StrWhereContact As String
        Dim Blank As Boolean                      'Whether or Not we should insert a blank field on the lists
        Dim BlankUse As Boolean = False

        'Execute request
        cmdTable.Fill(dsTable, strTableQBR)

        With dsTable.Tables(strTableQBR).Rows.Item(0)
            If strMode = "R" Then
                Blank = False
                BlankUse = False
                'If we are in consulting mode, only add date from the db QBR in the lists
                StrWhereCorp = " where CorporationNo=" & .Item("CorporationNo")
                StrWhereEmp = " where EmpNo=" & .Item("EmpNo")
                StrWhereClient = " where ClientNo=" & .Item("ClientNo")
                StrWhereContact = " where ContactNo=" & .Item("ContactNo")
                StrWhereOs = " where EmpNo=" & .Item("OsNo")

                If Not IsDBNull(.Item("UseNo")) Then
                    strWhereUse = " where UseNo=" & .Item("UseNo")
                Else
                    strWhereUse = " where 2=1"
                    BlankUse = True
                End If

                Input.Items.Clear()
                If Not IsDBNull(.Item("InputType")) Then
                    Input.Items.Add(.Item("InputType"))
                Else
                    Input.Items.Add("")
                End If
            Else
                StrWhereCorp = ""
                StrWhereEmp = ""
                StrWhereOs = " where title='OS'"
                StrWhereClient = ""
                strWhereUse = ""
                StrWhereContact = " where ClientNo=" & .Item("ClientNo")
                Blank = True
                BlankUse = True

                If Not IsDBNull(.Item("InputType")) Then
                    Input.SelectedValue = .Item("InputType")
                End If
            End If

            CreateListBox(dbConn, "Corporation", "Name", "CorporationNo", "Name", StrWhereCorp, Blank, Corporation)

            CreateListBox(dbConn, "Employee", "Name", "EMPNo", "Name", StrWhereEmp, Blank, Employe)
            If Not IsDBNull(.Item("EmpNo")) Then
                Employe.SelectedValue = .Item("EmpNo")
            End If

            CreateListBox(dbConn, "Employee", "Name", "EMPNo", "Name", StrWhereOs, Blank, OS)
            If Not IsDBNull(.Item("OsNo")) Then
                OS.SelectedValue = .Item("OsNo")
            End If

            CreateListBox(dbConn, "Client", "(Name + ', ' + Cast(ClientNo as Varchar))", "ClientNo", "Name", "", True, CustomerName)
            If Not IsDBNull(.Item("ClientNo")) Then
                CustomerName.SelectedValue = .Item("ClientNo")
                ShowCustomerInfos(dbConn)
            End If

            CreateListBox(dbConn, "Contact", "Name", "ContactNo", "Name", StrWhereContact, Blank, ContactName)
            If Not IsDBNull(.Item("ContactNo")) Then
                ContactName.SelectedValue = .Item("ContactNo")
                ShowContactInfos(dbConn)
            End If

            CreateListBox(dbConn, "UseQbr", "Name", "UseNo", "Name", strWhereUse, BlankUse, Use)
            If Not IsDBNull(.Item("UseNo")) Then
                Use.SelectedValue = .Item("UseNo")
            End If

            ShowQbrRelatedInfos(dbConn, Request.QueryString("Nu"))
            ShowResultInfos(dbConn, Request.QueryString("Nu"))
        End With

        ShowGraphs(dbConn, Qbrcache.Text, strMode)

        If strMode = "R" Then
            makeReadOnly()
        End If

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| makeReadOnly: makes sure every fields are in read only                                                           |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub makeReadOnly()
        SearchControls(Page, True)
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| SearchControls: runs trought all controls to see If we                                                           |
    '|   Notice: using recursion                                                                                        |
    '|   Parameters:    Parent: Contains the list of controls                                                           |
    '|                  Mode: choose we able or enable fields                                                           |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub SearchControls(ByVal Parent As Control, ByVal Mode As Boolean)
        Dim c As Control

        For Each c In Parent.Controls
            If c.GetType().ToString() = "System.Web.UI.WebControls.TextBox" Then
                CType(c, TextBox).ReadOnly = Mode
            End If

            If c.GetType().ToString() = "System.Web.UI.WebControls.CheckBox" Then
                CType(c, CheckBox).Enabled = Not Mode
            End If

            If c.GetType().ToString() = "System.Web.UI.WebControls.CheckBoxList" Then
                CType(c, CheckBoxList).Enabled = Not Mode
            End If

            If c.GetType().ToString() = "System.Web.UI.WebControls.RadioButtonList" Then
                CType(c, RadioButtonList).Enabled = Not Mode
            End If

            If c.Controls.Count > 0 Then
                SearchControls(c, Mode)
            End If
        Next

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| SaveQbr: Insert or update database                                                                               |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub SaveQbr(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dbConn As OleDbConnection
        Dim noQbr As Integer = -1

        EstablishConnection(dbConn)
        dbConn.Open()

        If RequiredFields() Then
            If ValidFields() Then
                If Qbrcache.Text = "-1" Then 'Insert mode
                    noQbr = InsertQbr(dbConn)
                    InsertResults(dbConn, noQbr)
                ElseIf IsNumeric(Qbrcache.Text) Then
                    If Qbrcache.Text > 0 Then ' UpdateMode
                        UpdateQbr(dbConn)
                        UpdateResults(dbConn)
                    End If
                End If

                If Session("products") <> Nothing Then
                    If noQbr = -1 Then
                        InsertProducts(dbConn, Qbrcache.Text)
                    Else
                        InsertProducts(dbConn, noQbr)
                    End If
                    Session("Products") = Nothing
                End If

                If noQbr = -1 Then
                    InsertGraphs(dbConn, Qbrcache.Text)
                Else
                    InsertGraphs(dbConn, noQbr)
                End If
                ResetSessionGraphs()

            End If
        End If

        dbConn.Close()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| RequiredFields: VerIfies If all required fields are entered                                                      |
    '|------------------------------------------------------------------------------------------------------------------|
    Function RequiredFields() As Boolean
        Dim Valid As Boolean = True
        Dim strMsg As String = "You must at least enter the employee name, the OS Name, " & _
                                    "the customer's name and the contact's name"

        If Employe.SelectedItem.Value = "" Or ContactName.SelectedItem.Value = "" Or OS.SelectedItem.Value = "" Then
            Response.Write("<script language=Javascript>alert(""" & strMsg & """);</script>")
            Valid = False
        End If
        RequiredFields = Valid
    End Function

    '|------------------------------------------------------------------------------------------------------------------|
    '| ValidFields: VerIfies If entered fields are in a valid Format                                                    |
    '|       Valids:                                                                                                    |
    '|          Date: Makes sure the date entered is a date Format                                                      |
    '|          Costs & Savings: Makes sure fields are in Number Format                                                 |
    '|------------------------------------------------------------------------------------------------------------------|
    Function ValidFields() As Boolean
        Dim Valid As Boolean = True
        Dim strMsg As String = ""
        Dim strInteger As String = ""

        If txtQbrDate.Text <> "" Then
            If Not IsDate(txtQbrDate.Text) Or Not ValDate.IsValid Then
                strMsg = "You must enter a valid date\n"
                Valid = False
            End If
        End If

        strInteger = Replace(Cost.Text, ",", "")
        strInteger = Replace(Cost.Text, ".", "")
        If strInteger <> "" Then
            If Not IsNumeric(strInteger) Then
                strMsg += "You must enter a valid cost (Number)\n"
                Valid = False
            End If
        End If

        strInteger = Replace(OnceSavings.Text, ",", "")
        strInteger = Replace(OnceSavings.Text, ".", "")
        If strInteger <> "" Then
            If Not IsNumeric(strInteger) Then
                strMsg += "You must enter valid Once Savings (Number)\n"
                Valid = False
            End If
        End If

        strInteger = Replace(AnnualSavings.Text, ",", "")
        strInteger = Replace(AnnualSavings.Text, ".", "")
        If strInteger <> "" Then
            If Not IsNumeric(strInteger) Then
                strMsg += "You must enter valid Annual Savings (Number)\n"
                Valid = False
            End If
        End If

        If Valid = False Then
            Response.Write("<script language=Javascript>alert(""" & strMsg & """);</script>")
        End If

        ValidFields = Valid
    End Function

    '|------------------------------------------------------------------------------------------------------------------|
    '| InsertQbr: Insert QBR in the Database                                                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Function InsertQbr(ByVal dbConn As OleDbConnection) As Integer
        Dim intNoQbr As Integer
        Dim strName As String
        Dim intNoEmp As String
        Dim intNoOS As String
        Dim intNoContact As String
        Dim intNoProduct As String
        Dim strDateQbr As String
        Dim strApplication As String
        Dim strUse As String
        Dim strSituation As String
        Dim strSolution As String
        Dim strTestimonial As String
        Dim strInputType As String

        Dim strReq As String = "SELECT TOP 1 QBRNo FROM QBR ORDER BY QBR.QBRNo DESC"

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet

        Dim cmdConn As New OleDbCommand

        'Execute request
        cmdTable.Fill(dsTable, "QBR")

        'QbrNumber
        If dsTable.Tables("QBR").Rows.Count <> 0 Then
            With dsTable.Tables("QBR").Rows.Item(0)
                intNoQbr = .Item("QbrNo") + 1
            End With
        Else
            intNoQbr = 1
        End If

        Qbrcache.Text = intNoQbr

        strName = SetsNull(ProjectName.Text, True)                          'ProjectName 
        intNoOS = OS.SelectedItem.Value                                     'OS
        intNoEmp = Employe.SelectedItem.Value                               'Employe
        intNoContact = ContactName.SelectedItem.Value                       'ContactName 

        strDateQbr = SetsNull(txtQbrDate.Text, True)                        'DateQbr

        strApplication = SetsNull(QBRApplication.Text, True)                   'Application
        strUse = SetsNull(Use.SelectedItem.Value, False)                    'Use
        strSituation = SetsNull(Situation.Text, True)                       'Situation
        strSolution = SetsNull(Solution.Text, True)                         'Solution
        strTestimonial = SetsNull(Testimonial.Text, True)                   'Testimonial

        strInputType = SetsNull(Input.SelectedItem.Value, True)             'Input Type

        strReq = "Insert into Qbr Values(" & intNoQbr & "," & strName & "," & intNoEmp & "," & intNoOS & "," & intNoContact & _
                                            "," & strDateQbr & "," & strApplication & "," & strUse & _
                                            "," & strSituation & "," & strSolution & "," & strTestimonial & "," & strInputType & ")"
        cmdConn.Connection = dbConn
        cmdConn.CommandText = strReq
        cmdConn.ExecuteNonQuery()

        InsertQbr = intNoQbr
    End Function

    '|------------------------------------------------------------------------------------------------------------------|
    '| UpdateQbr: Update QBR in the Database                                                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub UpdateQbr(ByVal dbConn As OleDbConnection)
        Dim intNoQbr As Integer
        Dim strName As String
        Dim intNoEmp As String
        Dim intNoOs As String
        Dim intNoContact As String
        Dim intNoProduct As String
        Dim strDateQbr As String
        Dim strApplication As String
        Dim strUse As String
        Dim strSituation As String
        Dim strSolution As String
        Dim strTestimonial As String
        Dim strInputType As String

        Dim strReq As String
        Dim cmdConn As New OleDbCommand

        intNoQbr = Qbrcache.Text                                                           'QbrNumber
        strName = "Name = " & SetsNull(ProjectName.Text, True)                          'ProjectName 
        intNoEmp = "EmpNo = " & Employe.SelectedItem.Value                              'Employe
        intNoOs = "OsNo = " & OS.SelectedItem.Value                                     'OS
        intNoContact = "ContactNo = " & SetsNull(ContactName.SelectedItem.Value, False) 'ContactName 

        strDateQbr = "QBRDate=" & SetsNull(txtQbrDate.Text, True)                       'DateQbr

        strApplication = "Application = " & SetsNull(QBRApplication.Text, True)            'Application
        strUse = "UseNo = " & SetsNull(Use.SelectedItem.Value, False)                   'Use
        strSituation = "Situation = " & SetsNull(Situation.Text, True)                  'Situation
        strSolution = "Solution = " & SetsNull(Solution.Text, True)                     'Solution
        strTestimonial = "Testimonial = " & SetsNull(Testimonial.Text, True)            'Testimonial

        strInputType = "InputType = " & SetsNull(Input.SelectedItem.Value, True)

        strReq = "Update Qbr Set " & strName & "," & intNoEmp & "," & intNoOs & "," & intNoContact & _
                "," & strDateQbr & "," & strApplication & "," & strUse & "," & strSituation & "," & strSolution & "," & _
                strTestimonial & "," & strInputType & " where QbrNo=" & intNoQbr

        cmdConn.Connection = dbConn
        cmdConn.CommandText = strReq
        cmdConn.ExecuteNonQuery()

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| InsertResults: Insert Qbr Results in the Database                                                                |
    '|          Parameters:                                                                                             |
    '|                  intNoQbr: Number of the Qbr entered                                                             |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub InsertResults(ByVal dbConn As OleDbConnection, ByVal intNoQbr As Integer)
        Dim strReq As String
        Dim strSummary As String
        Dim strCosts As String
        Dim strOnceSavings As String
        Dim strAnnualSavings As String

        Dim bApproved As Integer

        Dim cmdConn As New OleDbCommand

        strSummary = SetsNull(Results.Text, True)       'Summary 

        strCosts = SetsNull(Cost.Text, False)           'Costs
        If SetsNull(Cost.Text, False) <> "NULL" Then
            strCosts = FormatMoney(strCosts)
        End If

        strOnceSavings = SetsNull(OnceSavings.Text, False)          'OnceSavings 
        If SetsNull(OnceSavings.Text, False) <> "NULL" Then
            strOnceSavings = FormatMoney(strOnceSavings)
        End If

        strAnnualSavings = SetsNull(AnnualSavings.Text, False)      'AnnualSavings 
        If SetsNull(AnnualSavings.Text, False) <> "NULL" Then
            strAnnualSavings = FormatMoney(strAnnualSavings)
        End If


        bApproved = IIf(Approved.Checked, 1, 0)                  'Approved


        strReq = "Insert into Result Values(" & intNoQbr & "," & strSummary & "," & strCosts & "," & strOnceSavings & _
                                            "," & strAnnualSavings & "," & bApproved & ")"
        cmdConn.Connection = dbConn
        cmdConn.CommandText = strReq
        cmdConn.ExecuteNonQuery()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| UpdateResults: Update QBR Results in the Database                                                                |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub UpdateResults(ByVal dbConn As OleDbConnection)
        Dim strReq As String
        Dim intNoQbr As Integer
        Dim strSummary As String
        Dim strCosts As String
        Dim strOnceSavings As String
        Dim strAnnualSavings As String
        Dim bApproved As String

        Dim cmdConn As New OleDbCommand

        intNoQbr = Qbrcache.Text
        strSummary = "Summary=" & SetsNull(Results.Text, True)                  'Summary 

        strCosts = "Costs=" & SetsNull(Cost.Text, False)                        'Costs
        If SetsNull(Cost.Text, False) <> "NULL" Then
            strCosts = "Costs=" & SetsNull(FormatMoney(Cost.Text), False)
        End If

        strOnceSavings = "OnceSavings=" & SetsNull(OnceSavings.Text, False)                 'OnceSavings 
        If SetsNull(OnceSavings.Text, False) <> "NULL" Then
            strOnceSavings = "OnceSavings=" & SetsNull(FormatMoney(OnceSavings.Text), False)
        End If

        strAnnualSavings = "AnnualSavings=" & SetsNull(AnnualSavings.Text, False)           'AnnualSavings 
        If SetsNull(AnnualSavings.Text, False) <> "NULL" Then
            strAnnualSavings = "AnnualSavings=" & SetsNull(FormatMoney(AnnualSavings.Text), False)
        End If

        bApproved = "Approved=" & IIf(Approved.Checked, 1, 0)                               'Approved


        strReq = "Update Result set " & strSummary & "," & strCosts & "," & strOnceSavings & "," & strAnnualSavings & "," & _
                    bApproved & " where QbrNo = " & intNoQbr
        cmdConn.Connection = dbConn
        cmdConn.CommandText = strReq
        cmdConn.ExecuteNonQuery()

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| InsertProducts: Delete all existing rows and adds what is asked on the                                           |
    '|                   product page                                                                                   |
    '|                   Notice: This is called when the Qbr is saved                                                   |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub InsertProducts(ByVal dbConn As OleDbConnection, ByVal intNoQbr As Integer)
        Dim strTProducts() As String                  'Table that contains all the products
        Dim strTModels() As String                    'Table that contains all the Model Numbers
        Dim strReq As String = "Delete from product where QbrNo = " & intNoQbr
        Dim index As Integer

        Dim cmdTable As OleDbDataAdapter
        Dim dsTable As DataSet
        Dim cmdConn As New OleDbCommand

        'If update mode, we delete all rows contained in the bd and Then add the checked ones
        'If insert mode, doesn't matter If we delete, nothing should be in the table         

        cmdConn.Connection = dbConn
        cmdConn.CommandText = strReq
        cmdConn.ExecuteNonQuery()

        If Session("Products") <> Nothing Then
            If Session("Products") <> "" Then
                strTProducts = Split(Session("Products"), ",")
                strTModels = Split(Session("Model"), ",")
                For index = 0 To UBound(strTProducts)
                    If strTProducts(index) <> "" Then
                        'see If it already exists in the table
                        strReq = "SELECT * from product where QbrNo = " & intNoQbr & " And ProductNo=" & strTProducts(index) & _
                                       " And ModelNo='" & strTModels(index) & "'"
                        cmdTable = New OleDbDataAdapter(strReq, dbConn)
                        dsTable = New DataSet
                        cmdTable.Fill(dsTable, "Products")

                        If dsTable.Tables("Products").Rows.Count = 0 Then
                            strReq = "Insert into Product Values(" & intNoQbr & "," & strTProducts(index) & ",'" & strTModels(index) & "')"
                            cmdConn.Connection = dbConn
                            cmdConn.CommandText = strReq
                            cmdConn.ExecuteNonQuery()
                        End If
                    End If
                Next
            End If
        End If

        Session("Products") = Nothing
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| InsertGraphs:  Insert new links on graphs on database                                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub InsertGraphs(ByVal dbConn As OleDbConnection, ByVal intNoQbr As Integer)
        Dim strGraphs() As String = {"linksSituation", "linksSolution", "linksResult"}
        Dim strRemoveDb() As String
        Dim strNomFichier() As String
        Dim index As Integer
        Dim index1 As Integer = 1
        Dim index2 As Integer = 0
        Dim indexFichier As Integer
        Dim intNoGraph As Integer
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable
        Dim cmdConn As New OleDbCommand
        Dim Graph As File

        Dim strPath As String = ""

        Dim strReq As String = "SELECT TOP 1 LinkNo FROM Graphs Where QbrNo = " & intNoQbr & " ORDER BY LinkNo DESC"
        cmdTable = New OleDbDataAdapter(strReq, dbConn)
        cmdTable.Fill(dtTable)

        If dtTable.Rows.Count <> 0 Then
            With dtTable.Rows.Item(0)
                intNoGraph = .Item("LinkNo") + 1
            End With
        Else
            intNoGraph = 1
        End If

        cmdConn.Connection = dbConn

        For index = 0 To UBound(strGraphs)
            index1 = 1

            'delete Graphs from databse and server that the user wished to remove
            strRemoveDb = Split(Session(strGraphs(index) & "Remove"), ",")
            For index2 = 0 To UBound(strRemoveDb)

                If strRemoveDb(index2) <> "" Then
                    strReq = "SELECT Path From Graphs where LinkNo = " & strRemoveDb(index2) & " and QbrNo = " & Qbrcache.Text
                    cmdTable = New OleDbDataAdapter(strReq, dbConn)
                    dtTable = New DataTable
                    cmdTable.Fill(dtTable)

                    'Delete from database
                    strReq = "Delete from Graphs where LinkNo = " & strRemoveDb(index2) & " and QbrNo = " & Qbrcache.Text
                    cmdConn.CommandText = strReq
                    cmdConn.ExecuteNonQuery()

                    'Delete from server
                    With dtTable.Rows.Item(0)
                        If Not IsDBNull(.Item("Path")) Then
                            Graph.Delete(Server.MapPath("./") + .Item("Path"))
                        End If
                    End With
                End If
            Next

            'Add new Graphs

            Do While Session(strGraphs(index) & index1 & "set") <> Nothing
                If Session(strGraphs(index) & index1 & "set") = "true" Then
                    Try
                        Dim Basefile As String = System.IO.Path.GetFileName(Session(strGraphs(index) & index1).PostedFile.FileName)
                        Dim EndFile As String = Basefile
                        Dim paths As String = Server.MapPath("./") + "Qbr" & intNoQbr & "\"
                        Dim destDir As DirectoryInfo = New DirectoryInfo(paths)

                        If Not (destDir.Exists) Then
                            destDir.Create()
                        End If

                        If File.Exists(paths & Basefile) Then
                            indexFichier = 1
                            strNomFichier = Split(Basefile, ".")

                            Do While File.Exists(paths & EndFile)
                                If UBound(strNomFichier) >= 1 Then
                                    EndFile = strNomFichier(0) & "(" & indexFichier & ")." & strNomFichier(1)
                                Else
                                    EndFile = strNomFichier(0) & "(" & indexFichier & ")"
                                End If
                                indexFichier += 1
                            Loop
                        End If

                        Session(strGraphs(index) & index1).PostedFile.saveAs(paths & EndFile)

                        'insert Graph in database
                        strReq = "Insert into Graphs Values(" & intNoQbr & "," & intNoGraph & ",'" & _
                                    "Qbr" & intNoQbr & "\" & EndFile & "'," & index & ")"

                        cmdConn.CommandText = strReq
                        cmdConn.ExecuteNonQuery()

                        Session(strGraphs(index) & index1) = Nothing
                        Session(strGraphs(index) & index1 & "set") = Nothing
                    Catch

                    End Try
                    intNoGraph += 1

                End If
                index1 += 1
            Loop
            Session(strGraphs(index)) = Nothing
        Next

        ShowGraphs(dbConn, Qbrcache.Text, "I")
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| SetsNull: Give either a value or "Null" to a field so                                                            |
    '|                  we can enter it in de database                                                                  |
    '|          Parameters:                                                                                             |
    '|                  isString: identIfies If we must return a string with quotes                                     |
    '|------------------------------------------------------------------------------------------------------------------|
    Function SetsNull(ByVal strValue As String, ByVal isString As Boolean) As String
        strValue = Replace(strValue, "'", "''")

        If strValue <> "" Then
            If isString Then
                SetsNull = "'" & strValue & "'"
            Else
                SetsNull = strValue
            End If
        Else
            SetsNull = "NULL"
        End If
    End Function

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowMoney: puts , where needed                                                                                   |
    '|------------------------------------------------------------------------------------------------------------------|
    Function ShowMoney(ByVal dblMoney As Single) As String
        Dim strDec As String
        Dim strBegin As String
        Dim strEnd As String = ""
        Dim intRest As Integer
        Dim index As Integer = 1

        If dblMoney <> 0 and dblMoney <> 1 Then
            strDec = Round((dblMoney Mod 1), 2)
            If strDec.Length < 4 Then
                strDec += "0"
            End If
            If dblMoney > 0 Then
                strBegin = Floor(dblMoney)
            Else
                strBegin = Ceiling(dblMoney)
            End If

            intRest = strBegin.Length Mod 3

            If intRest = 0 Then intRest = 3


            Do While index < strBegin.Length And intRest <= strBegin.Length
                strEnd &= Mid(strBegin, index, intRest)
                strEnd &= ","
                index += intRest
                intRest = 3
            Loop

            If strDec <> "00" Then
                ShowMoney = Mid(strEnd, 1, strEnd.Length - 1) + "." + Mid(strDec, 3, 2)
            Else
                ShowMoney = Mid(strEnd, 1, strEnd.Length - 1)
            End If
        Else
            ShowMoney = "0,00"
        End If
    End Function

    '|------------------------------------------------------------------------------------------------------------------|
    '| FormatMoney: takes off , where needed                                                                            |
    '|                  so we can enter it in de database                                                               |
    '|------------------------------------------------------------------------------------------------------------------|
    Function FormatMoney(ByVal dblMoney As Single) As String
        Dim strDec As String
        Dim strEnd As String

        If dblMoney Mod 1 <> 0 Then
            strDec = Round((dblMoney Mod 1), 2)
            If strDec.Length < 4 Then
                strDec += "0"
            End If
        Else
            strDec = "0.00"

        End If

        If dblMoney > 0 Then
            strEnd = Floor(dblMoney)
        Else
            strEnd = Ceiling(dblMoney)
        End If

        strEnd = Replace(strEnd, ",", "")
        strEnd = Replace(strEnd, ".", "")

        FormatMoney = strEnd + "." + Mid(strDec, 3, 2)
    End Function

    '|------------------------------------------------------------------------------------------------------------------|
    '| openPrint: opens the print.aspx page                                                                             |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub openPrint(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("Print") = "Print"
        Response.Redirect("print.aspx?Nu=" & Qbrcache.Text)
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| PrintCustomerInfos: Shows customer infos on printable Qbr     		                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub PrintCustomerInfos(ByVal dbConn As OleDbConnection, ByVal intClientNo As Integer)
        Dim strReq As String = "Select Address, City, Province, IndustryNo from Client where ClientNo=" & intClientNo
        Dim strTable As String = "Client"

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        With dsTable.Tables(strTable).Rows.Item(0)
            If Not IsDBNull(.Item("Address")) Then
                lblAddress.Text = .Item("Address")
            Else
                lblAddress.Text = ""
            End If

            lblLocation.Text = ""
            If Not IsDBNull(.Item("City")) Then
                lblLocation.Text += .Item("City")
            End If

            If Not IsDBNull(.Item("Province")) Then
                If lblLocation.Text <> "" Then lblLocation.Text += ", "
                lblLocation.Text += .Item("Province")
            End If

            If Not IsDBNull(.Item("IndustryNo")) Then
                CreateLabel(dbConn, "Industry", "Industry", " where IndustryNo='" & .Item("IndustryNo") & "'", lblIndustry)
            Else
                lblIndustry.Text = ""
            End If
        End With
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| PrintContactInfos: Shows contact infos on printable Qbr         		                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub PrintContactInfos(ByVal dbConn As OleDbConnection, ByVal intContactNo As Integer)
        Dim strReq As String = "Select Title, EMail, Phone from Contact where ContactNo=" & intContactNo
        Dim strTable As String = "Contact"

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        With dsTable.Tables(strTable).Rows.Item(0)
            If Not IsDBNull(.Item("Title")) Then
                lblContactTitle.Text = .Item("Title")
            Else
                lblContactTitle.Text = ""
            End If

            If Not IsDBNull(.Item("EMail")) Then
                lblContactEMail.Text = .Item("EMail")
            Else
                lblContactEMail.Text = ""
            End If

            If Not IsDBNull(.Item("Phone")) Then
                lblContactTel.Text = "(" & Mid(.Item("Phone"), 1, 3) & ") " & Mid(.Item("Phone"), 4, 3) & "-" & _
                                         Mid(.Item("Phone"), 7, 4)
            Else
                lblContactTel.Text = ""
            End If

        End With
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| PrintProductInfos: Shows product infos on printable Qbr         		                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub PrintProductInfos(ByVal dbConn As OleDbConnection, ByVal intQbrNo As Integer)
        Dim strReq As String = "SELECT PrimaryP " & _
                                    "FROM Product_Service, Product WHERE Product_Service.ProductNo = Product.ProductNo " & _
                                    "GROUP BY Product_Service.PrimaryP, Product.QBRNo " & _
                                    "HAVING Product.QBRNo = " & Request.QueryString("Nu") & _
                                    " ORDER BY PrimaryP"
        Dim strTable As String = "Product"
        Dim strTableSub As String = "SubProduct"
        Dim strModelNo As String = ""


        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet
        Dim dsTable1 As New DataSet

        Dim myRow As DataRow
        Dim myRow1 As DataRow

        lblProduct.Text = ""

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        If dsTable.Tables(strTable).Rows.Count <> 0 Then
            lblProduct.Text = "<table>"

            For Each myRow In dsTable.Tables(strTable).Rows
                lblProduct.Text += "<tr valign='top'><td class='SmallText' colspan='3'>" & myRow("PrimaryP") & "</td></tr>"
                strReq = "SELECT SecondaryP, ModelNo " & _
                        "FROM Product_Service, Product WHERE Product_Service.ProductNo = Product.ProductNo AND " & _
                        "Product_Service.PrimaryP='" & myRow("PrimaryP") & "' And QBRNo=" & Request.QueryString("Nu") & _
                        " order by SecondaryP, ModelNo"

                cmdTable = New OleDbDataAdapter(strReq, dbConn)
                dsTable1 = New DataSet
                cmdTable.Fill(dsTable1, strTableSub)

                lblProduct.Text += "<tr height=""5px""><td></td></tr>"
                For Each myRow1 In dsTable1.Tables(strTableSub).Rows
                    lblProduct.Text += "<tr>"
                    lblProduct.Text += "<td width=""20%"" class='MiniText'></td>"  'blank space
                    If myRow1("ModelNo") = "NULL" Then
                        strModelNo = ""
                    Else
                        strModelNo = myRow1("ModelNo")
                    End If
                    lblProduct.Text += "<td class='MiniText'>- &nbsp;" & myRow1("SecondaryP") & "</td>"
                    If myRow1("ModelNo") <> "NULL" Then
                        lblProduct.Text += "<td class='MiniText'>&nbsp;&nbsp;&nbsp;Model # : " & myRow1("ModelNo") & "</td>"
                    End If
                    lblProduct.Text += "</tr>"
                Next
                lblProduct.Text += "<tr height=""20px""><td></td></tr>"
            Next

            lblProduct.Text += "</table>"
        Else
            lblProduct.Text += "&nbsp;"
        End If
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| PrintQbrRelatedInfos: Shows infos on printable                                                                   |
    '|          Parameters:                                                                                             |
    '|              intQbr: the number related to the Qbr in the database	                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub PrintQbrRelatedInfos(ByVal dbConn As OleDbConnection, ByVal intQbr As Integer)
        Dim strReq As String = "Select Name, QBRDate, Application, Situation, Solution, Testimonial, UseNo, InputType from Qbr " & _
                                    "where QBRNo=" & intQbr
        Dim strTable As String = "Qbr"

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        With dsTable.Tables(strTable).Rows.Item(0)
            If Not IsDBNull(.Item("Name")) Then
                lblProjectName.Text = .Item("Name")
            Else
                lblProjectName.Text = ""
            End If

            If Not IsDBNull(.Item("QBRDate")) Then
                lblDateQBR.Text = .Item("QbrDate")
            Else
                lblDateQBR.Text = ""
            End If

            If Not IsDBNull(.Item("Application")) Then
                lblApplication.Text = .Item("Application")
            Else
                lblApplication.Text = ""
            End If

            If Not IsDBNull(.Item("Situation")) Then
                lblSituation.Text = FormatString(.Item("Situation"))
            Else
                lblSituation.Text = ""
            End If

            If Not IsDBNull(.Item("Solution")) Then
                lblSolution.Text = FormatString(.Item("Solution"))
            Else
                lblSolution.Text = ""
            End If

            If Not IsDBNull(.Item("Testimonial")) Then
                lblTestimonial.Text = FormatString(.Item("Testimonial"))
            Else
                lblTestimonial.Text = ""
            End If

            If Not IsDBNull(.Item("UseNo")) Then
                CreateLabel(dbConn, "UseQbr", "Name", " where UseNo=" & .Item("UseNo"), lblUse)
            Else
                lblUse.Text = ""
            End If

            If Not IsDBNull(.Item("InputType")) Then
                lblInputType.Text = FormatString(.Item("InputType"))
            Else
                lblInputType.Text = ""
            End If
        End With
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| PrintResultInfos: Shows the results on printable Qbr                                                             |
    '|          Parameters:                                                                                             |
    '|              intQbr: the number related to the Qbr in the database	                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub PrintResultInfos(ByVal dbConn As OleDbConnection, ByVal intQbr As Integer)
        Dim strReq As String = "Select Summary, Costs, OnceSavings, AnnualSavings, Approved from Result where QBRNo=" & intQbr
        Dim strTable As String = "Qbr"

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        With dsTable.Tables(strTable).Rows.Item(0)
            If Not IsDBNull(.Item("Summary")) Then
                lblResults.Text = FormatString(.Item("Summary"))
            Else
                lblResults.Text = ""
            End If

            If Not IsDBNull(.Item("Costs")) Then
                lblCost.Text = "$ " & ShowMoney(.Item("Costs"))
            Else
                lblCost.Text = "-"
            End If

            If Not IsDBNull(.Item("OnceSavings")) Then
                lblOnceSavings.Text = "$ " & ShowMoney(.Item("OnceSavings"))
            Else
                lblOnceSavings.Text = "-"
            End If

            If Not IsDBNull(.Item("AnnualSavings")) Then
                lblAnnualSavings.Text = "$ " & ShowMoney(.Item("AnnualSavings"))
            Else
                lblAnnualSavings.Text = "-"
            End If

            If Not IsDBNull(.Item("Costs")) And .Item("Costs") <> 0 And (Not IsDBNull(.Item("OnceSavings")) Or Not IsDBNull(.Item("AnnualSavings"))) Then
                If Not IsDBNull(.Item("OnceSavings")) And Not IsDBNull(.Item("AnnualSavings")) Then
                    lblROI.Text = Round((.Item("OnceSavings") + .Item("AnnualSavings")) / .Item("Costs") * 100, 2) & " %"
                ElseIf Not IsDBNull(.Item("OnceSavings")) Then
                    lblROI.Text = Round(.Item("OnceSavings") / .Item("Costs") * 100, 2) & " %"
                ElseIf Not IsDBNull(.Item("AnnualSavings")) Then
                    lblROI.Text = Round(.Item("AnnualSavings") / .Item("Costs") * 100, 2) & " %"
                End If
            ElseIf .Item("Costs") = 0 Then
                lblROI.Text = 100 & " %"
            Else
                lblROI.Text = ""
            End If

            If .Item("Approved") = True Then
                lblApproved.Text = "YES"
            Else
                lblApproved.Text = "NO"
            End If
        End With
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowPrintableQbr: Shows the inFormations we have on the                                                          |
    '|                          demanded printable QBR                                                                  |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowPrintableQbr(ByVal dbConn As OleDbConnection)
        Dim strReq As String = "Select QbrNo, EmpNo, Qbr.OsNo, Qbr.ContactNo, Contact.ClientNo, CorporationNo " & _
                                            "from Qbr, Contact, Client where Qbr.ContactNo = Contact.ContactNo " & _
                                            "AND Contact.ClientNo = Client.ClientNo AND QbrNo=" & Request.QueryString("Nu")
        Dim strTable As String = "QBR"

        Dim StrWhereEmp As String
        Dim StrWhereClient As String
        Dim StrWhereContact As String
        Dim StrWhereCorp As String
        Dim StrWhereOs As String

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dsTable As New DataSet

        'Execute request
        cmdTable.Fill(dsTable, strTable)

        With dsTable.Tables(strTable).Rows.Item(0)
            StrWhereCorp = " where CorporationNo=" & .Item("CorporationNo")
            StrWhereEmp = " where EmpNo=" & .Item("EmpNo")
            StrWhereOs = " where EmpNo=" & .Item("OsNo")
            StrWhereClient = " where ClientNo=" & .Item("ClientNo")
            StrWhereContact = " where ContactNo=" & .Item("ContactNo")

            If Not IsDBNull(.Item("ContactNo")) Then
                CreateLabel(dbConn, "Corporation", "Name", StrWhereCorp, lblCorporation)
            Else
                lblCorporation.Text = ""
            End If

            If Not IsDBNull(.Item("EmpNo")) Then
                CreateLabel(dbConn, "Employee", "Name", StrWhereEmp, lblEmploye)
            Else
                lblEmploye.Text = ""
            End If

            If Not IsDBNull(.Item("OsNo")) Then
                CreateLabel(dbConn, "Employee", "Name", StrWhereOs, lblOS)
            Else
                lblOS.Text = ""
            End If

            If Not IsDBNull(.Item("ClientNo")) Then
                CreateLabel(dbConn, "Client", "Name", StrWhereClient, lblCustomerName)
                PrintCustomerInfos(dbConn, .Item("ClientNo"))
            Else
                lblCustomerName.Text = ""
            End If

            If Not IsDBNull(.Item("ContactNo")) Then
                CreateLabel(dbConn, "Contact", "Name", StrWhereContact, lblContactName)
                PrintContactInfos(dbConn, .Item("ContactNo"))
            Else
                lblContactName.Text = ""
            End If

            PrintProductInfos(dbConn, Request.QueryString("Nu"))

            PrintQbrRelatedInfos(dbConn, Request.QueryString("Nu"))
            PrintResultInfos(dbConn, Request.QueryString("Nu"))

            ShowGraphs(dbConn, Request.QueryString("Nu"), "R")
        End With
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| FormatString: Adds <br /> after x characters                                                                     |
    '|------------------------------------------------------------------------------------------------------------------|
    Function FormatString(ByVal myString As String) As String
        Dim x As Integer = 80
        Dim pos As Integer = 0
        Dim index As Integer = 0

        Dim strTemp As String
        Dim strTab() As String = Split(myString, Chr(13))

        myString = ""

        For index = 0 To UBound(strTab)
            strTemp = strTab(index)
            Do While strTemp.Length > x
                pos = InStr(Mid(strTemp, x, strTemp.Length), " ") + x - 1
                If pos <> Nothing Then
                    myString += Mid(strTemp, 1, pos) & "<br />"
                    strTemp = Mid(strTemp, pos + 1, strTemp.Length)
                Else
                    myString += strTemp
                    strTemp = ""
                End If
            Loop
            If strTemp <> "" Then myString += strTemp & "<br />"
        Next

        FormatString = myString
    End Function

    '|------------------------------------------------------------------------------------------------------------------|
    '| AddObjects: Dynamically adds field to enter products on Products.aspx                                            |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub AddObjects(ByVal dbConn As OleDbConnection)
        Dim i As Integer = 1

        Dim strReq As String
        Dim strTable As String = "Products"
        Dim strNoRows As String

        Dim cmdTable As OleDbDataAdapter
        Dim dsTable As DataSet
        Dim myRow As DataRow

        Dim strTProducts() As String                  'Table that contains all the products
        Dim strTModels() As String                    'Table that contains all the Model Numbers

        Dim index As Integer = 0

        If Request.QueryString("Mode") = "1" Or Request.QueryString("Mode") = "2" Or Request.QueryString("Mode") = "3" Then
            'If already open the web page, show the modIfications he may had made 
            If Session("Products") <> Nothing And Session("Products") <> "" And Session("Products") <> "Empty" Then
                strTProducts = Split(Session("Products"), ",")
                strTModels = Split(Session("Model"), ",")
                For index = 0 To UBound(strTProducts)
                    If strTProducts(index) <> "" Then
                        AddControls(dbConn, Request.QueryString("Mode"), i, strTProducts(index), strTModels(index))
                        i += 1
                    End If
                Next
            Else
                'Show infos from the database
                strReq = "SELECT ProductNo, ModelNo FROM product WHERE Product.QBRNo = " & Request.QueryString("Nu") & _
                           " order by ProductNo, ModelNo"

                cmdTable = New OleDbDataAdapter(strReq, dbConn)
                dsTable = New DataSet
                cmdTable.Fill(dsTable, strTable)
                For Each myRow In dsTable.Tables(strTable).Rows
                    AddControls(dbConn, Request.QueryString("Mode"), i, myRow("ProductNo"), myRow("ModelNo"))
                    i += 1
                Next
                If i > Session("NbProducts") Then
                    Session("NbProducts") = i - 1
                End If
            End If

            If Request.QueryString("Mode") = "3" And i = 1 Then
                strNoRows = "<span class='DarkBlue'>No products are entered for this QBR<br /><br /></span>"
                ph1.Controls.Add(New LiteralControl(strNoRows))
            End If

            If Request.QueryString("Mode") = "3" Then
                makeReadOnly()
            ElseIf Request.QueryString("Mode") = "1" Or Request.QueryString("Mode") = "2" Then
                For i = i To Session("NbProducts")
                    AddControls(dbConn, Request.QueryString("Mode"), i, "", "")
                Next
            End If
        End If
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| AddControls: Sets the properties of the dropdownlist and the RadioButtonList and adds them to the placeHolder       |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub AddControls(ByVal dbConn As OleDbConnection, ByVal strMode As String, ByVal i As Integer, ByVal strToSelect As String, ByVal strModel As String)
        Dim strWhere As String = ""
        Dim Blank As Boolean = True
        Dim ddlProduct As New DropDownList
        Dim rbProduct As New RadioButtonList
        Dim txtModel As New TextBox
        Dim strPrimary As String = ""

        Dim strReq = "select PrimaryP from Product_service where ProductNo =" & strToSelect
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable

        If strMode = "3" Then
            strWhere = " where ProductNo=" & strToSelect
            Blank = False
        End If

        ddlProduct.AutoPostBack = "True"

        ddlProduct.CssClass = "BGBabyBlue DarkBlue MiniText"
        rbProduct.CssClass = "BGBabyBlue DarkBlue MiniText"
        txtModel.CssClass = "BGBabyBlue DarkBlue MiniText ModelWidth"

        ddlProduct.ID = "ddlProduct" & i
        txtModel.ID = "txtModel" & i
        txtModel.MaxLength = "10"
        rbProduct.ID = "rbProduct" & i
        rbProduct.RepeatDirection = RepeatDirection.Horizontal

        AddHandler ddlProduct.SelectedIndexChanged, AddressOf ProductChange
        CreateListBox(dbConn, "Product_Service", "PrimaryP", "PrimaryP", "PrimaryP", strWhere, Blank, ddlProduct)

        If strToSelect <> "" Then
            cmdTable = New OleDbDataAdapter(strReq, dbConn)

            cmdTable.Fill(dtTable)

            With dtTable.Rows.Item(0)
                strPrimary = .Item("PrimaryP")
            End With

            ddlProduct.SelectedValue = strPrimary
        End If

        ph1.Controls.Add(New LiteralControl("<span class='smalltext darkBlue'>Product " & i & ": &nbsp;&nbsp;</span>"))
        ph1.Controls.Add(ddlProduct)
        ph1.Controls.Add(rbProduct)
        ph1.Controls.Add(New LiteralControl("<span class='smalltext darkBlue'>Model #: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>"))
        ph1.Controls.Add(txtModel)
        ph1.Controls.Add(New LiteralControl("<br /><br />"))


        ViewSubProducts(dbConn, ddlProduct, strToSelect)
        txtModel.Text = ViewModelNo(strModel)


    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ViewModelNo: Returns the Model #                                                                                 |
    '|------------------------------------------------------------------------------------------------------------------|
    Function ViewModelNo(ByVal strModel As String) As String

        If strModel = "NULL" Then
            ViewModelNo = ""
        Else
            ViewModelNo = strModel
        End If

    End Function
    '|------------------------------------------------------------------------------------------------------------------|
    '| ProductChange: Calls ViewSubProducts                                                                             |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ProductChange(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim dbConn As OleDbConnection

        EstablishConnection(dbConn)
        dbConn.Open()

        ViewSubProducts(dbConn, sender, "-1")

        dbConn.Close()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ViewSubProducts: Shows RadioButtonLists depending on the selected product                                           |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ViewSubProducts(ByVal dbConn As OleDbConnection, ByVal ddList As DropDownList, ByVal strToSelect As String)
        Dim ctl As Control
        Dim ct2 As Control
        Dim ct3 As Control

        Dim strNomDDL As String = "ddlProduct"
        Dim strNomCBL As String = "rbProduct"

        'Determine the index of the control we selected 
        Dim intIndexDDL As Integer = Mid(ddList.ID, strNomDDL.Length + 1, ddList.ID.Length - strNomDDL.Length)
        Dim intIndexCBL As Integer = 0

        'Go through all controls in the page to find the RadioButtonLists
        For Each ctl In Page.Controls
            If ctl.GetType Is GetType(Web.UI.HtmlControls.HtmlForm) Then
                For Each ct2 In ctl.Controls
                    If ct2.GetType Is GetType(System.Web.UI.WebControls.PlaceHolder) Then
                        For Each ct3 In ct2.Controls
                            If ct3.GetType Is GetType(System.Web.UI.WebControls.RadioButtonList) Then
                                'If it has the same index As the DropDownList
                                intIndexCBL = Mid(ct3.ID, strNomCBL.Length + 1, ct3.ID.Length - strNomCBL.Length)
                                If intIndexCBL = intIndexDDL Then
                                    ShowProductInfos(dbConn, CType(ct3, RadioButtonList), ddList.SelectedItem.Text, strToSelect)
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ConfirmProducts: Confirms the entries For the products                                                           |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ConfirmProducts(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctl As Control
        Dim ct2 As Control
        Dim ct3 As Control

        Dim i As Integer

        Dim selected As Boolean = False

        Dim strProducts As String = ""
        Session("Model") = ""

        'Go through all controls in the page to find the checkboxes
        For Each ctl In Page.Controls
            If ctl.GetType Is GetType(Web.UI.HtmlControls.HtmlForm) Then
                For Each ct2 In ctl.Controls
                    If ct2.GetType Is GetType(System.Web.UI.WebControls.PlaceHolder) Then
                        For Each ct3 In ct2.Controls
                            If ct3.GetType Is GetType(System.Web.UI.WebControls.RadioButtonList) Then
                                selected = False
                                For i = 0 To CType(ct3, RadioButtonList).Items.Count - 1
                                    If CType(ct3, RadioButtonList).Items(i).Selected Then
                                        strProducts += CType(ct3, RadioButtonList).Items(i).Value & ","
                                        selected = True
                                    End If
                                Next
                            End If
                            If ct3.GetType Is GetType(System.Web.UI.WebControls.TextBox) And selected Then
                                If CType(ct3, TextBox).Text = "" Then
                                    Session("Model") += "NULL,"
                                Else
                                    Session("Model") += CType(ct3, TextBox).Text & ","
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Next

        'affects a Session variable which we will use when inserting, modIfying the qbr
        If strProducts <> "" Then
            Session("Products") = strProducts
        Else
            Session("Products") = "Empty"
        End If

        Response.Write("<script language=Javascript>self.close();</script>")
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| openProducts: opens the products window                                                                          |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub openProducts(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim strComm As String = "<script language=Javascript>window.open('products.aspx?Nu=" & Qbrcache.Text & "&Mode=" & _
                                  Session("type") & "','new','menubar=no,scrollbars=yes,height=450,resizable=yes,width=900, " & _
                                  "left=100, top=300');</script>"
        Response.Write(strComm)
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| openNewContact: opens the Contact window                                                                         |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub openNewContact(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim strComm As String
        If CustomerName.SelectedItem.Value <> "" Then
            strComm = "<script language=Javascript>window.open('contact.aspx?Cli=" & CustomerName.SelectedItem.Value & "','new','menubar=no, " & _
                                "scrollbars=yes,height=350,resizable=no,width=680, " & _
                                "left=200, top=300');</script>"
        Else
            strComm = "<script language=Javascript>alert(""You must select a Customer first"");</script>"
        End If
        Response.Write(strComm)
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| EditContact: opens the Contact window for modifications                                                          |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub EditContact(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim strComm As String
        If ContactName.SelectedItem.Value <> "" Then
            strComm = "<script language=Javascript>window.open('contact.aspx?Cli=" & CustomerName.SelectedItem.Value & _
                            "&Con=" & ContactName.SelectedItem.Value & _
                            "','new','menubar=no, " & _
                            "scrollbars=yes,height=350,resizable=no,width=680, " & _
                            "left=200, top=300');</script>"
        Else
            strComm = "<script language=Javascript>alert(""You must select a Contact first"");</script>"
        End If
        Response.Write(strComm)
    End Sub

    '|-----------------------------------------------------------------------------------|
    '| ValidEMail : s'assure que le courriel a un format nom@serveur.domaine             |
    '|-----------------------------------------------------------------------------------|
    Protected Sub ValidEMail(ByVal obj As Object, ByVal Arguments As ServerValidateEventArgs)
        Try
            Dim pos As Integer
            pos = InStr(1, Arguments.Value, "@", CompareMethod.Text)
            Arguments.IsValid = InStr(pos, Arguments.Value, ".", CompareMethod.Text) <> 0 And _
                 Arguments.Value.Length > InStr(pos, Arguments.Value, ".", CompareMethod.Text)
        Catch Ex As Exception
            Arguments.IsValid = False
        End Try
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| PutsCaracters: returns the string with ' For the query                                                           |
    '|                replace @*@ by '                                                                                  |
    '|                (because ' doesn't work when sent in a link <a href...                                            | 
    '|------------------------------------------------------------------------------------------------------------------|
    Function PutsCaracters(ByVal strWhere As String) As String

        PutsCaracters = Replace(strWhere, "@*@", "'")

    End Function

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowCalendar: shows or hides the calendars                                                                       |
    '| Notice: image must have format "Img" & var1                                                                      |
    '|         calendar must have format "Cal" & var1                                                                   |
    '|         image var1 and calendar var1 must be the same                                                            |
    '|         - uses recursivity                                                                                       |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowCalendar(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Dim c As Control = New Control
        Dim Parent As Control = Page

Boucle:
        For Each c In Parent.Controls
            If c.GetType().ToString() = "System.Web.UI.WebControls.Calendar" And _
                c.ID = "Cal" & Mid(sender.id, 4, sender.id.length - 3) Then
                CType(c, Calendar).Visible = Not CType(c, Calendar).Visible
            End If
            If c.Controls.Count > 0 Then
                Parent = c
                c = New Control
                GoTo Boucle
            End If
        Next
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowDate: show selected date in the appropriate textfield                                                        |
    '| Notice: - calendar must have format "Cal" & var1                                                                 |
    '|           textBox must have format "txt" & var1 & "Date"                                                         |
    '|           calendar var1 and textbox var1 must be the same                                                        |
    '|         - uses recursivity                                                                                       |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowDate(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim c As Control = New Control
        Dim Parent As Control = Page

Boucle:
        For Each c In Parent.Controls
            If c.GetType().ToString() = "System.Web.UI.WebControls.TextBox" And _
                c.ID = "txt" & Mid(sender.id, 4, sender.id.length - 3) & "Date" Then
                CType(c, TextBox).Text = sender.SelectedDate
                sender.visible = False
            End If
            If c.Controls.Count > 0 Then
                Parent = c
                c = New Control
                GoTo Boucle
            End If
        Next


    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ValidDate: makes sure the dates entered are in a date Format                                                     |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ValidDate(ByVal source As Object, ByVal args As ServerValidateEventArgs)
        If Page.IsValid Then
            args.IsValid = args.Value = "" Or IsDate(args.Value)
        Else
            args.IsValid = True
        End If
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowContacts: Show the list of contacts depending on the customer selected                                       |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowContacts(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim strWhere As String
        Dim dbConn As OleDbConnection

        EstablishConnection(dbConn)
        dbConn.Open()

        If sender.SelectedItem.Value <> "" Then
            strWhere = " Where Contact.ClientNo = " & sender.SelectedItem.Value
            ShowCustomerInfos(dbConn)
            ContactTitle.Text = ""
            ContactEMail.Text = ""
            ContactTel.Text = ""
        Else
            strWhere = " Where 2=1"
            Address.Text = ""
            Location.Text = ""
            Industry.Text = ""
        End If

        ContactName.Items.Clear()
        CreateListBox(dbConn, "Contact", "Name", "ContactNo", "Name", strWhere, True, ContactName)

        dbConn.Close()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowCustomers: Show the list of customers depending on the corporation                                           | 
    '|                      selected                                                                                    |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowCustomers(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim strWhere As String
        Dim dbConn As OleDbConnection

        EstablishConnection(dbConn)
        dbConn.Open()

        If sender.SelectedItem.Value <> "" Then
            strWhere = " Where Client.CorporationNo = " & sender.SelectedItem.Value
            Address.Text = ""
            Location.Text = ""
            Industry.Text = ""
        Else
            strWhere = " Where 2=1"
        End If

        ContactName.Items.Clear()
        CustomerName.Items.Clear()

        CreateListBox(dbConn, "Client", "ClientNo & "", "" & Name", "ClientNo", "Name", strWhere, True, CustomerName)
        CreateListBox(dbConn, "Contact", "Name", "ContactNo", "Name", " where 2=1", True, ContactName)

        dbConn.Close()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| CallsShowContactInfos: calls ShowContactInfos Function when we select                                            | 
    '|                        a New contact in the list                                                                 |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub CallsShowContactInfos(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dbConn As OleDbConnection

        EstablishConnection(dbConn)

        dbConn.Open()

        If sender.SelectedItem.Value <> "" Then
            ShowContactInfos(dbConn)
        Else
            ContactTitle.Text = ""
            ContactEMail.Text = ""
            ContactTel.Text = ""
        End If

        dbConn.Close()

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| AddContact: Adds contact in the database                                                                         | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub AddContact(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dbConn As OleDbConnection

        Dim intNoContact As Integer

        Dim strNoCustomer As String = Request.QueryString("Cli")
        Dim strName As String
        Dim strTitle As String
        Dim strEmail As String
        Dim strPhone As String

        Dim cmdTable As OleDbDataAdapter
        Dim dsTable As New DataSet
        Dim cmdConn As New OleDbCommand

        'verifies the existence of the customer passed in the page
        Dim strReq As String = "Select * from Client where ClientNo=" & strNoCustomer

        EstablishConnection(dbConn)
        dbConn.Open()

        If Page.IsValid Then
            cmdTable = New OleDbDataAdapter(strReq, dbConn)

            cmdTable.Fill(dsTable, "Customer")

            If dsTable.Tables("Customer").Rows.Count > 0 Then 'if the custumer exists

                'gets the new contact's number
                strReq = "SELECT TOP 1 ContactNo FROM Contact ORDER BY ContactNo DESC"
                cmdTable = New OleDbDataAdapter(strReq, dbConn)
                cmdTable.Fill(dsTable, "Contact")

                If dsTable.Tables("Contact").Rows.Count <> 0 Then
                    With dsTable.Tables("Contact").Rows.Item(0)
                        intNoContact = .Item("ContactNo") + 1
                    End With
                Else
                    intNoContact = 1
                End If

                'creates the fields we enter in the db
                strName = SetsNull(txtcontactName.Text, True)                  'ContactName 
                strTitle = SetsNull(ContactTitle.Text, True)                   'contactTitle
                strEmail = SetsNull(Replace(ContactEMail.Text, " ", ""), True) 'contactEMail

                strPhone = Replace(ContactTel.Text, " ", "")
                strPhone = Replace(strPhone, "(", "")
                strPhone = Replace(strPhone, ")", "")
                strPhone = Replace(strPhone, "-", "")
                strPhone = SetsNull(Trim(strPhone), False) 'contactPhone        

                strReq = "Insert into Contact Values(" & intNoContact & "," & strNoCustomer & "," & strName & "," & _
                                                        strTitle & "," & strPhone & "," & strEmail & ")"
                cmdConn.Connection = dbConn
                cmdConn.CommandText = strReq
                cmdConn.ExecuteNonQuery()
            End If
            Session("postCustomer") = intNoContact
            Response.Write("<script language=Javascript>opener.document.forms[0].submit();self.close();</script>")
        End If

        dbConn.Close()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| UpdateContact: Updates contact in the database                                                                   | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub UpdateContact(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dbConn As OleDbConnection
        Dim strReq As String

        Dim intNoContact As String = Request.QueryString("Con")
        Dim strName As String
        Dim strTitle As String
        Dim strEmail As String
        Dim strPhone As String

        Dim cmdConn As New OleDbCommand

        EstablishConnection(dbConn)
        dbConn.Open()

        If Page.IsValid Then
            'creates the fields we enter in the db
            strName = "Name = " & SetsNull(txtcontactName.Text, True)                   'ContactName 
            strTitle = "Title = " & SetsNull(ContactTitle.Text, True)                   'contactTitle
            strEmail = "Email = " & SetsNull(Replace(ContactEMail.Text, " ", ""), True) 'contactEMail

            strPhone = Replace(ContactTel.Text, " ", "")
            strPhone = Replace(strPhone, "(", "")
            strPhone = Replace(strPhone, ")", "")
            strPhone = Replace(strPhone, "-", "")
            strPhone = "Phone = " & SetsNull(Trim(strPhone), False)                     'contactPhone        

            strReq = "Update Contact Set " & strName & "," & strTitle & "," & strPhone & "," & strEmail & _
                        " where ContactNo = " & intNoContact

            cmdConn.Connection = dbConn
            cmdConn.CommandText = strReq
            cmdConn.ExecuteNonQuery()
            Session("postCustomer") = Request.QueryString("Con")

            Response.Write("<script language=Javascript>opener.document.forms[0].submit();self.close();</script>")
        End If

        dbConn.Close()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| CreateReport: Creates the htm line report                                                                        |
    '|------------------------------------------------------------------------------------------------------------------|
    Sub CreateReport(ByVal dbConn As OleDbConnection)
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable
        Dim i As Integer
        Dim objbc As BoundColumn

        Dim strReq = "SELECT Convert(varchar,QBR.QbrDate,101) AS [Date], QBR.Name AS [Project Name], " & _
                        "Corporation.Name AS Enterprise, Client.Name AS Customer, Client.Address, " & _
                        "Client.City, Client.Province, Industry.Industry, Contact.Name AS [Contact Name], " & _
                        "Contact.Title, Contact.Phone, Contact.EMail, Employee AS [Employe Name], " & _
                        "OS.OS AS [OS Name], QBR.Application, UseQBR.Name AS [Use], Qbr.InputType as [Input Type], Product_Service.PrimaryP AS [Primary], " & _
                        "Product_Service.SecondaryP AS [Sub Product], QBR.Situation, QBR.Solution, " & _
                        "Result.Summary AS [Result Summary], '$' + CAST(ISNULL(Result.Costs,0) AS varchar) AS [Costs], " & _
                        "'$' + CAST(ISNULL(Result.OnceSavings,0) AS varchar) AS [Once Savings], " & _
                        "'$' + CAST(ISNULL(Result.AnnualSavings,0) AS varchar) AS [Annual Savings], " & _
                        "CAST(CASE WHEN Result.Costs = 0 THEN 0 ELSE Round((ISNULL(Result.OnceSavings,0) + ISNULL(Result.AnnualSavings,0)) / Result.Costs * 100, 2) END AS varchar) + '%' AS [ROI], " & _
                        "Approved = Case Result.Approved WHEN '1' THEN 'YES' WHEN '0' THEN 'NO' END, QBR.Testimonial " & _
                        "FROM Industry RIGHT JOIN (UseQbr RIGHT JOIN (((Corporation INNER JOIN " & _
                        "((Client INNER JOIN Contact ON Client.ClientNo = Contact.ClientNo) INNER JOIN " & _
                        "((Select EmpNo, Name AS OS from Employee) AS [OS] INNER JOIN " & _
                        "((Select EmpNo, Name AS Employee  from Employee) AS [Employees] " & _
                        "INNER JOIN QBR ON Employees.EmpNo = QBR.EmpNo) ON OS.EmpNo = QBR.OSNo) " & _
                        "ON Contact.ContactNo = QBR.ContactNo) ON Corporation.CorporationNo = " & _
                        "Client.CorporationNo) LEFT JOIN (Product_Service RIGHT JOIN Product ON " & _
                        "Product_Service.ProductNo = Product.ProductNo) ON QBR.QBRNo = Product.QBRNo) " & _
                        "INNER JOIN Result ON QBR.QBRNo = Result.QBRNo) ON UseQbr.UseNo = QBR.UseNo) ON " & _
                        "Industry.industryNo = Client.IndustryNo " & _
                        "ORDER BY QBr.QBRNo, QBR.QbrDate DESC"

        cmdTable = New OleDbDataAdapter(strReq, dbConn)
        cmdTable.Fill(dtTable)

        For i = 0 To dtTable.Columns.Count - 1
            objbc = New BoundColumn

            objbc.DataField = dtTable.Columns(i).ColumnName
            objbc.HeaderText = dtTable.Columns(i).ColumnName

            dgQbr.Columns.Add(objbc)
            dgQbr.DataSource = dtTable
            dgQbr.DataBind()
        Next

        AddButtons(dtTable)

        HideColumns()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| AddButtons: Adds buttons to show / hide columns                                                                  | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub AddButtons(ByVal dtTable As DataTable)
        Dim btnShow As New Button
        Dim i As Integer

        ph1.Controls.Add(New LiteralControl("<table><tr>"))

        For i = 0 To dtTable.Columns.Count - 1
            btnShow = New Button
            btnShow.ID = "btn" & i
            btnShow.Text = "hide"

            btnShow.CssClass = "smallbutton ButtonWidth"

            AddHandler btnShow.Click, AddressOf hideShowColumn

            ph1.Controls.Add(New LiteralControl("<td>"))
            ph1.Controls.Add(New LiteralControl("<span class='minitext darkBlue'>" & dtTable.Columns(i).ColumnName & ": </span>"))
            ph1.Controls.Add(New LiteralControl("</td><td>"))
            ph1.Controls.Add(btnShow)
            ph1.Controls.Add(New LiteralControl("&nbsp;&nbsp;"))
            ph1.Controls.Add(New LiteralControl("</td>"))

            'switch line each 8 buttons
            If (i + 1) Mod 7 = 0 And i Then
                ph1.Controls.Add(New LiteralControl("</tr><tr>"))
            End If
        Next
        ph1.Controls.Add(New LiteralControl("</tr></table>"))
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| hideShowColumn: hide or show columns when we click on the button associated to the column                        | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub hideShowColumn(ByVal sender As Object, ByVal e As System.EventArgs)
        If sender.text = "hide" Then
            sender.text = "show"
            dgQbr.Columns(Mid(sender.id, 4, sender.id.length - 3)).Visible = False
        Else
            sender.text = "hide"
            dgQbr.Columns(Mid(sender.id, 4, sender.id.length - 3)).Visible = True
        End If
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| HideColumns: hide all the columns that we have already choose to hide on page_load                               | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub HideColumns()
        Dim c As Control = New Control
        Dim Parent As Control = Page
        Dim btn As Button

Boucle:
        For Each c In Parent.Controls
            If c.GetType().ToString() = "System.Web.UI.WebControls.Button" Then
                btn = CType(c, Button)
                If btn.Text = "show" Then
                    dgQbr.Columns(Mid(c.ID, 4, btn.ID.Length - 3)).Visible = False
                End If
            End If
            If c.Controls.Count > 0 Then
                Parent = c
                c = New Control
                GoTo Boucle
            End If
        Next
    End Sub


    '|------------------------------------------------------------------------------------------------------------------|
    '| ExportToExcel: Exports the datagrid to a xls Excel file                                                          | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ExportToExcel(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.ContentType = "application/vnd.ms-excel"
        ' Remove the charset from the Content-Type header.
        Response.Charset = ""
        ' Turn off the view state.
        Me.EnableViewState = False
        Dim tw As New System.IO.StringWriter
        Dim hw As New HtmlTextWriter(tw)
        ' Get the HTML for the control.
        dgQbr.RenderControl(hw)
        ' Write the HTML back to the browser.
        Response.Write(tw.ToString())
        ' End the response.
        Response.End()

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ClearSearchSessionVariables: Sets the session variables from the search to nothing                               | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ClearSearchSessionVariables()
        Session("Corporation") = Nothing
        Session("Customer") = Nothing
        Session("Industry") = Nothing
        Session("Application") = Nothing
        Session("Employee") = Nothing
        Session("StartDate") = Nothing
        Session("EndDate") = Nothing
        Session("Product") = Nothing
        Session("SubProduct") = Nothing
        Session("ModelNo") = Nothing
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| AddFile: Add Files to a Session Variable                                                                         | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub AddFile(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim DbConn As OleDbConnection
        EstablishConnection(DbConn)
        DbConn.Open()

        ErrorLinkSituation.Text = ""
        ErrorLinkSolution.Text = ""
        ErrorLinkResult.Text = ""

        If (sender.id.ToString()) = "LinksSituation" Then
            AddToLinks("LinksSituation", MyFileSituation, ErrorLinkSituation)
        ElseIf (sender.id.ToString()) = "LinksSolution" Then
            AddToLinks("LinksSolution", MyFileSolution, ErrorLinkSolution)
        ElseIf (sender.id.ToString()) = "LinksResult" Then
            AddToLinks("LinksResult", MyFileResult, ErrorLinkResult)
        End If

        ShowGraphs(DbConn, Qbrcache.Text, "I")

        DbConn.Close()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| FileExists: Verifies if the user already entered this link                                                       | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub AddToLinks(ByVal strType As String, ByVal MyFile As HtmlInputFile, ByRef lblErreur As Label)
        Dim i As Integer = 1

        If MyFile.PostedFile.FileName <> "" Then
            If Not FileExists(strType, MyFile.PostedFile.FileName) Then
                Do While Session(strType & i & "set") <> Nothing
                    i += 1
                Loop

                Session(strType & i) = MyFile
                Session(strType & i & "set") = "true"

            Else
                lblErreur.Text = "Vous avez déjà entré ce fichier."
            End If
        End If

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| FileExists: Verifies if the user already entered this link                                                       | 
    '|------------------------------------------------------------------------------------------------------------------|
    Function FileExists(ByVal Links As String, ByVal LinkEntered As String)
        Dim index As Integer = 1
        Dim Exists As Boolean = False

        Do While Session(Links & index & "set") <> Nothing
            If Session(Links & index & "set") = True Then
                If Session(Links & index).PostedFile.FileName = LinkEntered Then
                    Exists = True
                End If
            End If
            index += 1
        Loop

        FileExists = Exists

    End Function

    '|------------------------------------------------------------------------------------------------------------------|
    '| ShowGraphs: Verifies if the user already entered this link                                                       | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ShowGraphs(ByVal dbConn As OleDbConnection, ByVal intNoQbr As Integer, ByVal strMode As String)
        lblLinksSituation.Text = ""
        lblLinksSolution.Text = ""
        lblLinksResult.Text = ""

        'If not new Qbr, show links from the database
        If intNoQbr <> -1 Then
            PutsLinksFromDBLabel(dbConn, intNoQbr, "LinksSituation", lblLinksSituation, 0, strMode)
            PutsLinksFromDBLabel(dbConn, intNoQbr, "LinksSolution", lblLinksSolution, 1, strMode)
            PutsLinksFromDBLabel(dbConn, intNoQbr, "LinksResult", lblLinksResult, 2, strMode)
        End If

        'Show Links from Session variables
        If strMode <> "R" Then
            PutLinksLabel("LinksSituation", lblLinksSituation, strMode)
            PutLinksLabel("LinksSolution", lblLinksSolution, strMode)
            PutLinksLabel("LinksResult", lblLinksResult, strMode)
        End If
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| PutLinksLabel: Puts graph links in the correct Label                                                             | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub PutLinksLabel(ByVal strLinks As String, ByRef lblLinks As Label, ByVal strMode As String)
        Dim index As Integer = 1
        Dim strTemp() As String
        Dim strNomFichier As String

        Do While Session(strLinks & index & "set") <> Nothing
            If Session(strLinks & index & "set") = "true" Then
                strTemp = Split(Session(strLinks & index).PostedFile.FileName, "\")
                strNomFichier = strTemp(UBound(strTemp))

                If strMode <> "R" Then
                    lblLinks.Text += "<input class=""SmallButton"" type=""Checkbox"" name=""Session" & index & """ />"
                End If

                lblLinks.Text += "<a target=""new"" href=""" & Session(strLinks & index).PostedFile.FileName & """>" & strNomFichier & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
            End If
            index += 1
        Loop

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| PutsLinksFromDBLabel: Puts graph links inserted in the databse in the correct Label                              | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub PutsLinksFromDBLabel(ByVal dbConn As OleDbConnection, ByVal intNoQbr As Integer, ByVal strLinks As String, _
                                ByRef lblLinks As Label, ByVal Type As Integer, ByVal strMode As String)

        Dim strReq As String = "Select LinkNo, Path from Graphs where Type=" & Type & " and QbrNo = " & intNoQbr & " order by LinkNo"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dtTable As New DataTable
        Dim strNomFichier() As String
        Dim myRow As DataRow

        cmdTable.Fill(dtTable)

        For Each myRow In dtTable.Rows
            strNomFichier = Split(myRow("Path"), "\")
            If Not RemoveExist(myRow("LinkNo"), strLinks) Then
                If strMode <> "R" Then
                    lblLinks.Text += "<input class=""SmallButton"" type=""Checkbox"" name=""DB" & myRow("LinkNo") & """ />"
                End If
                lblLinks.Text += "<a target=""new"" href=""" & myRow("Path") & """>" & strNomFichier(UBound(strNomFichier)) & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
            End If
        Next

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ResetSessionGraphs: Reset all the Session Variables we could have used concerning Graphs                         | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub ResetSessionGraphs()
        Dim strGraphs() As String = {"linksSituation", "linksSolution", "linksResult"}

        Dim index As Integer = 0
        Dim index1 = 1

        For index = 0 To UBound(strGraphs)
            Session(strGraphs(index) & "Remove") = Nothing
            Do While Session(strGraphs(index) & index1 & "set") <> Nothing
                Session(strGraphs(index) & index1 & "set") = Nothing
                Session(strGraphs(index) & index1) = Nothing
                index1 += 1
            Loop
        Next
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| DeleteGraphs: Delete the desired graphs                                                                          | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub DeleteGraphs(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim strGraphs As String = Mid(sender.id, "DeleteGraph".Length + 1, sender.id.length - "DeleteGraph".Length - 1)
        Dim index As String = Mid(sender.id, sender.id.length, 1)
        Dim index1 As Integer = 1
        Dim dbConn As OleDbConnection

        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As DataTable
        Dim myRow As DataRow

        EstablishConnection(dbConn)
        dbConn.Open()

        index1 = 1

        'Retirer de la base de données
        strReq = "Select LinkNo from Graphs where QbrNo = " & Qbrcache.Text & " and type = " & index
        cmdTable = New OleDbDataAdapter(strReq, dbConn)
        dtTable = New DataTable
        cmdTable.Fill(dtTable)
        For Each myRow In dtTable.Rows

            If Not RemoveExist(myRow("LinkNo"), strGraphs) Then
                If Request.Form("DB" & myRow("LinkNo")) = "on" Then

                    Session(strGraphs & "Remove") += myRow("LinkNo") & ","
                End If
            End If
        Next


        Do While Session(strGraphs & index1 & "set") <> Nothing

            'Remove from session Variable                 
            If Session(strGraphs & index1 & "set") = "true" Then
                If Request.Form("Session" & index1) = "on" Then
                    Session(strGraphs & index1 & "set") = "false"
                End If
            End If
            index1 += 1
        Loop

        ShowGraphs(dbConn, Qbrcache.Text, "I")
        dbConn.Close()
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| RemoveExist: Verifies if the file has already been asked to remove                                               | 
    '|------------------------------------------------------------------------------------------------------------------|
    Function RemoveExist(ByVal strNoLink As String, ByVal strType As String)
        Dim strTLinks() As String = Split(Session(strType & "Remove"), ",")
        Dim index As Integer
        Dim Exists As Boolean = False

        For index = 0 To UBound(strTLinks)
            If strTLinks(index) <> "" Then
                If strTLinks(index) = strNoLink Then
                    Exists = True
                End If
            End If
        Next

        RemoveExist = Exists
    End Function

    '|------------------------------------------------------------------------------------------------------------------|
    '| SendMailOutlook: Opens a new email window from outlook                                                           | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub SendMailOutlook(ByVal sender As Object, ByVal e As System.EventArgs)
        If Qbrcache.Text <> -1 Then
            Server.Transfer("Email.aspx?Nu=" & Qbrcache.Text)
        Else
            Response.Write("<script language=Javascript>alert(""The form must be saved first"");</script>")
        End If
    End Sub


    '|------------------------------------------------------------------------------------------------------------------|
    '| PreviewMail: Open the report in view mode                                                                       | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub PreviewMail(ByVal sender As Object, ByVal e As System.EventArgs)
        If txtTo.Text <> "" And txtFrom.Text <> "" Then
            Session("Print") = "Email"
            Server.Transfer("Print.aspx")
        Else
            Response.Write("<script language=Javascript>alert(""You must at least enter the recipient and the sender's name."");</script>")
        End If
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| SendMail: Sends the email                                                                                        | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub SendMail(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim EnvoiCourriel As MailMessage = New MailMessage

        SendEmail.Visible = False
        Try

            EnvoiCourriel.From = txtFrom.Text
            EnvoiCourriel.To = txtTo.Text
            EnvoiCourriel.Cc = TxtCC.Text
            EnvoiCourriel.Bcc = TxtBCC.Text
            EnvoiCourriel.Subject = txtSubject.Text

            'Attachements

            'Attach logo pictures 
            Dim attach As New MailAttachment(Server.MapPath("images\LAURENTIDE_LOGO.bmp"))
            EnvoiCourriel.Attachments.Add(attach)
            attach = New MailAttachment(Server.MapPath("images\emerson-small.bmp"))
            EnvoiCourriel.Attachments.Add(attach)

            'Graphs Attachments
            GraphAttachements(EnvoiCourriel.Attachments)

            Dim tw As New System.IO.StringWriter
            Dim hw As New HtmlTextWriter(tw)
            Page.RenderControl(hw)

            EnvoiCourriel.Body = tw.ToString()
            EnvoiCourriel.BodyFormat = MailFormat.Html

            hw.Close()
            tw.Close()
            EnvoiCourriel.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtauthenticate", 2)
            EnvoiCourriel.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserver", "lcl-exc")
            'SmtpMail.SmtpServer = "lcl-exc"
            SmtpMail.Send(EnvoiCourriel)

            Response.Redirect("Redirect.aspx?Nu=" & NoQbr.Text)

        Catch ex As Exception
            'If Error occures
            If ex.Message <> "Thread was being aborted." Then
                Response.Redirect("RedirectError.aspx?Nu=" & NoQbr.Text)
            End If
        End Try

        SendEmail.Visible = True
    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| GraphAttachements: attach the web links to the mail message                                                      | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub GraphAttachements(ByRef Attachement As IList)
        Dim dbconn As OleDbConnection

        EstablishConnection(dbconn)
        dbconn.Open()

        lblLinksSituation.Text = ""
        lblLinksSolution.Text = ""
        lblLinksResult.Text = ""

        'Show Links from Database
        PutsLinksFromDBLabelEmail(dbconn, Val(NoQbr.Text), "LinksSituation", lblLinksSituation, 0, Attachement)
        PutsLinksFromDBLabelEmail(dbconn, Val(NoQbr.Text), "LinksSolution", lblLinksSolution, 1, Attachement)
        PutsLinksFromDBLabelEmail(dbconn, Val(NoQbr.Text), "LinksResult", lblLinksResult, 2, Attachement)

        dbconn.Close()

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| PutsLinksFromDBLabelEmail: Puts graph links inserted in the databse in the correct Label For Email               | 
    '|------------------------------------------------------------------------------------------------------------------|
    Sub PutsLinksFromDBLabelEmail(ByVal dbConn As OleDbConnection, ByVal intNoQbr As Integer, ByVal strLinks As String, _
                                ByRef lblLinks As Label, ByVal Type As Integer, ByRef Attachement As IList)

        Dim strReq As String = "Select LinkNo, Path from Graphs where Type=" & Type & " and QbrNo = " & intNoQbr & " order by LinkNo"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dtTable As New DataTable
        Dim strNomFichier() As String
        Dim myRow As DataRow
        Dim attach As MailAttachment

        cmdTable.Fill(dtTable)

        For Each myRow In dtTable.Rows
            strNomFichier = Split(myRow("Path"), "\")
            If Not RemoveExist(myRow("LinkNo"), strLinks) Then

                'Ajouter l'attachement
                attach = New MailAttachment(Server.MapPath(myRow("Path")))
                Attachement.Add(attach)
                lblLinks.Text += "<a target=""new"" href=""" & strNomFichier(UBound(strNomFichier)) & """>" & strNomFichier(UBound(strNomFichier)) & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
            End If
        Next

    End Sub

    '|------------------------------------------------------------------------------------------------------------------|
    '| ReturnField: Return the searched value from strNomTable                                                          | 
    '|------------------------------------------------------------------------------------------------------------------|
    Function ReturnField(ByVal dbConn As OleDbConnection, ByVal strNomTable As String, ByVal strColonne As String, ByVal strColonneRecherche As String, ByVal strValeur As String) As String
        Dim strReq = "select " & strColonne & " from " & strNomTable & " where " & strColonneRecherche & " = " & strValeur
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConn)
        Dim dttable As New DataTable
        Dim strRetour As String

        If strValeur <> "" Then
            cmdTable.Fill(dttable)

            If Not IsDBNull(dttable.Rows.Item(0)(strColonne)) Then
                strRetour = dttable.Rows.Item(0)(strColonne)
            End If
        End If

        ReturnField = strRetour
    End Function

    Private Sub InitializeComponent()

    End Sub

End Class

