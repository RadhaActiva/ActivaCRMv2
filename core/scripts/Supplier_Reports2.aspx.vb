Imports System.Data.OleDb
Imports System.Data

Partial Class SupplierReports2
    Inherits System.Web.UI.Page

    Private objConn As New OleDbConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ToString)
    Protected ds1 As DataSet
    Protected strChkstatus As String
    Private strSuppid As String = String.Empty


#Region "Page Load"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            strSuppid = Request.Cookies("GAPS")("SupplierID")

            If Not strSuppid = String.Empty Then

                lblErrMsg.Text = ""

                Dim str As String = Page.Request.Cookies("GAPS")("lastlogindate")
                lblLastLoginDate.Text = Replace(str, "%2D", "-")

                Dim strSql As String = ""


                If Not Page.IsPostBack Then

                    PopulateRootLevel()

                    CreatedDateFrom.Text = DateTime.Today
                    CreatedDateTo.Text = DateTime.Today

                    Calendar1.Visible = False
                    Calendar2.Visible = False

                End If

                strSql = "SELECT id,title,parentid,link FROM tblSuppReportTypeHeader WHERE parentid IS NOT NULL ORDER BY parentid,id;SELECT id,title,parentid FROM tblSuppReportTypeHeader WHERE parentid IS NULL;"

                ds1 = getDataSet(strSql)

                Dim subTitle As String = String.Empty
                Dim dt As New DataTable
                dt = ds1.Tables(0)

                For Each dr As DataRow In dt.Rows
                    If dr("id") = Page.Request("id") And Page.Request("id") <> "" Then
                        subTitle = dr("title").ToString()
                    End If
                Next

                If Page.Request("PID") = 2 Then
                    lbltext.Text = "Repeat - " & subTitle
                Else
                    lbltext.Text = "Seasonal - " & subTitle
                End If
            Else
                Response.Write("Error: Invalid SupplierID")
                Exit Sub
            End If

        Catch ex As Exception
            Response.Write("Error:" & ex.Message)
        End Try
    End Sub

#End Region

#Region "Adding Details"
    Protected Sub btnAddOrders_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddOrders.Click
        Dim i As Integer

        If rdbLogSheet.Checked = True Then
            For i = 1 To lstLogSheet.Items.Count - 1
                If lstLogSheet.Items(i).Selected Then
                    lstActiveLogSheet.Items.Add(lstLogSheet.Items(i))
                End If
            Next

            For i = 0 To lstActiveLogSheet.Items.Count - 1
                lstLogSheet.Items.Remove(lstActiveLogSheet.Items(i))
            Next

            lstLogSheet.ClearSelection()
            lstActiveLogSheet.ClearSelection()
        End If
    End Sub
#End Region

#Region "Delete Details"

    Protected Sub btnRmvOrders_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRmvOrders.Click
        Dim i As Integer
        Dim str As ArrayList = New ArrayList(1)
        str.Clear()

        If rdbLogSheet.Checked = True Then
            For i = 0 To lstActiveLogSheet.Items.Count - 1
                If lstActiveLogSheet.Items(i).Selected Then
                    lstLogSheet.Items.Add(lstActiveLogSheet.Items(i).Text)
                    str.Add(lstActiveLogSheet.Items(i).Text)
                End If
            Next

            For i = 0 To str.Count - 1
                lstActiveLogSheet.Items.Remove(str(i).ToString)
            Next

            lstLogSheet.ClearSelection()
            lstActiveLogSheet.ClearSelection()

        End If
    End Sub

#End Region

#Region "Clear Controls"

    Protected Sub ClearControls()
        lstLogSheet.Items.Clear()
        lstActiveLogSheet.Items.Clear()

        lblMessage.Text = ""
        lblErrMsg.Text = ""
    End Sub

#End Region

#Region "Get Details Button"

    Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGo.Click
        Try
            Dim dt1, dt2 As Date
            Try
                lblErrMsg.Text = ""
                dt1 = DateValue(CreatedDateFrom.Text)
            Catch ex As Exception
                lblErrMsg.Text = " *Invalid From-Date entered!"
                Exit Sub
            End Try

            Try
                lblErrMsg.Text = ""
                dt2 = DateValue(CreatedDateTo.Text)
            Catch ex As Exception
                lblErrMsg.Text = " *Invalid To-Date entered!"
                Exit Sub
            End Try

            If dt1 > System.DateTime.Today Then
                lblErrMsg.Text = " *From date must not be greater than todays date!"
                Exit Sub
            End If

            If dt2 > DateAndTime.Today Then
                lblErrMsg.Text = " *To date must not be greater than todays date!"
                Exit Sub
            End If

            If dt1 <= dt2 Then
                Call LoadData(dt1, dt2)
            End If

        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub

#End Region

#Region "Calender Operations"

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        lblErrMsg.Text = ""
        If Calendar1.SelectedDate > System.DateTime.Today Then
            lblErrMsg.Text = " *From Date should not be greater than todays date!"
        Else
            CreatedDateFrom.Text = Calendar1.SelectedDate
            Calendar1.Visible = False
        End If
    End Sub

    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click
        Calendar1.Visible = True
    End Sub

    Protected Sub ImageButton2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton2.Click
        Calendar2.Visible = True
    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged
        lblErrMsg.Text = ""
        If Calendar2.SelectedDate > System.DateTime.Today Then
            lblErrMsg.Text = " *To Date should not be greater than todays date!"
        Else
            CreatedDateTo.Text = Calendar2.SelectedDate
            Calendar2.Visible = False
        End If

    End Sub

#End Region

#Region "TreeView Operations"

    Protected Sub PopulateRootLevel()

        Dim objCommand As New OleDbCommand()
        objCommand.CommandText = "SELECT id,title,(SELECT count(*) FROM tblSuppReportTypeHeader WHERE parentid=A.id) childnodecount FROM tblSuppReportTypeHeader A WHERE parentID IS NULL"
        objCommand.Connection = objConn

        Dim da As New OleDbDataAdapter(objCommand)
        Dim dt As New DataTable()
        da.Fill(dt)

        PopulateNodes(dt, TreeView1.Nodes)
    End Sub
    Protected Sub PopulateSubNodes(ByVal dt As DataTable, ByVal nodes As TreeNodeCollection)
        For Each dr As DataRow In dt.Rows
            Dim tn As New TreeNode()
            tn.Text = dr("title").ToString
            tn.Value = dr("id").ToString
            nodes.Add(tn)
        Next
    End Sub
    Protected Sub PopulateNodes(ByVal dt As DataTable, ByVal nodes As TreeNodeCollection)
        For Each dr As DataRow In dt.Rows
            Dim tn As New TreeNode()
            tn.Text = dr("title").ToString
            tn.Value = dr("id").ToString
            nodes.Add(tn)

            'If node has child nodes, then enable on-demand populating
            tn.PopulateOnDemand = (CInt(dr("childnodecount")) > 0)

        Next
    End Sub
    Protected Sub PopulateSubLevel(ByVal parentid As Integer, ByVal parentNode As TreeNode)
        Try

            Dim objCommand As New OleDbCommand

            objCommand.CommandText = "SELECT id,title,parentid,(SELECT count(*) FROM tblSuppReportTypeHeader WHERE parentid=A.id) AS childnodecount FROM tblSuppReportTypeHeader A WHERE parentid=" & parentid & " ORDER BY parentid,id"
            objCommand.Connection = objConn

            Dim da As New OleDbDataAdapter(objCommand)
            Dim dt As New DataTable()
            da.Fill(dt)

            PopulateSubNodes(dt, parentNode.ChildNodes)
        Catch ex As Exception
            Response.Write(ex.Message)

        End Try

    End Sub
    Protected Sub TreeView1_SelectedNodeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TreeView1.SelectedNodeChanged
        For Each dr As DataRow In ds1.Tables(1).Rows
            If TreeView1.SelectedNode.Value = CStr(dr("id")) Then
                Response.Redirect("Supplier_Reports.aspx?id=" + CStr(dr("id")) + "&PID=" + CStr(dr("id")))
            End If
        Next
        For Each dr As DataRow In ds1.Tables(0).Rows
            If TreeView1.SelectedNode.Value = CStr(dr("id")) Then
                Response.Redirect(dr("link") + "?id=" + CStr(dr("id")) + "&PID=" + CStr(dr("parentid")))
            End If
        Next

    End Sub
    Protected Sub TreeView1_TreeNodePopulate(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.TreeNodeEventArgs) Handles TreeView1.TreeNodePopulate
        PopulateSubLevel(CInt(e.Node.Value), e.Node)
    End Sub

#End Region

#Region "Load Data To Controls"

    Protected Sub LoadData(ByVal dtFrom As Date, ByVal dtTo As Date)

        lblMessage.Text = "Selected details from {" & dtFrom & "} to {" & dtTo & "} "

        Dim strSql As String = ""
        Try
            If Page.Request("PID") = 1 Then
                strSql = "SELECT DISTINCT tblOrderAllocationList.LogSheetNo FROM tblOrderAllocationList INNER JOIN tblMasterAllocationList ON tblOrderAllocationList.MALNo=tblMasterAllocationList.MALNo INNER JOIN tblLogSheet ON tblOrderAllocationList.LogSheetNo=tblLogSheet.LogSheetNo " & _
                         "WHERE [tblMasterAllocationList].[SupplierID]='" & strSuppid & "' AND [tblMasterAllocationList].[MATotalOrderqty] <>0 AND tblOrderAllocationList.Status='ShippedTL' AND CONVERT(DATE,tblLogSheet.CreatedDate,101) BETWEEN '" & dtFrom & "' AND '" & dtTo & "' ORDER BY tblOrderAllocationList.LogSheetNo"
            ElseIf Page.Request("PID") = 2 Then
                strSql = "SELECT DISTINCT tblOrderRepeatList.LogSheetNo FROM tblOrderRepeatList INNER JOIN tblLogSheet ON tblOrderRepeatList.LogSheetNo=tblLogSheet.LogSheetNo " & _
                         "WHERE tblOrderRepeatList.Status='ShippedTL' AND CONVERT(DATE,tblLogSheet.CreatedDate,101) BETWEEN '" & dtFrom & "' AND '" & dtTo & "' ORDER BY tblOrderRepeatList.LogSheetNo"
            End If

            Dim ds As DataSet
            ds = getDataSet(strSql)

            With lstLogSheet
                .DataSource = ds.Tables(0)
                .DataValueField = "LogSheetNo"
                .DataTextField = "LogSheetNo"
                .DataBind()
                .Items.Insert(0, "--Select LogSheets--")
            End With

        Catch ex As Exception
            Response.Write("<b>Error:</b>" & ex.Message)
        End Try
    End Sub

#End Region

#Region "Get Report Button"

    Protected Sub btnGetReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGetReport.Click
        Dim i As Integer
        Dim rpt As Integer = 0
        Dim strLogSheet As String = String.Empty

        Dim rptName As String = String.Empty
        Dim rptQuery As String = String.Empty


        For i = 0 To lstActiveLogSheet.Items.Count - 1
            If i = 0 Then
                strLogSheet = "'" + Trim(lstActiveLogSheet.Items(i).ToString()) + "'"
            Else
                strLogSheet = strLogSheet + ",'" + Trim(lstActiveLogSheet.Items(i).ToString()) + "'"
            End If
        Next

        If Not strLogSheet = String.Empty Then

            If Page.Request("PID") = 1 Then
                rpt = 11
                rptName = "rptSeasonLogSheet.rpt"
                rptQuery = "SELECT [tblOrderAllocationList].[OrderNo], [tblOrderAllocationList].[MALNo], [tblOrderAllocationList].[OutletID], [tblOrderAllocationList].[Status], " & _
                            "[tblOrderAllocationList].[OrderType], [tblOrderAllocationList].[Remark], [tblOrderAllocationList].[SeasonCodeID], [tblOrderAllocationList].[VoidDate], " & _
                            "[tblOrderAllocationList].[VoidUserID], [tblOrderAllocationList].[PostDate], [tblOrderAllocationList].[PostUserID], [tblOrderAllocationList].[ShippedDate], " & _
                            "[tblOrderAllocationList].[ShippedUserID], [tblOrderAllocationList].[ReceiveDate], [tblOrderAllocationList].[ReceiveUserID], [tblOrderAllocationList].[DeliveredDate], " & _
                            "[tblOrderAllocationList].[DeliveredUserID], [tblOrderAllocationList].[AcceptedDate], [tblOrderAllocationList].[AcceptedUserID], " & _
                            "[tblOrderAllocationList].[RODate], [tblOrderAllocationList].[ROUserID], [tblOrderAllocationList].[RejectDate], [tblOrderAllocationList].[RejectUserID], " & _
                            "[tblOrderAllocationList].[ApprovedDate], [tblOrderAllocationList].[ApprovedUserID], [tblOrderAllocationList].[OrderQty], [tblOrderAllocationList].[OrderCartonQty], " & _
                            "[tblOrderAllocationList].[OrderTotalAmt], [tblOrderAllocationList].[OrderCostAmt], [tblOrderAllocationList].[shippedQty], [tblOrderAllocationList].[shippedRemark], " & _
                            "[tblOrderAllocationList].[shippedCartonQty], [tblOrderAllocationList].[shippedCartonAmt], [tblOrderAllocationList].[shippedCostAmt], [tblOrderAllocationList].[ReceivedQty], " & _
                            "[tblOrderAllocationList].[ReceivedCartonQty], [tblOrderAllocationList].[ReceivedTotalAmt], [tblOrderAllocationList].[ReceivedCostAmt], [tblOrderAllocationList].[DeliveredQty], " & _
                            "[tblOrderAllocationList].[DeliveredCartonQty], [tblOrderAllocationList].[DeliveredTotalAmt], [tblOrderAllocationList].[DeliveredCostAmt], [tblOrderAllocationList].[ROQty], " & _
                            "[tblOrderAllocationList].[ROCartonQty], [tblOrderAllocationList].[ROTotalAmt], [tblOrderAllocationList].[RORemark], [tblOrderAllocationList].[ROCostAmt], " & _
                            "[OutletMaster].[OutletLocation],[tblOrderAllocationList].[CreatedDate],[tblMasterAllocationList].[ExpiryDate] AS MALExpiryDate,[tblMasterAllocationList].[shippeddate] AS MALShipmentDate, " & _
                            "[tblMasterAllocationList].[DeliveryDate] AS MALDeliveryDate, tblMasterAllocationList].[SupplierID] AS SupplierId, " & _
                            "'trans1' AS [Transporter],0,[tblLogSheet].[LogSheetNo],[tblLogSheet].[SupplierId] AS TransId,[SUPPLIER].[SupplierName] AS TransName, [OutletArea].[OutletArea],[OutletMaster].[BoContactNo],[tblLogSheet].[CreatedDate] As LogSheetDate " & _
                            "FROM [tblOrderAllocationList] INNER JOIN [OutletMaster] ON [tblOrderAllocationList].[OutletID]=[OutletMaster].[OutletID] " & _
                            "INNER JOIN [OutletArea] ON [OutletArea].[OutletAreaID] = [OutletMaster].[OutletAreaID] " & _
                            "INNER JOIN [tblLogSheet] ON [tblLogSheet].[LogSheetNo] = [tblOrderAllocationList].[LogSheetNo]  " & _
                            "INNER JOIN [Supplier] ON [tblLogSheet].[SupplierId]=[Supplier].[SupplierId] " & _
                            "INNER JOIN [tblMasterAllocationList] ON [tblOrderAllocationList].[MALNo]=[tblMasterAllocationList].[MALNo] " & _
                            "WHERE [tblMasterAllocationList].[SupplierID]='" & strSuppid & "' AND [tblMasterAllocationList].[MATotalOrderqty] <>0 AND [tblOrderAllocationList].[STATUS]='ShippedTL' AND tblOrderAllocationList.LogSheetNo IN (" + strLogSheet + ")"
            End If

            If Page.Request("PID") = 2 Then
                rpt = 38
                rptName = "rptRepeatLogSheet.rpt"
                rptQuery = "SELECT [tblOrderRepeatList].[OrderNo],[tblOrderRepeatList].[OutletID],[tblOrderRepeatList].[Supplierid],[tblOrderRepeatList].[Status], " & _
                            "[tblOrderRepeatList].[OrderDeliveryDate],[tblOrderRepeatList].[ShippedDate],[tblOrderRepeatList].[ReceiveDate],[tblOrderRepeatList].[DeliveredDate], " & _
                            "[tblOrderRepeatList].[AcceptedDate],[tblOrderRepeatList].[RODate],[tblOrderRepeatList].[ROQty],[tblOrderRepeatList].[OrderDC], " & _
                            "[tblOrderRepeatList].[ShippedCartonQty] AS [OrderCartonQty],[tblOrderRepeatList].[PostDate],'trans1' AS [Transporter],[tblLogSheet].[LogSheetNo], " & _
                            "[tblLogSheet].[SupplierId] AS TransId,[SUPPLIER].[SupplierName] AS TransName,[tblLogSheet].[CreatedDate] As LogSheetDate, " & _
                            "[OutletArea].[OutletArea], [OutletMaster].[BoContactNo], [OutletMaster].[OutletLocation] " & _
                            "FROM [tblOrderRepeatList], [OutletMaster], [OutletArea], [tblLogSheet], [SUPPLIER] " & _
                            "WHERE [OutletMaster].[OutletId] = [tblOrderRepeatList].[OutletId] And [OutletArea].[OutletAreaID] = [OutletMaster].[OutletAreaID] " & _
                            "AND [tblOrderRepeatList].[LogSheetNo]=[tblLogSheet].[LogSheetNo] AND [tblLogSheet].[SupplierId]=[Supplier].[SupplierId] " & _
                            "AND [tblOrderRepeatList].[STATUS]='ShippedTL' AND [tblOrderRepeatList].[LogSheetNo] IN (" & strLogSheet & ")"

            End If

            Session("rptName") = rptName
            Session("rptQuery") = rptQuery

            ClientScript.RegisterStartupScript(Me.GetType(), "onclick", "<script language='Javascript'>window.open('Supplier_Repeat_ReportViewer.aspx?ID=" & rpt & "'); </script>")

        End If
    End Sub

#End Region

#Region "DB Functions"

    Protected Function getDataSet(ByVal sqlStr As String) As DataSet
        Try

            Dim objCommand As New OleDb.OleDbCommand()

            objCommand.Connection = objConn
            objCommand.CommandText = sqlStr
            objCommand.CommandType = CommandType.Text

            Dim da As New OleDb.OleDbDataAdapter()
            da.SelectCommand = objCommand

            Dim ds As New DataSet
            da.Fill(ds)

            Return ds
        Catch ex As Exception
            Throw New Exception(ex.Message, ex.InnerException)
        End Try
    End Function

#End Region

End Class

