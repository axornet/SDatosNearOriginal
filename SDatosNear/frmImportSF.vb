Public Class frmImportSF

    Private Sub cmdStartCancelExit_Click(sender As System.Object, e As System.EventArgs) Handles cmdStartCancelExit.Click
        If chkTables.CheckedItems.Count = 0 Then
            MsgBox("You must select at least one table", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Select Case cmdStartCancelExit.Text
            Case "Start"
                If MsgBox("Do you whant to import Sales Force's Data?", vbYesNo Or vbQuestion) = vbYes Then
                    cmdStartCancelExit.Text = "Cancel"
                    Call lp_ImportData()
                End If
            Case "Cancel"
                If MsgBox("Do you want to cancel the process", vbYesNo Or vbQuestion) = vbYes Then
                    goGlobalCancel = True
                End If
            Case "Close"
                Me.Close()
        End Select
    End Sub

    Private Sub lp_ImportData()
        Dim oTableCollection As New List(Of cImportSF.cSalesForceItem)
        For Each oItem As cImportSF.cSalesForceItem In chkTables.CheckedItems
            oTableCollection.Add(oItem)
        Next
        Me.txtResultado.Text = cImportSF.Import(Me.pgbGlobal, Me.pgbCurrent, Me.lblCurrentOp, Me.lblTable, oTableCollection)
        cmdStartCancelExit.Text = "Close"
    End Sub


    Private Sub frmImportSF_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        cmdStartCancelExit.Text = "Start"
        For Each oItem As cImportSF.cSalesForceItem In cImportSF.GetTablesCollection
            Me.chkTables.Items.Add(oItem, True)
        Next
    End Sub

    Private Sub cmdAll_Click(sender As System.Object, e As System.EventArgs) Handles cmdAll.Click
        Call SelectItems(True)
    End Sub

    Private Sub cmdNone_Click(sender As System.Object, e As System.EventArgs) Handles cmdNone.Click
        Call SelectItems(False)
    End Sub

    Private Sub SelectItems(pSelected As Boolean)
        For i As Integer = 0 To chkTables.Items.Count - 1
            chkTables.SetItemChecked(i, pSelected)
        Next
    End Sub

End Class