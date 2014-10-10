Public Class frmImportBizContent

    Private Sub cmdStartCancelExit_Click(sender As System.Object, e As System.EventArgs) Handles cmdStartCancelExit.Click
        Select Case cmdStartCancelExit.Text
            Case "Start"
                If MsgBox("Do you whant to import BIZ Content's Data?", vbYesNo Or vbQuestion) = vbYes Then
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
        Me.txtResultado.Text = cImportContentGeneric.Import("[BIZCONTENT]",Me.pgbGlobal, Me.pgbCurrent, Me.lblTable, Me.lblCurrentOp, GC_BizSystem)
        cmdStartCancelExit.Text = "Close"
    End Sub

    Private Sub frmImportContent_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        cmdStartCancelExit.Text = "Start"
    End Sub
End Class
