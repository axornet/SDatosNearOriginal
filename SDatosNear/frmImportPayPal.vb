Public Class frmImportPayPal
    Private Sub cmdStartCancelExit_Click(sender As System.Object, e As System.EventArgs) Handles cmdStartCancelExit.Click
        Select Case cmdStartCancelExit.Text
            Case "Start"
                If MsgBox("Do you whant to import PayPal's Data?", vbYesNo Or vbQuestion) = vbYes Then
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
        Me.txtResultado.Text = cImportPayPal.Import(Me.txtResultado, _
            Me.pgbGlobal, Me.pgbCurrent, _
            Me.lblCurrentOp, Me.lblTable, My.Settings.PayPalFolderIn, _
            My.Settings.PayPalFolderProcessed)
        cmdStartCancelExit.Text = "Close"
    End Sub


    Private Sub frmImportPayPal_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        cmdStartCancelExit.Text = "Start"
    End Sub
End Class