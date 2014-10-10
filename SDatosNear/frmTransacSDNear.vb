Public Class frmTransacSDNear

    Private Sub cmdStartCancelExit_Click(sender As System.Object, e As System.EventArgs) Handles cmdStartCancelExit.Click
        Select Case cmdStartCancelExit.Text
            Case "Start"
                If MsgBox("Do you whant to Execute the Create TransacSdNear", vbYesNo Or vbQuestion) = vbYes Then
                    cmdStartCancelExit.Text = "Cancel"
                    Call lp_GenerateTransac()
                End If
            Case "Cancel"
                If MsgBox("Do you want to cancel the process", vbYesNo Or vbQuestion) = vbYes Then
                    goGlobalCancel = True
                End If
            Case "Close"
                Me.Close()
        End Select
    End Sub

    Private Sub lp_GenerateTransac()
        Me.txtResultado.Text = cTransacSDNear.Generate(Me.pgbGlobal, Me.pgbCurrent, Me.lblCurrentOp, Me.lblTable)
        cmdStartCancelExit.Text = "Close"
    End Sub

    Private Sub frmApareao_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        cmdStartCancelExit.Text = "Start"
    End Sub



End Class