Public Class frmTruncateAll

    Public Sub StartCancelStop(Optional pblnNoAsk As Boolean = False)
        Select Case cmdStartCancelExit.Text
            Case "Start"
                If pblnNoAsk OrElse MsgBox("Do you whant to Truncate All Datamart Data?", vbYesNo Or vbQuestion) = vbYes Then
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

    Private Sub cmdStartCancelExit_Click(sender As System.Object, e As System.EventArgs) Handles cmdStartCancelExit.Click
        Call StartCancelStop()
    End Sub

    Private Sub lp_ImportData()

        Dim sResulta As String = "Starting Truncate all at " & Now.ToString & vbCrLf
        Dim sError As String = ""
        Dim oEx As Exception = Nothing

        If Not goGlobalCancel Then
            Me.Text = "Truncate - Importing Phoenix..." : Application.DoEvents()
            oEx = Nothing
            sResulta += cTruncateAll.Execute(oEx) & vbCrLf
            Me.txtResultado.Text = sResulta
            If Not oEx Is Nothing Then
                sError += "Error Truncate all: " & oEx.ToString & vbCrLf
                sResulta += sError
            End If
        End If

        Me.Text = "Truncate all - Done..." : Application.DoEvents()
        cmdStartCancelExit.Text = "Close"
        '
        If sError <> "" Then
            sResulta += "Truncate all with ERRORS at " & Now.ToString & vbCrLf
            Call GrabarLog(eLogType.eERROR, sError)
        Else
            sResulta += "Truncate all: at " & Now.ToString & vbCrLf
        End If
        
    End Sub

    Private Sub frmImportAll_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        cmdStartCancelExit.Text = "Start"
        goGlobalCancel = False
    End Sub

End Class