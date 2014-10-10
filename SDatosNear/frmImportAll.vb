Public Class frmImportAll

    Public Sub StartCancelStop(Optional pblnNoAsk As Boolean = False)
        Select Case cmdStartCancelExit.Text
            Case "Start"
                If pblnNoAsk OrElse MsgBox("Do you want to import Datamart Data?", vbYesNo Or vbQuestion) = vbYes Then
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
        Dim sResulta As String = "Starting all import at " & Now.ToString & vbCrLf
        Dim sError As String = ""
        Dim oEx As Exception = Nothing
        Dim oSemaforo As New CSemaphore

        ' Marco el arranque del proceso
        oSemaforo.BeginProcess()

        ' Mientras va importando Sales Force, arranco con Content
        If Process_BatSFDCin Then
            System.Diagnostics.Process.Start("C:\DataSFDC\ImportAll.bat")
        End If
        
        If Not goGlobalCancel Then
            'Me.Text = "Truncating All ..." : Application.DoEvents()
            'oEx = Nothing
            'sResulta += cTruncateAll.Execute(oEx) & vbCrLf
            'Me.txtResultado.Text = sResulta
            'If Not oEx Is Nothing Then
            ' sError += "Error Truncating All: " & oEx.ToString & vbCrLf
            'sResulta += sError
            'End If
        End If


        If Not goGlobalCancel Then
            Me.Text = "Import All - Importing Phoenix..." : Application.DoEvents()
            oEx = Nothing
            sResulta += cImportPhoenixGeneric.Import("[PHOENIX]", Me.pgbGlobal, Me.pgbCurrent, Me.lblCurrentOp, Me.lblTable, GC_EduSystem, oEx) & vbCrLf
            Me.txtResultado.Text = sResulta
            If Not oEx Is Nothing Then
                sError += "Error importing Phoenix: " & oEx.ToString & vbCrLf
                sResulta += sError
            End If
        End If

        If Not goGlobalCancel Then
            Me.Text = "Import All - Importing SalesForce..." : Application.DoEvents()
            oEx = Nothing
            sResulta += cImportSF.Import(Me.pgbGlobal, Me.pgbCurrent, Me.lblCurrentOp, Me.lblTable, , oEx) & vbCrLf
            Me.txtResultado.Text = sResulta
            If Not oEx Is Nothing Then
                sError += "Error importing SalesForce: " & oEx.ToString & vbCrLf
                sResulta += sError
            End If
        End If

        If Not goGlobalCancel Then
            Me.Text = "Import All - Importing Content..." : Application.DoEvents()
            oEx = Nothing
            sResulta += cImportContentGeneric.Import("[CONTENT]", Me.pgbGlobal, Me.pgbCurrent, Me.lblTable, Me.lblCurrentOp, GC_EduSystem, oEx) & vbCrLf
            Me.txtResultado.Text = sResulta
            If Not oEx Is Nothing Then
                sError += "Error importing Content: " & oEx.ToString & vbCrLf
                sResulta += sError
            End If
        End If

        If Not (noBusiness) Then
            If Not goGlobalCancel Then
                Me.Text = "Import All - Importing BIZ Phoenix..." : Application.DoEvents()
                oEx = Nothing
                sResulta += cImportPhoenixGeneric.Import("[BIZPHOENIX]", Me.pgbGlobal, Me.pgbCurrent, Me.lblCurrentOp, Me.lblTable, GC_BizSystem, oEx) & vbCrLf
                Me.txtResultado.Text = sResulta
                If Not oEx Is Nothing Then
                    sError += "Error importing Phoenix: " & oEx.ToString & vbCrLf
                    sResulta += sError
                End If
            End If

            If Not goGlobalCancel Then
                Me.Text = "Import All - Importing Biz Content..." : Application.DoEvents()
                oEx = Nothing
                sResulta += cImportContentGeneric.Import("[BIZCONTENT]", Me.pgbGlobal, Me.pgbCurrent, Me.lblTable, Me.lblCurrentOp, GC_BizSystem, oEx) & vbCrLf
                Me.txtResultado.Text = sResulta
                If Not oEx Is Nothing Then
                    sError += "Error importing Content: " & oEx.ToString & vbCrLf
                    sResulta += sError
                End If
            End If
        End If

        If Not goGlobalCancel Then
            Me.Text = "Import All - Importing Google..." : Application.DoEvents()
            oEx = Nothing
            sResulta += cImportGoogle.Import(Me.txtResultado, _
                Me.pgbGlobal, Me.pgbCurrent, _
                Me.lblCurrentOp, Me.lblTable, My.Settings.GoogleFolderIn, _
                My.Settings.GoogleFolderProcessed)

            Me.txtResultado.Text = sResulta
            If Not oEx Is Nothing Then
                sError += "Error importing Google: " & oEx.ToString & vbCrLf
                sResulta += sError
            End If
        End If

        If Not goGlobalCancel Then
            Me.Text = "Import All - Importing Apple..." : Application.DoEvents()
            oEx = Nothing
            sResulta += cImportApple.Import(Me.txtResultado, _
                Me.pgbGlobal, Me.pgbCurrent, _
                Me.lblCurrentOp, Me.lblTable, My.Settings.AppleFolderIn, _
                My.Settings.AppleFolderProcessed)

            Me.txtResultado.Text = sResulta
            If Not oEx Is Nothing Then
                sError += "Error importing Apple: " & oEx.ToString & vbCrLf
                sResulta += sError
            End If
        End If

        If Not goGlobalCancel Then
            Me.Text = "Import All - Importing PayPal..." : Application.DoEvents()
            oEx = Nothing
            sResulta += cImportPayPal.Import(Me.txtResultado, _
                Me.pgbGlobal, Me.pgbCurrent, _
                Me.lblCurrentOp, Me.lblTable, My.Settings.PayPalFolderIn, _
                My.Settings.PayPalFolderProcessed)

            Me.txtResultado.Text = sResulta
            If Not oEx Is Nothing Then
                sError += "Error importing PayPal: " & oEx.ToString & vbCrLf
                sResulta += sError
            End If
        End If

        If sError = "" Then
            If Not goGlobalCancel Then
                Me.Text = "Export SFDC - Exporting Stats to SalesForce..." : Application.DoEvents()
                oEx = Nothing
                sResulta += cExportSF.Export(Me.pgbGlobal, Me.pgbCurrent, Me.lblCurrentOp, Me.lblTable) & vbCrLf
                Me.txtResultado.Text = sResulta
                If Not oEx Is Nothing Then
                    sError += "Error Exporting to SalesForce Stats: " & oEx.ToString & vbCrLf
                    sResulta += sError
                End If
            End If
        Else
            sError += "----------> SKIP EXPORT SFDC " & vbCrLf
        End If

        If Not goGlobalCancel Then
            Me.Text = "Generate Transactions....." : Application.DoEvents()
            oEx = Nothing
            sResulta += cTransacSDNear.Generate(Me.pgbGlobal, Me.pgbCurrent, Me.lblCurrentOp, Me.lblTable) & vbCrLf
            Me.txtResultado.Text = sResulta
            If Not oEx Is Nothing Then
                sError += "Error Generate Transactions.....: " & oEx.ToString & vbCrLf
                sResulta += sError
            End If
        End If

        If Process_BatSFDCout Then
            System.Diagnostics.Process.Start("C:\DataSFDC\ExportAll.bat")
        End If


        Me.Text = "Import All - Done..." : Application.DoEvents()
        cmdStartCancelExit.Text = "Close"
        '
        If sError <> "" Then
            sResulta += "End all import with ERRORS at " & Now.ToString & vbCrLf
            Call GrabarLog(eLogType.eERROR, sError)
        Else
            sResulta += "End all import at " & Now.ToString & vbCrLf
        End If
        '
        oSemaforo.EndProcess()

        If sError <> "" Then
            cMail.SendMail("oscarw@nearpod.com", "sdatosnear@gmail.com", "Import data results with ERRORS", sResulta)
        Else
            cMail.SendMail("oscarw@nearpod.com", "sdatosnear@gmail.com", "Import data results", sResulta)
        End If

    End Sub

    Private Sub frmImportAll_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        cmdStartCancelExit.Text = "Start"
        goGlobalCancel = False
    End Sub

End Class