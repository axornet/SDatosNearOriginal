Public Class cImportPayPal
    Public Shared Function Import( _
            ptxtStatus As TextBox,
            pPgbGlobal As ProgressBar, pPgbParcial As ProgressBar, _
            plblCurrentOp As Label, plblTable As Label, _
            pvstrInFolder As String, pvstrMoveFolder As String,
            Optional ByRef pexError As Exception = Nothing) As String
        Dim sResulta As String = "Import PayPal" & vbCrLf

        Try
            sResulta += "Start " & Now.ToString & vbCrLf
            ' Averiguo la lista de archivos a procesar
            Dim oFiles = GetFiles(pvstrInFolder, "*.csv")
            Dim ofileQuery = From file In oFiles _
                            Where file.Extension = ".csv" _
                            Order By file.Name _
                            Select file

            pPgbGlobal.Maximum = ofileQuery.Count
            pPgbGlobal.Value = 0
            If Not pvstrMoveFolder.EndsWith("\") Then
                pvstrMoveFolder += "\"
            End If
            For Each oFile In ofileQuery
                ProgressBarAdd(pPgbGlobal)
                plblTable.Text = "T_PayPalTxn"
                sResulta += gfstr_ImportBulkFromTxt(oFile.FullName, "T_PayPalTxn", False, pexError, plblCurrentOp, pPgbParcial, ",", True, New String() {"TRANSACTION_ID"}) & vbCrLf
                If pexError Is Nothing Then
                    My.Computer.FileSystem.MoveFile(oFile.FullName, pvstrMoveFolder & oFile.Name)
                End If
                ptxtStatus.Text = sResulta
                If goGlobalCancel Then Exit For
            Next

            plblTable.Text = "Done"
            plblCurrentOp.Text = "Done"
            pPgbParcial.Value = 0
            pPgbGlobal.Value = pPgbGlobal.Maximum

            sResulta += "End " & Now.ToString & vbCrLf

        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf
        End Try
        Call GrabarLog(eLogType.ePAYPAL, sResulta)

        Return sResulta

    End Function




End Class
