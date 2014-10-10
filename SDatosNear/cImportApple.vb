Public Class cImportApple

    Public Shared Function Import( _
            ptxtStatus As TextBox,
            pPgbGlobal As ProgressBar, pPgbParcial As ProgressBar, _
            plblCurrentOp As Label, plblTable As Label, _
            pvstrInFolder As String, pvstrMoveFolder As String,
            Optional ByRef pexError As Exception = Nothing) As String
        Dim sResulta As String = "Import Apple" & vbCrLf

        Try
            sResulta += "Start " & Now.ToString & vbCrLf
            ' Averiguo la lista de archivos a procesar
            Dim oFiles = GetFiles(pvstrInFolder, "*.txt")
            Dim ofileQuery = From file In oFiles _
                            Where file.Extension = ".txt" _
                            Order By file.Name _
                            Select file

            pPgbGlobal.Maximum = ofileQuery.Count
            pPgbGlobal.Value = 0
            If Not pvstrMoveFolder.EndsWith("\") Then
                pvstrMoveFolder += "\"
            End If
            For Each oFile In ofileQuery
                ProgressBarAdd(pPgbGlobal)
                plblTable.Text = "T_StatsApple"
                sResulta += gfstr_ImportBulkFromTxt(oFile.FullName, "T_StatsApple", False, pexError, plblCurrentOp, pPgbParcial, vbTab, True) & vbCrLf
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
        Call GrabarLog(eLogType.eAPPLE, sResulta)

        Return sResulta

    End Function


End Class
