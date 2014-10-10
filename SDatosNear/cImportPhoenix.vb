Public Class cImportPhoenixGeneric
    Const C_LINKED As Boolean = True
    'Const C_DATABASE As String = "[PHOENIX]"
    Const C_DeltaNumINI As String = "1000000000"
    Const C_DeltaNumEND As String = "2000000000"
    Const C_DeltaText As String = """B_"""

    Public Shared Function Import( _
                pDataBase As String, _
                pPgbGlobal As ProgressBar, pPgbParcial As ProgressBar, _
                plblCurrentOp As Label, plblTable As Label, _
                pBizSystem As Boolean, _
                Optional ByRef pexError As Exception = Nothing) As String
        Dim lvstrExpSql As String
        Dim lvstrColumns As String
        Dim sResulta As String
        Dim sSpeacialSqlDelete As String

        If (pBizSystem) Then
            sResulta = "Import Biz Phoenix" & vbCrLf
        Else
            sResulta = "Import Phoenix" & vbCrLf
        End If

        Try
            sResulta += "Start " & Now.ToString & vbCrLf

            pPgbGlobal.Maximum = 6
            pPgbGlobal.Value = 1


            If goLocal251 Then
                If (pBizSystem) Then
                    lvstrExpSql = _
                        "Select " + IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                        "l.id " + IIf(pBizSystem, "+" & C_DeltaNumINI, "") & " as lead_uid, " & _
                        "s.teacher_id " + IIf(pBizSystem, "+" & C_DeltaNumINI, "") & " as teacher_id, " & _
                        IIf(pBizSystem, "concat(" & C_DeltaText & ",s.presentation_uid)", "s.presentation_uid") & " as presentation_uid, " & _
                        "s.homework, from_unixtime(s.timestamp) As session_date, " & _
                        IIf(pBizSystem, "concat(" & C_DeltaText & ",s.uid)", "s.uid") & " as session_uid, " & _
                        IIf(pBizSystem, "concat(" & C_DeltaText & ",l.device_uid)", "l.device_uid") & " as device_uid, " & _
                        "l.is_Deleted," & _
                        "q.qQuizSlides, q.qQuiz, q.qQuizDeleted, q.qQuizSkip, q.qQuizCorrect, " & _
                        "qa.qQA, qa.qQADeleted, qa.qQASkip, qa.qQACorrect, " & _
                        "p.qPoll, p.qPollDeleted, p.qPollSkip, " & _
                        "d.qDraw, d.qDrawDeleted, d.qDrawSkip  " & _
                        "from session s  " & _
                        "	left join lead l on s.uid = l.session_uid  " & _
                        "	left join v_quiz q   on l.session_uid = q.session_uid  and l.device_uid = q.device_uid  " & _
                        "	left join v_qa   qa  on l.session_uid = qa.session_uid and l.device_uid = qa.device_uid  " & _
                        "	left join v_poll p   on l.session_uid = p.session_uid  and l.device_uid = p.device_uid  " & _
                        "	left join v_drawit d on l.session_uid = d.session_uid  and l.device_uid = d.device_uid " & GC_LIMITRESULT
                Else
                    lvstrExpSql = _
                        "Select " + IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                        "l.id " + IIf(pBizSystem, "+" & C_DeltaNumINI, "") & " as lead_uid, " & _
                        "s.teacher_id " + IIf(pBizSystem, "+" & C_DeltaNumINI, "") & " as teacher_id, " & _
                        IIf(pBizSystem, "concat(" & C_DeltaText & ",s.presentation_uid)", "s.presentation_uid") & " as presentation_uid, " & _
                        "s.type, from_unixtime(s.timestamp) As session_date, " & _
                        IIf(pBizSystem, "concat(" & C_DeltaText & ",s.uid)", "s.uid") & " as session_uid, " & _
                        IIf(pBizSystem, "concat(" & C_DeltaText & ",l.device_uid)", "l.device_uid") & " as device_uid, " & _
                        "l.is_Deleted," & _
                        "q.qQuizSlides, q.qQuiz, q.qQuizDeleted, q.qQuizSkip, q.qQuizCorrect, " & _
                        "qa.qQA, qa.qQADeleted, qa.qQASkip, qa.qQACorrect, " & _
                        "p.qPoll, p.qPollDeleted, p.qPollSkip, " & _
                        "d.qDraw, d.qDrawDeleted, d.qDrawSkip  " & _
                        "from session s  " & _
                        "	left join lead l on s.uid = l.session_uid  " & _
                        "	left join v_quiz q   on l.session_uid = q.session_uid  and l.device_uid = q.device_uid  " & _
                        "	left join v_qa   qa  on l.session_uid = qa.session_uid and l.device_uid = qa.device_uid  " & _
                        "	left join v_poll p   on l.session_uid = p.session_uid  and l.device_uid = p.device_uid  " & _
                        "	left join v_drawit d on l.session_uid = d.session_uid  and l.device_uid = d.device_uid " & GC_LIMITRESULT
                End If

            Else
                lvstrExpSql = _
                    "Select " + IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                    "l.id " + IIf(pBizSystem, "+" & C_DeltaNumINI, "") & " as lead_uid, " & _
                    "s.teacher_id " + IIf(pBizSystem, "+" & C_DeltaNumINI, "") & " as teacher_id, " & _
                    IIf(pBizSystem, "concat(" & C_DeltaText & ",s.presentation_uid)", "s.presentation_uid") & " as presentation_uid, " & _
                    "s.type, from_unixtime(s.timestamp) As session_date, " & _
                    IIf(pBizSystem, "concat(" & C_DeltaText & ",s.uid)", "s.uid") & " as session_uid, " & _
                    IIf(pBizSystem, "concat(" & C_DeltaText & ",l.device_uid)", "l.device_uid") & " as device_uid, " & _
                    "l.is_Deleted," & _
                    "(select count(distinct slide) from quiz q where l.session_uid = q.session_uid and l.device_uid = q.device_uid) As qQuizSlides," & _
                    "(select count(*) from quiz q   where l.session_uid = q.session_uid and l.device_uid = q.device_uid  and not is_deleted) As qQuiz," & _
                    "(select count(*) from quiz q   where l.session_uid = q.session_uid and l.device_uid = q.device_uid and is_deleted) As qQuizDeleted," & _
                    "(select count(*) from quiz q   where l.session_uid = q.session_uid and l.device_uid = q.device_uid and is_skip  and not is_deleted) As qQuizSkip," & _
                    "(select count(*) from quiz q   where l.session_uid = q.session_uid and l.device_uid = q.device_uid and is_correct  and not is_deleted) As qQuizCorrect," & _
                    "(select count(*) from qa   a   where l.session_uid = a.session_uid and l.device_uid = a.device_uid and not is_deleted) As qQA," & _
                    "(select count(*) from qa   a   where l.session_uid = a.session_uid and l.device_uid = a.device_uid and is_deleted) As qQADeleted," & _
                    "(select count(*) from qa   a   where l.session_uid = a.session_uid and l.device_uid = a.device_uid and is_skip  and not is_deleted) As qQASkip," & _
                    "(select count(*) from qa   a   where l.session_uid = a.session_uid and l.device_uid = a.device_uid and is_correct  and not is_deleted) As qQACorrect," & _
                    "(select count(*) from poll p   where l.session_uid = p.session_uid and l.device_uid = p.device_uid  and not is_deleted) As qPoll," & _
                    "(select count(*) from poll p   where l.session_uid = p.session_uid and l.device_uid = p.device_uid and is_deleted) As qPollDeleted," & _
                    "(select count(*) from poll p   where l.session_uid = p.session_uid and l.device_uid = p.device_uid and is_skip  and not is_deleted) As qPollSkip," & _
                    "(select count(*) from drawit d where l.session_uid = d.session_uid and l.device_uid = d.device_uid  and not is_deleted) As qDraw," & _
                    "(select count(*) from drawit d where l.session_uid = d.session_uid and l.device_uid = d.device_uid and is_deleted) As qDrawDeleted," & _
                    "(select count(*) from drawit d where l.session_uid = d.session_uid and l.device_uid = d.device_uid and is_skip  and not is_deleted) As qDrawSkip " & _
                    "from session s left join lead l on s.uid = l.session_uid " & GC_LIMITRESULT

            End If
            If (pBizSystem) Then
                lvstrColumns = _
                "sd_source, lead_uid, teacher_id,presentation_uid,homework,session_date,session_uid,device_uid,is_deleted," & _
                "qQuizSlides,qQuiz, qQuizDeleted,qQuizSkip,qQuizCorrect,qQA,qQADeleted,qQASkip," & _
                "qQACorrect,qPoll,qPollDeleted,qPollSkip,qDraw,qDrawDeleted,qDrawSkip"
            Else
                lvstrColumns = _
                "sd_source, lead_uid, teacher_id,presentation_uid,type,session_date,session_uid,device_uid,is_deleted," & _
                "qQuizSlides,qQuiz, qQuizDeleted,qQuizSkip,qQuizCorrect,qQA,qQADeleted,qQASkip," & _
                "qQACorrect,qPoll,qPollDeleted,qPollSkip,qDraw,qDrawDeleted,qDrawSkip"
            End If

            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Sessions where sd_source = 1"
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizPhoenix, lvstrExpSql, "T_Sessions", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Sessions", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Sessions where sd_source = 0"
                    sResulta += gfstr_ImportaBulked(goConNear, goConnPhoenix, lvstrExpSql, "T_Sessions", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Sessions", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If

            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                If (pBizSystem) Then
                    pPgbParcial.Maximum = gflng_GetNumReg(goConnBizPhoenix, "SELECT COUNT(*) FROM lead")
                    sResulta += gfstr_Importa(goConnBizPhoenix, lvstrExpSql, goConNear, "T_Sessions", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
                Else
                    pPgbParcial.Maximum = gflng_GetNumReg(goConnPhoenix, "SELECT COUNT(*) FROM lead")
                    sResulta += gfstr_Importa(goConnPhoenix, lvstrExpSql, goConNear, "T_Sessions", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
                End If
            End If

            Dim oCmd As SqlClient.SqlCommand
            If (pBizSystem) Then
                ProgressBarAdd(pPgbGlobal)
                plblCurrentOp.Text = "Convert type to homework and embed ...." : Application.DoEvents()
                sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
                oCmd = goConNear.CreateCommand
                oCmd.CommandTimeout = 99999
                oCmd.CommandText = "update T_Sessions set type = Case when homework = 1 then 1 else 0 end where sd_source = 1 "
                oCmd.ExecuteNonQuery()
            Else
                ProgressBarAdd(pPgbGlobal)
                plblCurrentOp.Text = "Convert type to homework and embed ...." : Application.DoEvents()
                sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
                oCmd = goConNear.CreateCommand
                oCmd.CommandTimeout = 99999
                oCmd.CommandText = "update T_Sessions set homework = Case when type = 1 then 1 else 0 end, embed = case when type = 2 then 1 else 0 end where sd_source = 0 "
                oCmd.ExecuteNonQuery()
            End If
            
            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Session Size Calculation ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 99999
            oCmd.CommandText = "update T_Sessions set sessionSize = cont " & _
                               " from (select ts1.session_uid as session_uid, count(*) as cont  " & _
                               "        from T_Sessions ts1 " & _
                               "        where(ts1.is_Deleted = 0) " & _
                               "        and ts1.lead_uid > 0 " & _
                               "        group by ts1.session_uid ) as T1, T_Sessions T2 " & _
                               "where(T1.session_uid = T2.session_uid)"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Is Deleted ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 99999
            oCmd.CommandText = "update T_Sessions Set is_Deleted = 0 where is_deleted is null"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Homework ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 99999
            oCmd.CommandText = "update t_sessions set homework = 0 where homework = 1 and lead_uid is null"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Embed ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 99999
            oCmd.CommandText = "update t_sessions set embed = 0 where embed = 1 and lead_uid is null"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("T_SESSIONS", "session_date", "_Ses")

            plblTable.Text = "Done"
            plblCurrentOp.Text = "Done"
            pPgbParcial.Value = 0
            pPgbGlobal.Value = pPgbGlobal.Maximum

            sResulta += "End " & Now.ToString & vbCrLf

        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString
        End Try

        Call GrabarLog(eLogType.ePHOENIX, sResulta)
        Return sResulta

    End Function
End Class
