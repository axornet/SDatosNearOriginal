Public Class cExportSF
    Public Shared Function Export( _
        ByRef pPgbGlobal As ProgressBar, _
        pPgbParcial As ProgressBar, _
        plblCurrentOp As Label, _
        plblTable As Label, _
        Optional ByRef pexError As Exception = Nothing) As String

        Dim sResulta As String = "Export SalesForce" & vbCrLf
        Dim lvstrSql As String
        Dim lvlngNumReg As Integer
        Dim oCmd As SqlClient.SqlCommand

        Try
            sResulta += "Start all process" & Now.ToString & vbCrLf

            pPgbGlobal.Maximum = 20
            pPgbGlobal.Value = 0

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Deleting Temp Table" : Application.DoEvents()
            oCmd.CommandText = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TEMP_EXPORTSFDC]') AND type in (N'U')) " & _
                              " DROP TABLE [dbo].[TEMP_EXPORTSFDC]"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Temp Table (Students Use)" : Application.DoEvents()
            oCmd.CommandText = _
                "SELECT TU.[sforceId] as ID, " & _
                "TU.[firstName] as FIRSTNAME, " & _
                "TU.[lastName] as LASTNAME, " & _
                "TU.[email] as EMAIL, " & _
                "TU.[sfType] as USERTYPE__C, " & _
                "CASE WHEN [TU].isdeleted = 1 or [TU].active = 0 THEN getdate() ELSE NULL END as INACTIVE__C, " & _
                "CASE WHEN [TU].[mailAnnouncements] = 'True' THEN 'FALSE' ELSE 'TRUE' END as HasOptedOutOfEmail, " & _
                "CASE  [TU].sdn_Stage " & _
                "   WHEN '20' then 'Stage 020 - Presentation Executed' " & _
                "   WHEN '30' then 'Stage 030 - Upgraded User' " & _
                "   WHEN '40' then 'Stage 040 - Presentation Created' " & _
                "   WHEN '50' then 'Stage 050 - Use in Class' " & _
                "   WHEN '55' then 'Stage 055 - Created 3 presentations' " & _
                "   WHEN '60' then 'Stage 060 - Heavy User' " & _
                "   ELSE 'Stage 010 - App downloaded' " & _
                " END as User_Stage__c, " & _
                "TU.[sdn_Stage40Date] as Created_Presentation_040__c, " & _
                "TU.[sdn_Stage55Date] as Created_X_presentations_055__c," & _
                "TU.[sdn_Stage60Date] as Heavy_User_060__c, " & _
                "TU.[sdn_Stage20Date] as Experienced_Nearpod__c, " & _
                "0 as USX_Exceed_Stud_Limit__c, " & _
                "0 as USX_Exceed07_Stud_Limit__c," & _
                "0 as USX_Exceed30_Stud_Limit__c," & _
                "(SELECT Count(Distinct dbo.T_Sessions.lead_uid)     FROM   dbo.T_Sessions   WHERE  dbo.T_Sessions.teacher_id = TU.id AND T_sessions.is_deleted = 0         AND  DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 7) As USX_Tot07_Students__c,    " & _
                "(SELECT Count(Distinct dbo.T_Sessions.lead_uid)     FROM   dbo.T_Sessions   WHERE  dbo.T_Sessions.teacher_id = TU.id AND T_sessions.is_deleted = 0         AND  DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 30) As USX_Tot30_Students__c,    " & _
                "(SELECT Count(Distinct dbo.T_Sessions.lead_uid)     FROM   dbo.T_Sessions   WHERE  dbo.T_Sessions.teacher_id = TU.id AND T_sessions.is_deleted = 0 ) As USX_Tot_Students__c, " & _
                "(SELECT count(DISTINCT dbo.T_Presentation.applicationUid)     FROM   dbo.T_Presentation   WHERE  TU.id=dbo.T_Presentation.userId AND dbo.T_Presentation.isDeleted  =  0 AND  not(dbo.T_Presentation.FromStore = 1) AND  DATEDIFF(day, dbo.T_Presentation.created, GETDATE()) <= 7) As USX_Tot07_Created_Presentations__c,   " & _
                "(SELECT count(DISTINCT dbo.T_Presentation.applicationUid)     FROM   dbo.T_Presentation   WHERE  TU.id=dbo.T_Presentation.userId AND dbo.T_Presentation.isDeleted  =  0 AND  not(dbo.T_Presentation.FromStore = 1) AND  DATEDIFF(day, dbo.T_Presentation.created, GETDATE()) <= 30) As USX_Tot30_Created_Presentations__c,  " & _
                "(SELECT count(DISTINCT dbo.T_Presentation.applicationUid)     FROM   dbo.T_Presentation   WHERE  TU.id=dbo.T_Presentation.userId AND dbo.T_Presentation.isDeleted  =  0 AND  not(dbo.T_Presentation.FromStore = 1)) As USX_Tot_Created_Presentations__c," & _
                "(SELECT Count(distinct T_Sessions.session_uid) FROM dbo.T_Presentation, dbo.T_Sessions WHERE TU.id=dbo.T_Presentation.userId AND ( dbo.T_Presentation.applicationUid=dbo.T_Sessions.presentation_uid  ) AND  not(dbo.T_Presentation.FromStore = 1) and (dbo.T_Sessions.is_Deleted = 0) AND  DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 7) as USX_Tot07_Exec_Presentations__c, " & _
                "(SELECT Count(distinct T_Sessions.session_uid) FROM dbo.T_Presentation, dbo.T_Sessions WHERE TU.id=dbo.T_Presentation.userId AND ( dbo.T_Presentation.applicationUid=dbo.T_Sessions.presentation_uid  ) AND  not(dbo.T_Presentation.FromStore = 1) and (dbo.T_Sessions.is_Deleted = 0) AND  DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 30) as USX_Tot30_Exec_Presentations__c, " & _
                "(SELECT Count(distinct T_Sessions.session_uid) FROM dbo.T_Presentation, dbo.T_Sessions WHERE TU.id=dbo.T_Presentation.userId AND ( dbo.T_Presentation.applicationUid=dbo.T_Sessions.presentation_uid  ) AND  not(dbo.T_Presentation.FromStore = 1) and (dbo.T_Sessions.is_Deleted = 0)) as USX_Tot_Exec_Presentations__c," & _
                "(SELECT Count(distinct T_Sessions.session_uid) FROM dbo.T_Presentation, dbo.T_Sessions WHERE TU.id=dbo.T_Presentation.userId AND ( dbo.T_Presentation.applicationUid=dbo.T_Sessions.presentation_uid  ) AND  dbo.T_Presentation.FromStore = 1 and (dbo.T_Sessions.is_Deleted = 0) AND  DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 7) as USX_Tot07_Exec_Featured_Presentations__c, " & _
                "(SELECT Count(distinct T_Sessions.session_uid) FROM dbo.T_Presentation, dbo.T_Sessions WHERE TU.id=dbo.T_Presentation.userId AND ( dbo.T_Presentation.applicationUid=dbo.T_Sessions.presentation_uid  ) AND  dbo.T_Presentation.FromStore = 1 and (dbo.T_Sessions.is_Deleted = 0) AND  DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 30) as USX_Tot30_Exec_Featured_Presentations__c, " & _
                "(SELECT Count(distinct T_Sessions.session_uid) FROM dbo.T_Presentation, dbo.T_Sessions WHERE TU.id=dbo.T_Presentation.userId AND ( dbo.T_Presentation.applicationUid=dbo.T_Sessions.presentation_uid  ) AND  dbo.T_Presentation.FromStore = 1 and (dbo.T_Sessions.is_Deleted = 0)) as USX_Tot_Exec_Featured_Presentations__c, " & _
                "(select max(session_date) from dbo.T_Sessions  where T_Sessions.is_Deleted = 0 and TU.id = T_sessions.teacher_id) as Stats_last_activity__c, " & _
                "TU.[goldbyReferral] as goldbyReferral__C, " & _
                "(SELECT T2.sforceId from T_User T2 where TU.referredBy = T2.id) as Referred_By__C " & _
                "INTO TEMP_EXPORTSFDC " & _
                "FROM T_User TU, TSF_Contact " & _
                "WHERE NOT TU.sforceId IS NULL And not TU.sforceId = '' " & _
                "      AND TU.sforceId = TSF_Contact.ID "
            oCmd.ExecuteNonQuery()


            ' ------------------- SILVER --------------------
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed 7 Silver)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed07_Stud_Limit__c = USX_Exceed07_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'Silver') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "			AND DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 7 " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 30) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed 30 Silver)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed30_Stud_Limit__c = USX_Exceed30_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'Silver') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "			AND DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 30 " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 30) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed All Time Silver)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed_Stud_Limit__c = USX_Exceed_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'Silver') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 30) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ' ------------------- GOLD --------------------

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed 7 GOLD)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed07_Stud_Limit__c = USX_Exceed07_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'Nearpod Gold Edition') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "			AND DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 7 " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 50) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed 30 GOLD)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed30_Stud_Limit__c = USX_Exceed30_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'Nearpod Gold Edition') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "			AND DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 30 " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 50) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed All Time GOLD)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed_Stud_Limit__c = USX_Exceed_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'Nearpod Gold Edition') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 50) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ' ------------------- SCHOOL --------------------

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed 7 School)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed07_Stud_Limit__c = USX_Exceed07_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'School') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "			AND DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 7 " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 100) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed 30 School)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed30_Stud_Limit__c = USX_Exceed30_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'School') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "			AND DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 30 " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 100) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed All Time School)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed_Stud_Limit__c = USX_Exceed_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'School') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 100) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ' ------------------- Cart --------------------

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed 7 Cart)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed07_Stud_Limit__c = USX_Exceed07_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'Cart') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "			AND DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 7 " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 50) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed 30 Cart)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed30_Stud_Limit__c = USX_Exceed30_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'Cart') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "			AND DATEDIFF(day, dbo.T_Sessions.session_date, GETDATE()) <= 30 " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 50) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Updating Temp (Exceed All Time Cart)" : Application.DoEvents()
            oCmd.CommandText = _
                "update [dbo].[TEMP_EXPORTSFDC] " & _
                "set USX_Exceed_Stud_Limit__c = USX_Exceed_Stud_Limit__c + " & _
                "	(select count(SESSIONEXCEDIDAS.QStudents) from " & _
                "		(SELECT Count(Distinct dbo.T_Sessions.lead_uid) as QStudents " & _
                "		 FROM " & _
                "			dbo.T_user TU,  " & _
                "			dbo.T_Sessions " & _
                "		WHERE " & _
                "			(TU.sforceId = TEMP_EXPORTSFDC.ID) and " & _
                "			(TU.type = 'Cart') and " & _
                "			(TU.ID=dbo.T_Sessions.teacher_id  ) " & _
                "			AND (T_Sessions.is_deleted = 0) " & _
                "		GROUP BY " & _
                "			dbo.T_Sessions.session_uid " & _
                "		HAVING Count(Distinct dbo.T_Sessions.lead_uid) > 50) SESSIONEXCEDIDAS) "
            oCmd.ExecuteNonQuery()




            plblTable.Text = "Done"
            plblCurrentOp.Text = "Done"
            pPgbParcial.Value = 0
            pPgbGlobal.Value = pPgbGlobal.Maximum

            plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
            lvlngNumReg = gflng_GetNumReg(goConNear, "T_User", "NOT sforceId IS NULL")
            lvstrSql = _
                "SELECT * from TEMP_EXPORTSFDC"

            plblTable.Text = "C:\Datamart\TxOut\TUserStats.csv" : Application.DoEvents()
            sResulta += gfstr_BackupTable("C:\Datamart\TxOut\TUserStats.csv", lvstrSql, lvlngNumReg, plblCurrentOp, pPgbParcial)

            sResulta += "End all process" & Now.ToString & vbCrLf

        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Call GrabarLog(eLogType.eSFO, sResulta)

        Return sResulta


    End Function

End Class
