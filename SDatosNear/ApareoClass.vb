Public Class ApareoClass
    Public Shared Function Compare( _
            ByRef pPgbGlobal As ProgressBar, _
            pPgbParcial As ProgressBar, _
            plblCurrentOp As Label, _
            plblTable As Label, _
            Optional ByRef pexError As Exception = Nothing) As String

        Dim sResulta As String = "Apareo SalesForce" & vbCrLf
        Dim lvstrSql As String
        Dim oCmd As SqlClient.SqlCommand

        Try
            sResulta += "Start all process" & Now.ToString & vbCrLf

            pPgbGlobal.Maximum = 100
            pPgbGlobal.Value = 0

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Deleting Temp Table" : Application.DoEvents()
            oCmd.CommandText = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TEMP_APAREOSFDC]') AND type in (N'U')) " & _
                              " DROP TABLE [dbo].[TEMP_APAREOSFDC]"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Create Temp Table" : Application.DoEvents()
            oCmd.CommandText = "CREATE TABLE [dbo].[TEMP_APAREOSFDC](" & _
                            "[Proceso] [varchar](2) NOT NULL, " & _
                            "[U_Id] [int] NOT NULL, " & _
                            "[U_sforceId] [nvarchar](30) NULL, " & _
                            "[Observ] [varchar](200) NOT NULL " & _
                               ") ON [PRIMARY]"
            oCmd.ExecuteNonQuery()


            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Usuarios sin SFDC ID" : Application.DoEvents()
            oCmd.CommandText = "Insert into TEMP_APAREOSFDC " & _
                               "Select  'P1' as Proceso, " & _
                               "       T_User.id as U_Id, " & _
                               "       T_User.sforceId as U_sforceId, " & _
                               "       'SforceID en cero' as Observ " & _
                               "from dbo.t_user left join dbo.TSF_Contact on t_user.sforceId = TSF_Contact.ID " & _
                               "   where (t_user.isDeleted = 0)" & _
                               "       and (t_user.sforceId = '') "
            oCmd.ExecuteNonQuery()


            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Diff en First/Last/Email" : Application.DoEvents()
            oCmd.CommandText = "Insert into TEMP_APAREOSFDC " & _
                                "Select    'P2' as Proceso, " & _
                                "           T_User.id as U_Id, " & _
                                "           T_User.sforceId as U_sforceId, " & _
                                "           'First/LastName/Email with Differences' as Observ " & _
                                "from dbo.t_user left join dbo.TSF_Contact on t_user.sforceId = TSF_Contact.ID " & _
                                "where (t_user.isDeleted = 0)" & _
                                "   and (" & _
                                "           (t_user.firstName != TSF_Contact.FIRSTNAME) " & _
                                "       or " & _
                                "           (t_user.lastName != TSF_Contact.LASTNAME)" & _
                                "       or " & _
                                "           (t_user.email != TSF_Contact.EMAIL)" & _
                                "        ) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Diff en User Type" : Application.DoEvents()
            oCmd.CommandText = "Insert into TEMP_APAREOSFDC " & _
                                "Select    'P3' as Proceso, " & _
                                "           T_User.id as U_Id, " & _
                                "           T_User.sforceId as U_sforceId, " & _
                                "           'User Type Diferente User: ' + t_user.sfType +' En SFDC: ' + TSF_Contact.USERTYPE__C " & _
                                "from dbo.t_user left join dbo.TSF_Contact on t_user.sforceId = TSF_Contact.ID " & _
                                "where (t_user.isDeleted = 0)" & _
                                "   and (t_user.sfType != TSF_Contact.USERTYPE__C) "
            oCmd.ExecuteNonQuery()


                'plblTable.Text = "C:\Datamart\TxOut\TUserStats.csv" : Application.DoEvents()
                'sResulta += gfstr_BackupTable("C:\Datamart\TxOut\TUserStats.csv", lvstrSql, lvlngNumReg, plblCurrentOp, pPgbParcial)

            sResulta += "End all process" & Now.ToString & vbCrLf

        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Call GrabarLog(eLogType.eSFO, sResulta)

        Return sResulta


    End Function

End Class
