Public Class cTruncateAll

    Const C_LINKED As Boolean = True
    
    Public Shared Function Execute(Optional ByRef pexError As Exception = Nothing) As String

        Dim sResulta As String = ""
        Dim oCmd As SqlClient.SqlCommand



        Try    
            oCmd = goConNear.CreateCommand
            With oCmd

                sResulta += " Truncating T_Answer " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_Answer"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_District " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_District"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_PollAnswer " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_PollAnswer"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_PollQuestion " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_PollQuestion"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_Presentation " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_Presentation"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_Product " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_Product"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_QAAnswer " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_QAAnswer"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_QAQuestion " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_QAQuestion"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_Question " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_Question"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_QuizAnswer " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_QuizAnswer"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_QuizQuestion " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_QuizQuestion"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_School " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_School"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_Sessions " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_Sessions"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_SharePresentation " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_SharePresentation"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_Slide " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_Slide"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_SlideShow " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_SlideShow"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_User " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_User"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_UserHomeworks " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_UserHomeworks"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_UserWebJoins " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_UserWebJoins"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

                sResulta += " Truncating T_Webpage " & Now.ToString & vbCrLf

                .CommandText = "TRUNCATE TABLE T_Webpage"
                .CommandTimeout = 99999999
                .ExecuteNonQuery()

            End With
            sResulta += " and Ending at " & Now.ToString & vbCrLf
            Return sResulta
        Catch ex As Exception
            Execute = ex.ToString
            pexError = ex
            Call GlobalErrorHandler(ex, "modGlobales.gfstr_ImportaLinked")
        End Try

    End Function


End Class