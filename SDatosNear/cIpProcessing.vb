Imports System.Data.SqlClient

Public Class cIpProcessing

    Public Shared Function Generate(ByRef pPgbGlobal As ProgressBar, _
                                    PgbParcial As ProgressBar, _
                                    plblCurrentOp As Label, _
                                    plblTable As Label, _
                                    Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try

            Dim arrauserid As New ArrayList
            Dim arrazipcode As New ArrayList
            Dim arracountry As New ArrayList
            Dim arrastate As New ArrayList
            Dim arracity As New ArrayList

            Dim olduserid = 0
            Dim lCommand As String = ""

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Zip from IP" : Application.DoEvents()

            Dim Com As New SqlCommand("select userid, country, region, city, TD_states.STATE_2, TD_ZipCodes.ZC_Zip, count(*) " & _
                                        "from T_UserLogin " & _
                                        "left join TD_States on T_UserLogin.Region = TD_States.STATE_FULL " & _
                                        "left join TD_ZipCodes on TD_States.STATE_2 = TD_ZipCodes.ZC_State and T_UserLogin.city = TD_ZipCodes.ZC_PrimaryCity and ZC_default = 1 " & _
                                        "  " & _
                                        "group by userId, country, region, city,  TD_states.STATE_2 , TD_ZipCodes.ZC_Zip " & _
                                        "order by count(*) desc ", goConNear)
            Dim RDR = Com.ExecuteReader()
            If RDR.HasRows Then
                Do While RDR.Read

                    If (olduserid <> RDR.Item("userId")) Then


                        arrauserid.Add(RDR.Item("userId"))
                        If IsDBNull(RDR.Item("ZC_Zip")) Then
                            arrazipcode.Add(0)
                        Else
                            arrazipcode.Add(RDR.Item("ZC_Zip").replace("'", ""))
                        End If
                        If IsDBNull(RDR.Item("country")) Then
                            arracountry.Add("N/A")
                        Else
                            arracountry.Add(RDR.Item("country").replace("'", ""))
                        End If
                        If IsDBNull(RDR.Item("STATE_2")) Then
                            arrastate.Add("N/A")
                        Else
                            arrastate.Add(RDR.Item("STATE_2").replace("'", ""))
                        End If
                        If IsDBNull(RDR.Item("city")) Then
                            arracity.Add("N/A")
                        Else
                            arracity.Add(RDR.Item("city").replace("'", ""))
                        End If


                        olduserid = RDR.Item("userId")
                    End If

                Loop
            End If
            RDR.Close()


            For i = 0 To arrauserid.Count - 1
                lCommand = lCommand & "update T_User " & _
                                      " set fromip_zipcode = " & arrazipcode.Item(i).ToString() & _
                                      "    , fromip_city = '" & arracity.Item(i).ToString() & "'" & _
                                      "    , fromip_state = '" & arrastate.Item(i).ToString() & "'" & _
                                      "    , fromip_country = '" & arracountry.Item(i).ToString() & "'" & _
                                      " where id = " & arrauserid.Item(i).ToString() & "; "

                If (i / 20 = Int(i / 20) Or i = arrauserid.Count - 1) Then
                    oCmd = goConNear.CreateCommand
                    oCmd.CommandTimeout = 999999
                    oCmd.CommandText = lCommand
                    oCmd.ExecuteNonQuery()
                    lCommand = ""
                End If

            Next


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function
End Class
