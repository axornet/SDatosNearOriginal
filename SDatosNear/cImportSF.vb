Imports System.IO

Public Class cImportSF
    Public Enum eImportMetod
        eBulk = 1
        eStandard = 2
    End Enum

    Public Class cSalesForceItem
        Public strLabel As String
        Public strPath As String
        Public strTable As String
        Public strSeparator As String
        Public strMetod As eImportMetod = eImportMetod.eBulk
        Public strDecimalPoint As String

        Public Sub New(pLabel As String, pPath As String, pTable As String, pMetod As eImportMetod, Optional pSeparator As String = Chr(34) & "," & Chr(34), Optional pDecimalPoint As String = Nothing)
            Me.strLabel = pLabel
            Me.strPath = pPath
            Me.strTable = pTable
            Me.strSeparator = pSeparator
            Me.strMetod = pMetod
            Me.strDecimalPoint = pDecimalPoint
        End Sub

        Public Overrides Function ToString() As String
            Return strLabel
        End Function

    End Class

    Public Shared Function GetTablesCollection() As List(Of cSalesForceItem)
        Dim oLista As New List(Of cSalesForceItem)
        oLista.Add(New cSalesForceItem("TSF_Contact", "C:\DataSFDC\Contact.csv", "TSF_Contact", eImportMetod.eBulk, ","))
        oLista.Add(New cSalesForceItem("TSF_Leads", "C:\DataSFDC\Lead.csv", "TSF_Leads", eImportMetod.eBulk, ","))
        oLista.Add(New cSalesForceItem("T_NPStore", "C:\DataSFDC\npstore.csv", "T_NPStore", eImportMetod.eBulk, ","))
        oLista.Add(New cSalesForceItem("TSF_Opportunity", "C:\DataSFDC\opportinity.csv", "TSF_Opportunity", eImportMetod.eBulk, ",", "."))
        oLista.Add(New cSalesForceItem("TSF_District", "C:\DataSFDC\District.csv", "TSF_District", eImportMetod.eBulk, ","))
        oLista.Add(New cSalesForceItem("TSF_School", "C:\DataSFDC\school.csv", "TSF_School", eImportMetod.eBulk, ",", "."))
        'No van Mas oLista.Add(New cSalesForceItem("TSF_PayPalTxn", "C:\DataSFDC\PayPalTxn.csv", "TSF_PayPalTxn", eImportMetod.eBulk, ","))
        'No van Mas oLista.Add(New cSalesForceItem("TSF_StatsApple", "C:\DataSFDC\StatsApple.csv", "TSF_StatsApple", eImportMetod.eBulk, ","))
        'No van Mas oLista.Add(New cSalesForceItem("TSF_StatsGoogle", "C:\DataSFDC\StatsGoogle.csv", "TSF_StatsGoogle", eImportMetod.eBulk, ","))
        oLista.Add(New cSalesForceItem("TSF_DnsDomains", "C:\DataSFDC\DnsDomains.csv", "TSF_DnsDomains", eImportMetod.eBulk, ","))
        oLista.Add(New cSalesForceItem("TSF_Webinear", "C:\DataSFDC\Webinear.csv", "TSF_Webinear", eImportMetod.eBulk, ","))
        oLista.Add(New cSalesForceItem("TSF_Account", "C:\DataSFDC\Account.csv", "TSF_Account", eImportMetod.eBulk, ","))
        Return oLista
    End Function


    Public Shared Function Import( _
            pPgbGlobal As ProgressBar, pPgbParcial As ProgressBar, _
            plblCurrentOp As Label, plblTable As Label, _
            Optional ByRef pTablesCollection As List(Of cSalesForceItem) = Nothing,
            Optional ByRef pexError As Exception = Nothing) As String
        Dim sResulta As String = "Import SalesForce" & vbCrLf
        Dim lTablesCollection As List(Of cSalesForceItem)

        Try
            sResulta += "Start " & Now.ToString & vbCrLf
            If pTablesCollection Is Nothing Then
                lTablesCollection = cImportSF.GetTablesCollection
            Else
                lTablesCollection = pTablesCollection
            End If

            pPgbGlobal.Maximum = 15 + lTablesCollection.Count
            pPgbGlobal.Value = 0

            For Each oItem As cSalesForceItem In lTablesCollection
                ProgressBarAdd(pPgbGlobal)
                plblTable.Text = oItem.strLabel
                If oItem.strMetod = eImportMetod.eBulk Then
                    sResulta += gfstr_ImportBulkFromTxt(oItem.strPath, oItem.strTable, True, pexError, plblCurrentOp, pPgbParcial, oItem.strSeparator, True, , oItem.strDecimalPoint)
                Else
                    sResulta += gfstr_ImportFromTxt(oItem.strPath, oItem.strTable, True, pexError, plblCurrentOp, pPgbParcial, oItem.strSeparator)
                End If
            Next

            Dim oCmd As SqlClient.SqlCommand

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Copy Biz Premium a Gold Date ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "update tsf_contact set Gold_Upgraded__c = Biz_Premium_Upgraded__c " & _
                "where Biz_Premium_Upgraded__c is not null and (Gold_Upgraded__c is null or Gold_Upgraded__c is not null and  Biz_Premium_Upgraded__c > Gold_Upgraded__c)"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Copy Biz Pro a Gold Date ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "update tsf_contact set Gold_Upgraded__c = Biz_Pro_Upgraded__c " & _
                 "where Biz_Pro_Upgraded__c is not null and (Gold_Upgraded__c is null or Gold_Upgraded__c is not null and  Biz_Pro_Upgraded__c > Gold_Upgraded__c)"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Copy Biz Team a Gold Date ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "update tsf_contact set Gold_Upgraded__c = Biz_Team_Upgraded__c " & _
                 "where Biz_Team_Upgraded__c is not null and (Gold_Upgraded__c is null or Gold_Upgraded__c is not null and  Biz_Team_Upgraded__c > Gold_Upgraded__c)"
            oCmd.ExecuteNonQuery()


            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates Contact Gold...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("TSF_Contact", "GOLD_UPGRADED__C", "_GoldUp")


            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates Vip Asigned Date...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("TSF_Contact", "VIP_OWNER_ASSIGNED_DATE__C", "_VIPAD")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit CountryName Contact...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritCountry("TSF_Contact", "MAILINGCOUNTRY", "_MAILINGCOUNTRY")


            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Fixing UNITED STATES ZIP CODES" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "update tsf_contact set zc_zip = cast(SUBSTRING(MAILINGPOSTALCODE,1,5) as int) " & _
                "where COUNTRYNAME_MAILINGCOUNTRY = 'UNITED STATES' and  isnumeric(SUBSTRING(MAILINGPOSTALCODE,1,5)) = 1 "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Update Geo info from United States " : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                         "UPDATE tsf_contact " & _
                         "SET " & _
                         "   tsf_contact.[ZC_Type] = z.[ZC_Type], " & _
                         "   tsf_contact.[ZC_PrimaryCity] =z.[ZC_PrimaryCity], " & _
                         "   tsf_contact.[ZC_AcceptableCities] = z.[ZC_AcceptableCities], " & _
                         "   tsf_contact.[ZC_UnAcceptableCities] = z.[ZC_UnAcceptableCities]," & _
                         "   tsf_contact.[ZC_State] = z.[ZC_State]," & _
                         "   tsf_contact.[ZC_County]= z.[ZC_County]," & _
                         "   tsf_contact.[ZC_TimeZone]= z.[ZC_TimeZone]," & _
                         "   tsf_contact.[ZC_AreaCodes]= z.[ZC_AreaCodes]," & _
                         "   tsf_contact.[ZC_Latitude]= z.[ZC_Latitude]," & _
                         "   tsf_contact.[ZC_Longitude]= z.[ZC_Longitude]," & _
                         "   tsf_contact.[ZC_WorldRegion]= Z.[ZC_WorldRegion]," & _
                         "   tsf_contact.[ZC_Country]= z.[ZC_Country]," & _
                         "   tsf_contact.[ZC_Decommissioned]= z.[ZC_Decommissioned]," & _
                         "tsf_contact.[ZC_Estimated_population] = z.[ZC_Estimated_population] " & _
                         "FROM tsf_contact " & _
                         "JOIN " & _
                         "   TD_ZipCodes as z ON tsf_contact.zc_zip = z.ZC_Zip"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Other CountryName Contact...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritCountry("TSF_Contact", "OTHERCOUNTRY", "_OTHERCOUNTRY")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates Opp Create Date...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("TSF_Opportunity", "CREATEDDATE", "CR")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates Opp Close Date ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("TSF_Opportunity", "CLOSEDATE", "CLOSE")

            'Schools 
            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates School Initial Contract Date ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("TSF_School", "SCHOOL_CONTRACT_INITIAL_DATE__C", "1CID")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates School 2nd Initial Contract Date ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("TSF_School", "SCHOOL_2ND_CONTRACT_INITIAL_DATE__C", "2CID")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates School End Contract Date ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("TSF_School", "SCHOOL_CONTRACT_END_DATE__C", "1CED")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates School 2nd End Contract Date ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("TSF_School", "SCHOOL_2ND_CONTRACT_END_DATE__C", "2CED")


            plblTable.Text = "Done"
            plblCurrentOp.Text = "Done"
            pPgbParcial.Value = 0
            pPgbGlobal.Value = pPgbGlobal.Maximum

            sResulta += "End " & Now.ToString & vbCrLf

        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf
        End Try
        Call GrabarLog(eLogType.eSF, sResulta)

        Return sResulta

    End Function

End Class
