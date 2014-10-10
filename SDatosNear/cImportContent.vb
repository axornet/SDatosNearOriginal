Public Class cImportContentGeneric
    Const C_LINKED As Boolean = True
    'Const C_DATABASE As String = "[CONTENT]"
    Const C_DeltaNumINI As String = "1000000000"
    Const C_DeltaNumEND As String = "2000000000"
    Const C_DeltaText As String = """B_"""

    Public Shared Function Import( _
            pDataBase As String, _
            pPgbGlobal As ProgressBar, pPgbParcial As ProgressBar, _
            plblTable As Label, plblCurrentOp As Label, _
            pBizSystem As Boolean, _
            Optional ByRef pexError As Exception = Nothing) As String
        Dim lvstrExpSql As String
        Dim lvstrColumns As String
        Dim sResulta As String
        Dim sSpeacialSqlDelete As String

        If (pBizSystem) Then
            sResulta = "Import Biz Content" & vbCrLf
        Else
            sResulta = "Import Content" & vbCrLf
        End If

        Try
            sResulta += "Start " & Now.ToString & vbCrLf

            pPgbGlobal.Maximum = 72
            pPgbGlobal.Value = 0

            'GoTo ZonaUpdates

            If Not (pBizSystem) Then
                'Importa de NearAdmin los schools 
                lvstrExpSql = _
                  "SELECT " & _
                  FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                  IIf(pBizSystem, "1", "0") & " as sd_source, " & _
                  "schoolNCESId, districtNCESId, lowGrade, HighGrade, name, " & _
                  "countyname, address, city, state, zipcode, phone, localeCode, locale, " & _
                  "charter, magnet, title1School, title1SchoolWide, enrollment, teachers, schools, " & _
                  "studentsTeachersRatio, freeLunch, reducedLunch, type, gsId, fax, website, " & _
                  "domainfromwebsite, lat, lon " & _
                  "FROM nearpodschoolinfo order by id " & GC_LIMITRESULT
                lvstrColumns = _
                    "id, sd_source, schoolNCESId, districtNCESId, lowGrade, HighGrade, name, " & _
                    "countyname, address, city, state, zipcode, phone, localeCode, locale, " & _
                    "charter, magnet, title1School, title1SchoolWide, enrollment, teachers, schools, " & _
                    "studentsTeachersRatio, freeLunch, reducedLunch, type, gsId, fax, website, " & _
                    "domainfromwebsite, lat, lon "
                If C_LINKED Then
                    sSpeacialSqlDelete = "DELETE FROM T_NearpodSchoolInfo where sd_source = 0"
                    sResulta += gfstr_ImportaBulked(goConNear, goConnNearAdmin, lvstrExpSql, "T_NearpodSchoolInfo", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)

                Else

                    pPgbParcial.Minimum = 0
                    plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                    pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM nearpodschoolinfo ")
                    sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_NearpodSchoolInfo", True, pPgbParcial, plblCurrentOp, plblTable, pexError)

                End If

            
                lvstrExpSql = _
                    "SELECT " & _
                    FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                    IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                    FieldAddNumber(pBizSystem, "userid ") & " as userid, " & _
                    "from_unixtime(created) as created, " & _
                    "ip,  country, region, city, full_dump, tries " & _
                    "FROM userlogin order by id " & GC_LIMITRESULT
                lvstrColumns = _
                    "id, sd_source, created, ip,  country, region, city, full_dump, tries "
                If C_LINKED Then
                    sSpeacialSqlDelete = "DELETE FROM T_UserLogin where sd_source = 0"
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_UserLogin", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    pPgbParcial.Minimum = 0
                    plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                    pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM userlogin ")
                    sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_UserLogin", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
                End If
            End If


            lvstrExpSql = _
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                FieldAddNumber(pBizSystem, "questionId ") & " as questionId, " & _
                "answerText, orden, " & _
                "manualAnswer,manualAnswerText,multilineAnswer," & _
                "isDeleted FROM Answer order by id " & GC_LIMITRESULT
            lvstrColumns = _
                "id, sd_source, questionId,answerText,orden," & _
                "manualAnswer,manualAnswerText,multilineAnswer," & _
                "isDeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Answer where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Answer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Answer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Answer where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Answer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Answer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else

                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Answer ")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Answer", True, pPgbParcial, plblCurrentOp, plblTable, pexError)

            End If


            ProgressBarAdd(pPgbGlobal)
            'lvstrExpSql = "SELECT id,isDeleted,name,maxAdmins,maxUsers,maxSchools,SUBSTR(sforceId,1,15) as sforceId FROM District order by id"
            lvstrExpSql =
                "SELECT " & _
                 FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "isDeleted,name,maxAdmins,maxUsers,maxSchools,sforceId FROM District order by id"
            lvstrColumns = "id,sd_source, isDeleted,name,maxAdmins,maxUsers,maxSchools,sforceId"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_District where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_District", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_District", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_District where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_District", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_District", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM District")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_District", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "name, isdeleted " & _
                "FROM Age order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, name, isdeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Age where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Age", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Age", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Age where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Age", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Age", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Age")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Age", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "name, isdeleted " & _
                "FROM Author order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, name, isdeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Author where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Author", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Author", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Author where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Author", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Author", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Author")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Author", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "presentationId, gradeId, isDeleted " & _
                "FROM PresentationGrades order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, presentationId, gradeId, isDeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                   
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_PresentationGrades where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Grade", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_PresentationGrades", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM PresentationGrades")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_PresentationGrades", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If


            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "name, isdeleted " & _
                "FROM Grade order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, name, isdeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Grade where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Grade", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Grade", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Grade where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Grade", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Grade", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Grade")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Grade", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "name, isdeleted " & _
                "FROM Initiative order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, name, isdeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Initiative where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Initiative", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Initiative", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Initiative where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Initiative", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Initiative", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Initiative")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Initiative", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "name, isdeleted " & _
                "FROM Level order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, name, isdeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Level where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Level", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Level", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Level where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Level", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Level", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Level")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Level", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "name, isdeleted " & _
                "FROM Publisher order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, name, isdeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Publisher where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Publisher", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Publisher", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Publisher where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Publisher", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Publisher", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Publisher")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Publisher", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "name, isdeleted " & _
                "FROM Subject order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, name, isdeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Subject where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Subject", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Subject", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Subject where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Subject", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Subject", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Subject")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Subject", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                FieldAddNumber(pBizSystem, "answerId ") & " as answerId " & _
                "FROM PollAnswer order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, answerId"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_PollAnswer where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_PollAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_PollAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_PollAnswer where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_PollAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_PollAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM PollAnswer")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_PollAnswer", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                FieldAddNumber(pBizSystem, "questionid ") & " as questionid " & _
                "FROM PollQuestion order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, questionid"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_PollQuestion where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_PollQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_PollQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_PollQuestion where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_PollQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_PollQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM PollQuestion")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_PollQuestion", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql = _
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "trim(cast(name as char(150))) as name, " & _
                "from_unixtime(created) as created, " & _
                FieldAddNumber(pBizSystem, "userId ") & " as userId, " & _
                "published, " & _
                "IFNULL(" & FieldAddText(pBizSystem, "applicationUid") & ",'')" & " as applicationUid, " & _
                "from_unixtime(modified) As modified," & _
                "isDeleted,featured,size, " & _
                FieldAddNumber(pBizSystem, "parentId") & " as parentId, " & _
                "publishers," & _
                "skinId, language, productStoreId, price, " & _
                FieldAddNumber(pBizSystem, "publisherId ") & " as publisherId, " & _
                FieldAddNumber(pBizSystem, "initiativeId ") & " as initiativeId, " & _
                FieldAddNumber(pBizSystem, "levelId ") & " as levelId, " & _
                FieldAddNumber(pBizSystem, "ageId ") & " as ageId, " & _
                FieldAddNumber(pBizSystem, "subjectId ") & " as subjectId, " & _
                FieldAddNumber(pBizSystem, "authorId ") & " as authorId, " & _
                " archived " & _
                " FROM Presentation order by id " & GC_LIMITRESULT
            lvstrColumns = _
                "id, sd_source, name,created,userId," & _
                "published, applicationUid," & _
                "modified," & _
                "isDeleted,featured,size,parentId,publishers," & _
                "skinId,language,productStoreId,price, " & _
                "publisherId,initiativeId,gradeId,levelId,ageId,subjectId,authorId, archived "
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Presentation where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Presentation", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Presentation", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Presentation where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Presentation", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Presentation", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Presentation")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Presentation", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql = _
               "SELECT " & _
               FieldAddNumber(pBizSystem, "id ") & " as id, " & _
               IIf(pBizSystem, "1", "0") & " as sd_source ," & _
               "name, price, isDeleted, " & _
               "maxPresentations,maxSizeOfPresentations," & _
               "maxStorage,maxStudents,maxStudentsForTrial," & _
               "shareWithUsers,extendedFeatures,sfType,watermark FROM Product order by id " & GC_LIMITRESULT
            lvstrColumns = _
                "id, sd_source, name,price,isDeleted," & _
                "maxPresentations,maxSizeOfPresentations," & _
                "maxStorage,maxStudents,maxStudentsForTrial," & _
                "shareWithUsers,extendedFeatures,sfType,watermark "
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Product where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Product", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Product", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Product where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Product", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Product", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If

            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Product")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Product", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            'ProgressBarAdd(pPgbGlobal)
            'lvstrExpSql = _
            '   "SELECT " & _
            '   FieldAddNumber(pBizSystem, "id ") & ", " & _
            '   IIf(pBizSystem, "1", "0") & " as sd_source ," & _
            '   FieldAddNumber(pBizSystem, "productId ") & ", " & _
            '   "price, regularity, unity " & _
            '   "FROM ProductPrice order by id " & GC_LIMITRESULT
            'lvstrColumns = _
            '    "id, sd_source, productId, price, regularity, unity"
            'If C_LINKED Then
            '    If (pBizSystem) Then
            '        sSpeacialSqlDelete = "DELETE FROM T_ProductPrice where sd_source = 1"
            '        sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_ProductPrice", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
            '    Else
            '        sSpeacialSqlDelete = "DELETE FROM T_ProductPrice where sd_source = 0"
            '        sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_ProductPrice", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
            '    End If
            'Else
            '    pPgbParcial.Minimum = 0
            '    plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
            '    pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM ProductPrice")
            '    sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_ProductPrice", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            'End If


            lvstrExpSql =
              "SELECT " & _
              FieldAddNumber(pBizSystem, "id ") & " as id, " & _
               IIf(pBizSystem, "1", "0") & " as sd_source ," & _
              FieldAddNumber(pBizSystem, "answerId ") & " as answerId, " & _
               "isCorrect FROM QAAnswer Order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, answerId, isCorrect"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_QAAnswer where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_QAAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_QAAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_QAAnswer where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_QAAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_QAAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM QAAnswer")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_QAAnswer", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If


            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
              "SELECT " & _
              FieldAddNumber(pBizSystem, "id ") & " as id, " & _
               IIf(pBizSystem, "1", "0") & " as sd_source ," & _
              FieldAddNumber(pBizSystem, "questionId ") & " as questionId " & _
              "FROM QAQuestion  Order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, questionId"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_QAQuestion where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_QAQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_QAQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_QAQuestion where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_QAQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_QAQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If

            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM QAQuestion")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_QAQuestion", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                FieldAddNumber(pBizSystem, "slideId ") & " as slideId, " & _
                "questionText, orden, type, onlyOneAnswer," & _
                "onlyOneAnswerText,isDeleted " & _
                "FROM Question  Order by id  " & GC_LIMITRESULT
            lvstrColumns = _
                "id, sd_source, slideId,questionText,orden,type,onlyOneAnswer," & _
                "onlyOneAnswerText,isDeleted "
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Question where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Question", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Question", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Question where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Question", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Question", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Question")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Question", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                FieldAddNumber(pBizSystem, "answerId ") & " as answerId, " & _
                "value FROM QuizAnswer Order by id  " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, answerId,value"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_QuizAnswer where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_QuizAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_QuizAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_QuizAnswer where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_QuizAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_QuizAnswer", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM QuizAnswer")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_QuizAnswer", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
               "SELECT " & _
               FieldAddNumber(pBizSystem, "id ") & " as id, " & _
               IIf(pBizSystem, "1", "0") & " as sd_source ," & _
               FieldAddNumber(pBizSystem, "questionId ") & " as questionId, " & _
               "value FROM QuizQuestion  Order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, questionId,value"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_QuizQuestion where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_QuizQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_QuizQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_QuizQuestion where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_QuizQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_QuizQuestion", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM QuizQuestion")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_QuizQuestion", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                FieldAddText(pBizSystem, "uid") & " as uid, " & _
                "isDeleted,name,maxAdmins,maxUsers,sforceId, " & _
                FieldAddNumber(pBizSystem, "districtId ") & " as districtId " & _
                "FROM School  Order by id "
            lvstrColumns = "id, sd_source, uid,isDeleted,name,maxAdmins,maxUsers,sforceId,districtId"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_School where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_School", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_School", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_School where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_School", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_School", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM School ")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_School", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "code, " & _
                FieldAddNumber(pBizSystem, "presentationId ") & " as presentationId, " & _
                "from_unixtime(created) as created FROM SharePresentation  Order by id  " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, code, presentationId,created"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_SharePresentation where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_SharePresentation", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_SharePresentation", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_SharePresentation where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_SharePresentation", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_SharePresentation", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM SharePresentation")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_SharePresentation", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                 "SELECT " & _
                 FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                 FieldAddNumber(pBizSystem, "presentationId ") & " as presentationId, " & _
                 "orden, " & _
                 "slideType, " & _
                 "title,resized,originalOrden, " & _
                 "presentationFileId, " &
                 "isDeleted, size " & _
                 "FROM Slide order by id " & GC_LIMITRESULT
            lvstrColumns = _
                "id, sd_source, presentationId, orden, slideType, " & _
                "title,resized,originalOrden, " & _
                "presentationFileId,isDeleted,size "
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Slide where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Slide", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Slide", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Slide where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Slide", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Slide", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Slide")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Slide", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)

            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                FieldAddNumber(pBizSystem, "slideId ") & " as slideId, " & _
                "orden,title,iconExtension,icon,isDeleted FROM SlideShow  Order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, slideId,orden,title,iconExtension,icon,isDeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_SlideShow where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_SlideShow", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_SlideShow", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_SlideShow where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_SlideShow", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_SlideShow", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM SlideShow")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_SlideShow", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
               "SELECT " & _
               FieldAddNumber(pBizSystem, "id ") & " as id , " & _
               IIf(pBizSystem, "1", "0") & " as sd_source ," & _
               FieldAddNumber(pBizSystem, "userId ") & " as userId, " & _
                FieldAddText(pBizSystem, "sessionUid") & " as sessionUid, " & _
               "isDeleted FROM UserHomeworks  Order by id " & GC_LIMITRESULT
            lvstrColumns = " id, sd_source, userId,sessionUid,isDeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_UserHomeworks where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_UserHomeworks", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_UserHomeworks", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_UserHomeworks where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_UserHomeworks", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_UserHomeworks", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM UserHomeworks")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_UserHomeworks", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
               "SELECT " & _
               FieldAddNumber(pBizSystem, "id ") & " as id, " & _
               IIf(pBizSystem, "1", "0") & " as sd_source ," & _
               FieldAddNumber(pBizSystem, "userId ") & " as userId, " & _
               FieldAddText(pBizSystem, "sessionUid") & " as sessionUid, " & _
               "isDeleted FROM UserWebJoins  Order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, userId,sessionUid,isDeleted"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_UserWebJoins where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_UserWebJoins", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_UserWebJoins", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_UserWebJoins where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_UserWebJoins", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_UserWebJoins", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM UserWebJoins")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_UserWebJoins", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
               "SELECT " & _
               FieldAddNumber(pBizSystem, "id ") & " as id, " & _
               IIf(pBizSystem, "1", "0") & " as sd_source ," & _
               FieldAddNumber(pBizSystem, "slideId ") & " as slideId, " & _
               "slideUrl,allowedUrls,allowToBrowse,browseAnyPage,isDeleted, type FROM Webpage  Order by id " & GC_LIMITRESULT
            lvstrColumns = "id, sd_source, slideId,slideUrl,allowedUrls,allowToBrowse,browseAnyPage,isDeleted, type"
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Webpage where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Webpage", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Webpage", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Webpage where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Webpage", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Webpage", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Webpage")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Webpage", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If


            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                 "SELECT " & _
                 FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                 IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                 "userName,email,firstName,lastName,from_unixtime(lastLogin) as lastLogin, " & _
                 "from_unixtime(created) as created,active,admin,lang,isDeleted," & _
                 FieldAddText(pBizSystem, "presenterUid") & " as presenterUid , " & _
                 "institute, qtyReportViews, " & _
                 "qtyLogins,storage," & _
                 "from_unixtime(lastSforceSync) as lastSforceSync,sforceId," & _
                 "qtyDashboardLogins,qtyBuilderLogins,maxPresentations,maxSizeOfPresentations," & _
                 "roleId, " & _
                 "maxStorage, maxStudents, maxStudentsForTrial, maxStudentsExceeded, " & _
                 "extendedFeaturesCount,shareWithUsers,extendedFeatures," & _
                 "type, " & _
                 FieldAddNumber(pBizSystem, "productId ") & " as productId, " & _
                 FieldAddNumber(pBizSystem, "schoolId ") & " as schoolId, " & _
                 FieldAddNumber(pBizSystem, "districtId ") & " as districtId, " & _
                 "sfType,referral,dismissPublishMessage," & _
                 "dismissGLogout,watermark,from_unixtime(nextFormCall) as nextFormCall,lastForm, archived, registeredFrom, mailAnnouncements, " & _
                 "promoCode, " & _
                 "from_unixtime(referralReminder) as referralReminder , " & _
                 "referralReminderCount, " & _
                 "skinId, " & _
                 "mailHomeworkActivities, " & _
                 "from_unixtime(expirationDate) as expirationDate, " & _
                 "productBuyBy, " & _
                 "referralOneLeftToGoldCount, " & _
                 "from_unixtime(referralOneLeftToGold) as referralOneLeftToGold, " & _
                 "downgradeMessageShowed, " & _
                 "hasShowSmartUsageEdit, " & _
                 "hasShowSmartUsage, " & _
                 "mailWeekSummaryActivities " & _
                 "FROM User  Order by id " & GC_LIMITRESULT
            lvstrColumns = _
                "id, sd_source, userName,email,firstName,lastName,lastLogin, " & _
                "created,active,admin,lang,isDeleted," & _
                "presenterUid,institute,qtyReportViews," & _
                "qtyLogins,storage," & _
                "lastSforceSync,sforceId," & _
                "qtyDashboardLogins,qtyBuilderLogins,maxPresentations,maxSizeOfPresentations," & _
                "roleId,maxStorage,maxStudents,maxStudentsForTrial,maxStudentsExceeded," & _
                "extendedFeaturesCount,shareWithUsers,extendedFeatures," & _
                "type,productId,schoolId,districtId,sfType,referral,dismissPublishMessage," & _
                "dismissGLogout,watermark,nextFormCall,lastForm, archived, registeredFrom, mailAnnouncements, " & _
                "promoCode, " & _
                "referralReminder , " & _
                "referralReminderCount, " & _
                "skinId, " & _
                "mailHomeworkActivities, " & _
                "expirationDate, " & _
                "productBuyBy, " & _
                "referralOneLeftToGoldCount, " & _
                "referralOneLeftToGold, " & _
                "downgradeMessageShowed, " & _
                "hasShowSmartUsageEdit, " & _
                "hasShowSmartUsage, " & _
                "mailWeekSummaryActivities "
            ' Le saque el If C_LINKED And False Then
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_User where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_User", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_User", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_User where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_User", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_User", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM User")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_User", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If


            'UserProductHistoric
            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                FieldAddNumber(pBizSystem, "userId ") & " as userId, " & _
                "from_unixtime(upgradeTime) as upgradeTime, " & _
                FieldAddNumber(pBizSystem, "oldProductId ") & " as oldProductId, " & _
                FieldAddNumber(pBizSystem, "productId ") & " as productId, " & _
                "productName, " & _
                "price, " & _
                "isDeleted, " & _
                "maxPresentations, " & _
                "maxSizeOfPresentations, " & _
                "maxStorage, " & _
                "maxStudents, " & _
                "maxStudentsForTrial, " & _
                "shareWithUsers, " & _
                "extendedFeatures, " & _
                "sfType, " & _
                "watermark, " & _
                "maxHomeworkJoins, " & _
                "source, " & _
                "upgradeAuthorizationManager, " & _
                "upgradeAuthorizationUser, " & _
                "upgradeAuthorizationMonths, " & _
                "from_unixtime(expirationDate) as expirationDate, " & _
                FieldAddNumber(pBizSystem, "upgradeAuthorizationUserId ") & " as upgradeAuthorizationUserId, " & _
                FieldAddNumber(pBizSystem, "sourceId ") & " as sourceId, " & _
                "regularity, " & _
                "unity " & _
                 "FROM UserProductHistoric  Order by id " & GC_LIMITRESULT
            lvstrColumns = _
                "id, " & _
                "sd_source, " & _
                "userId, " & _
                "upgradeTime, " & _
                "oldProductId, " & _
                "productId, " & _
                "productName, " & _
                "price, " & _
                "isDeleted, " & _
                "maxPresentations, " & _
                "maxSizeOfPresentations, " & _
                "maxStorage, " & _
                "maxStudents, " & _
                "maxStudentsForTrial, " & _
                "shareWithUsers, " & _
                "extendedFeatures, " & _
                "sfType, " & _
                "watermark, " & _
                "maxHomeworkJoins, " & _
                "source, " & _
                "upgradeAuthorizationManager, " & _
                "upgradeAuthorizationUser, " & _
                "upgradeAuthorizationMonths, " & _
                "expirationDate, " & _
                "upgradeAuthorizationUserId, " & _
                "sourceId, " & _
                "regularity, " & _
                "unity "
            ' Le saque el If C_LINKED And False Then
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_UserProductHistoric where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_UserProductHistoric", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_UserProductHistoric", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_UserProductHistoric where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_UserProductHistoric", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_UserProductHistoric", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM UserProductHistoric")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_UserProductHistoric", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            'UserTypeInfo
            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                FieldAddNumber(pBizSystem, "userId ") & " as userId, " & _
                FieldAddNumber(pBizSystem, "type ") & " as type, " & _
                "status, " & _
                "isMagic, " & _
                "createdAndPublished, " & _
                "downloadedFeatured, " & _
                "homework, " & _
                "reportsViewedQty, " & _
                "launchedWithJoins, " & _
                "sessionsQty, " & _
                "studentsJoinedQty, " & _
                "from_unixtime(modified) as modified, " & _
                "isDeleted, " & _
                "sessionsQtyLastPeriod, " & _
                "studentsJoinedQtyLastPeriod " & _
                "FROM usertypesinfo Order by id " & GC_LIMITRESULT
            lvstrColumns = _
                "id " & _
                "sd_source, " & _
                "userId " & _
                "type " & _
                "status, " & _
                "isMagic, " & _
                "createdAndPublished, " & _
                "downloadedFeatured, " & _
                "homework, " & _
                "reportsViewedQty, " & _
                "launchedWithJoins, " & _
                "sessionsQty, " & _
                "studentsJoinedQty, " & _
                "modified, " & _
                "isDeleted, " & _
                "sessionsQtyLastPeriod, " & _
                "studentsJoinedQtyLastPeriod "
            ' Le saque el If C_LINKED And False Then
            If C_LINKED Then
                If (pBizSystem) Then
                    'sSpeacialSqlDelete = "DELETE FROM T_UserTypesInfo where sd_source = 1"
                    'sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_UserTypesInfo", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_UserTypesInfo where sd_source = 0"
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_UserTypesInfo", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM usertypesinfo")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_UserTypesInfo", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If


            'AppleReceipt
            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "from_unixtime(created) as created, " & _
                FieldAddNumber(pBizSystem, "userId ") & " as userId, " & _
                FieldAddNumber(pBizSystem, "productId ") & " as productId, " & _
                FieldAddNumber(pBizSystem, "presentationId ") & " as presentationId, " & _
                "product_id, " & _
                "status, " & _
                "error, " & _
                "quantity, " & _
                "transaction_id, " & _
                "sforceId, " & _
                "from_unixtime(expires_date) as expires_date " & _
                "FROM AppleReceipt  Order by id " & GC_LIMITRESULT
            lvstrColumns = _
                "id, " & _
                "sd_source, " & _
                "created, " & _
                "userId, " & _
                "productId, " & _
                "presentationId, " & _
                "product_id, " & _
                "status, " & _
                "error, " & _
                "quantity, " & _
                "transaction_id, " & _
                "sforceId, " & _
                "expires_date "
            ' Le saque el If C_LINKED And False Then
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_AppleReceipt where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_AppleReceipt", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_AppleReceipt", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_AppleReceipt where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_AppleReceipt", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_AppleReceipt", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM AppleReceipt")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_AppleReceipt", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            'Paypal 
            ProgressBarAdd(pPgbGlobal)
            lvstrExpSql =
                "SELECT " & _
                FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                "from_unixtime(created) as created, " & _
                "cast(mc_gross as decimal(10,2)) as mc_gross, " & _
                "cast(tax as decimal(10,2)) as tax, " & _
                "first_name, " & _
                "last_name, " & _
                "cast(mc_fee as decimal(10,2)) as mc_fee, " & _
                "quantity, " & _
                "payer_email, " & _
                "txn_id, " & _
                "cast(payment_fee as decimal(10,2)) as payment_fee, " & _
                "txn_type, " & _
                "item_name, " & _
                "residence_country, " & _
                "reason_code, " & _
                "parent_txn_id, " & _
                "sforceId, " & _
                "custom, " & _
                "payment_status, " & _
                FieldAddNumber(pBizSystem, "userId ") & " as userId, " & _
                FieldAddNumber(pBizSystem, "presentationId ") & " as presentationId, " & _
                FieldAddNumber(pBizSystem, "productId ") & " as productId " & _
                "from Paypal  " & GC_LIMITRESULT

            lvstrColumns = _
                "id, " & _
                "sd_source, " & _
                "created, " & _
                 "mc_gross, " & _
                "tax, " & _
                "first_name, " & _
                "last_name, " & _
                "mc_fee, " & _
                "quantity, " & _
                "payer_email, " & _
                "txn_id, " & _
                "payment_fee, " & _
                "txn_type, " & _
                "item_name, " & _
                "residence_country, " & _
                "reason_code, " & _
                "parent_txn_id, " & _
                "sforceId, " & _
                "custom, " & _
                "payment_status, " & _
                "userId, " & _
                "presentationId, " & _
                "productId "
            ' Le saque el If C_LINKED And False Then
            If C_LINKED Then
                If (pBizSystem) Then
                    sSpeacialSqlDelete = "DELETE FROM T_Paypal where sd_source = 1"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Paypal", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Paypal", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_Paypal where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Paypal", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Paypal", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Paypal")
                sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Paypal", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            'OSCAR 

            'Oscar Octubre/2014 ----> sacada porque Lucas que dijo que no se usa mas
            ''T_Bundle 
            'ProgressBarAdd(pPgbGlobal)
            'lvstrExpSql =
            '    "SELECT " & _
            '    FieldAddNumber(pBizSystem, "id ") & " as id, " & _
            '    IIf(pBizSystem, "1", "0") & " as sd_source ," & _
            '    "name, " & _
            '    "isDeleted, " & _
            '    "cast(price as decimal(10,2)) as price " & _
            '    "from Bundle  " & GC_LIMITRESULT
            'lvstrColumns = _
            '    "id, " & _
            '    "sd_source, " & _
            '    "name, " & _
            '     "isDeleted, " & _
            '    "price "
            '' Le saque el If C_LINKED And False Then
            'If C_LINKED Then
            '    If (pBizSystem) Then
            '        sSpeacialSqlDelete = "DELETE FROM T_Bundle where sd_source = 1"
            '        'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Paypal", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
            '        sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_Bundle", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
            '    Else
            '        sSpeacialSqlDelete = "DELETE FROM T_Bundle where sd_source = 0"
            '        'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Paypal", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
            '        sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_Bundle", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
            '    End If
            'Else
            '    pPgbParcial.Minimum = 0
            '    plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
            '    pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM Bundle")
            '    sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_Bundle", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            'End If

            '
            'Ojo que esta ahora toma la data de MARKETPLACE
            'T_BundlePresentation 
            'Oscar Octubre/2014 
            '      Me construi una tabla que une el T_BundlePresentation con el codigo de magento de bundel y los codigos de presnetaiociones el content tool
            ' Agregue un group by para eliminar los repetido erroneos
            ProgressBarAdd(pPgbGlobal)

            lvstrExpSql = <![CDATA[
select sd_source, bundleId, presentationId from 
(select * from
(select 0 as sd_source, bundle_id as bundleId , (SELECT cpei.value FROM catalog_product_entity_int cpei 
    WHERE bundle_presentations.presentation_id = cpei.entity_id
    AND cpei.attribute_id = 144 LIMIT 1) as 'presentationId'
from bundle_presentations
where is_fixed = 0
) subquery
where subquery.presentationId is not null and 
      subquery.bundleId is not null  and
      subquery.presentationId > 0 and
      subquery.bundleId > 0) pn
group by pn.bundleId, pn.presentationId

]]>.Value

            lvstrColumns = _
                "sd_source, " & _
                "bundleId, " & _
                "presentationId "
            ' Le saque el If C_LINKED And False Then
            If C_LINKED Then
                If (pBizSystem) Then
                   
                Else
                    sSpeacialSqlDelete = "DELETE FROM T_BundlePresentation where sd_source = 0"
                    'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Paypal", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    sResulta += gfstr_ImportaBulked(goConNear, goMagento, lvstrExpSql, "T_BundlePresentation", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)

                End If
            Else
                pPgbParcial.Minimum = 0
                plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM BundlePresentation")
                sResulta += gfstr_Importa(goMagento, lvstrExpSql, goConNear, "T_BundlePresentation", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            End If

            'Oscar Octubre/2014 -----> la saque porque me dice LUCAS que esta no va mas
            ''T_UserPresentationsBuy 
            'ProgressBarAdd(pPgbGlobal)
            'lvstrExpSql =
            '    "SELECT " & _
            '    FieldAddNumber(pBizSystem, "id ") & " as id, " & _
            '    IIf(pBizSystem, "1", "0") & " as sd_source ," & _
            '    FieldAddNumber(pBizSystem, "userId ") & " as userId, " & _
            '    FieldAddNumber(pBizSystem, "entityId ") & " as entityId, " & _
            '    "entityType, " & _
            '    "presentations, " & _
            '    "cast(price as decimal(10,2)) as price,  " & _
            '     "from_unixtime(created) as created " & _
            '    "from UserPresentationsBuy  " & GC_LIMITRESULT
            'lvstrColumns = _
            '    "id, " & _
            '    "sd_source, " & _
            '    "userId, " & _
            '    "entityId, " & _
            '    "entityType, " & _
            '    "presentations, " & _
            '    "price, " & _
            '    "created "
            '' Le saque el If C_LINKED And False Then
            'If C_LINKED Then
            '    If (pBizSystem) Then
            '        sSpeacialSqlDelete = "DELETE FROM T_UserPresentationsBuy where sd_source = 1"
            '        'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Paypal", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
            '        sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_UserPresentationsBuy", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
            '    Else
            '        sSpeacialSqlDelete = "DELETE FROM T_UserPresentationsBuy where sd_source = 0"
            '        'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Paypal", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
            '        sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_UserPresentationsBuy", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
            '    End If
            'Else
            '    pPgbParcial.Minimum = 0
            '    plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
            '    pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM UserPresentationsBuy")
            '    sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_UserPresentationsBuy", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
            'End If

            If Not (pBizSystem) Then
                'T_MasterReferral
                ProgressBarAdd(pPgbGlobal)
                lvstrExpSql =
                    "SELECT " & _
                    FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                    IIf(pBizSystem, "1", "0") & " as sd_source ," & _
                    "hostCode, " & _
                    "email, " & _
                    "state, " & _
                    "site, " & _
                    "isDeleted, " & _
                    "from_unixtime(reminder) as reminder, " & _
                    "reminderCount " & _
                    "from masterreferral  " & GC_LIMITRESULT
                lvstrColumns = _
                    "id, " & _
                    "sd_source, " & _
                    "hostCode, " & _
                    "email, " & _
                    "state, " & _
                    "site, " & _
                    "isDeleted, " & _
                    "reminder, " & _
                    "reminderCount "
                ' Le saque el If C_LINKED And False Then
                If C_LINKED Then
                    If (pBizSystem) Then
                        sSpeacialSqlDelete = "DELETE FROM T_MasterReferral where sd_source = 1"
                        sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_MasterReferral", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    Else
                        sSpeacialSqlDelete = "DELETE FROM T_MasterReferral where sd_source = 0"
                        sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_MasterReferral", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    End If
                Else
                    pPgbParcial.Minimum = 0
                    plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                    pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM masterreferral")
                    sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_MasterReferral", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
                End If

                'T_MasterUser
                ProgressBarAdd(pPgbGlobal)

                lvstrExpSql =
                    "SELECT " & _
                    FieldAddNumber(pBizSystem, "id ") & " as id, " & _
                    IIf(pBizSystem, "1", "0") & " as sd_source, " & _
                    "userName, " & _
                    "email, " & _
                    "hostCode, " & _
                    "isDeleted, " & _
                    "server " & _
                    "from masteruser  "
                If (pBizSystem) Then
                    lvstrExpSql &= "where server = 'business'"
                Else
                    lvstrExpSql &= "where server = 'education'"
                End If
                lvstrExpSql &= GC_LIMITRESULT & " "
                lvstrColumns = _
                    "id, " & _
                    "sd_source, " & _
                    "userName, " & _
                    "email, " & _
                    "hostCode, " & _
                    "isDeleted, " & _
                    "server "
                ' Le saque el If C_LINKED And False Then
                If C_LINKED Then
                    If (pBizSystem) Then
                        sSpeacialSqlDelete = "DELETE FROM T_MasterUser where sd_source = 1"
                        sResulta += gfstr_ImportaBulked(goConNear, goConnBizContent, lvstrExpSql, "T_MasterUser", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    Else
                        sSpeacialSqlDelete = "DELETE FROM T_MasterUser where sd_source = 0"
                        sResulta += gfstr_ImportaBulked(goConNear, goConnContent, lvstrExpSql, "T_MasterUser", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
                    End If
                Else
                    pPgbParcial.Minimum = 0
                    plblCurrentOp.Text = "Counting records..." : Application.DoEvents()
                    pPgbParcial.Maximum = gflng_GetNumReg(goConnContent, "SELECT COUNT(*) FROM masteruser")
                    sResulta += gfstr_Importa(goConnContent, lvstrExpSql, goConNear, "T_MasterUser", True, pPgbParcial, plblCurrentOp, plblTable, pexError)
                End If
            End If

            Dim oCmd As SqlClient.SqlCommand

            'Oscar Octubre/2014 puse los precios en cero
            ' Tiene que estar primero porque inicializa algunos indices
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating FromStore" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_Presentation Set FromStore = 1, " & _
                "OriginalName = T2.Name, " & _
                "OriginalCreatedDate = Convert(Date,t2.Created), " & _
                "OriginalPrice = 0, " & _
                "Price = 0, " & _
                "publisherId = T2.publisherId, " & _
                "initiativeId = T2.initiativeId, " & _
                "gradeId = T2.gradeId, " & _
                "levelId = T2.levelId, " & _
                "ageId = T2.ageId, " &
                "subjectId = T2.subjectId, " &
                "authorId = T2.authorId " &
                "FROM T_Presentation, T_Presentation T2 where " & _
                "t_Presentation.parentid <> 0  And T_Presentation.parentId  = t2.id and (t2.userId in " & GC_AUTHORS & " )  "  ' si pablo me da el ok cambio esto y le agrego que el T_presentation.userid=1152
            oCmd.ExecuteNonQuery()
            'oscar octubre/2014

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Append UserPresentationbuy Manual" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            If (pBizSystem) Then
                oCmd.CommandText = _
              "INSERT INTO T_UserPresentationsBuy " & _
              " SELECT * FROM T_UserPresentationsBuyManual where sd_source = 1"
            Else
                oCmd.CommandText = _
               "INSERT INTO T_UserPresentationsBuy " & _
               " SELECT * FROM T_UserPresentationsBuyManual where sd_source = 0"
            End If
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Getting Started" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_Presentation Set FromGettingStarted = 1,  " & _
                "OriginalName = T2.Name, OriginalCreatedDate = Convert(Date,t2.Created) " & _
                "FROM T_Presentation, T_Presentation T2 " & _
                "where " & _
                "(t_Presentation.parentid <> 0 " & _
                "	And " & _
                "T_Presentation.parentId = t2.id " & _
                "   And " & _
                "(t2.userId = 88000 " & _
                "or t2.userId = 652 " & _
                "or t2.userId = 538)) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Not From Getting Started" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "update T_Presentation set FromGettingStarted = 0 where FromGettingStarted is null"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Presentation Age" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_Presentation set " & _
                "Age = T_Age.name " & _
                "FROM T_Presentation, T_Age  " & _
                "Where T_Presentation.ageId = T_Age.id"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Presentation Author" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_Presentation set " & _
                "Author = T_Author.name " & _
                "FROM T_Presentation, T_Author  " & _
                "Where T_Presentation.authorId = T_Author.id"
            oCmd.ExecuteNonQuery()


            ' Arreglos de Grades
            ' Inserto un grade nuevo llamado multigrade
            If Not (pBizSystem) Then
                ProgressBarAdd(pPgbGlobal)
                oCmd = goConNear.CreateCommand
                oCmd.CommandTimeout = 999999
                plblCurrentOp.Text = "Generating codigo de Multi Grade" : Application.DoEvents()
                sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
                oCmd.CommandText = "INSERT INTO [dbo].[T_Grade] ([id], [sd_source], [name] ,[IsDeleted]) VALUES( 99999999,0,'MULTIGRADE',0)"
                oCmd.ExecuteNonQuery()
            End If

            'Pongo un codigo 999999 en el caso que sean multiples y sino pongo el detall
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Actualiza Presentations con los Codigos de Grades" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = <![CDATA[
    update T_Presentation 
    set gradeId = 
    case 
         WHEN countGrades > 1 THEN 99999999
         ELSE firstGrade 
    END 
    from T_Presentation, (select presentationId, count(*) as countGrades, min(gradeId) as firstGrade from T_PresentationGrades group by presentationId) PG
    Where T_Presentation.id = PG.presentationId
    ]]>.Value
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Presentation Grade" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_Presentation set " & _
                "Grade = T_Grade.name " & _
                "FROM T_Presentation, T_Grade  " & _
                "Where T_Presentation.gradeId = T_Grade.id"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Presentation Initiative" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_Presentation set " & _
                "Initiative = T_Initiative.name " & _
                "FROM T_Presentation, T_Initiative  " & _
                "Where T_Presentation.initiativeId = T_Initiative.id"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Presentation Level" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_Presentation set " & _
                "LevelType = T_Level.name " & _
                "FROM T_Presentation, T_Level  " & _
                "Where T_Presentation.levelId = T_Level.id"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Presentation Publisher" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_Presentation set " & _
                "publishers = T_Publisher.name " & _
                "FROM T_Presentation, T_Publisher  " & _
                "Where T_Presentation.publisherId = T_Publisher.id"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Presentation Subject" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_Presentation set " & _
                "Subject = T_Subject.name " & _
                "FROM T_Presentation, T_Subject  " & _
                "Where T_Presentation.SubjectId = T_Subject.id"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Not From Store" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "update T_Presentation set FromStore = 0 where FromStore is null"
            oCmd.ExecuteNonQuery()


            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Convert WebPage Collab Type" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                    " update T_slide set slideType = 'Webpage Collab Draw' from T_slide, T_Webpage where slideType = 'Webpage' and T_Webpage.type = 3 and T_Slide.id = T_Webpage.slideId "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Convert WebPage PDF" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                    " update T_slide set slideType = 'Webpage Pdf' from T_slide, T_Webpage where slideType = 'Webpage' and T_Webpage.type = 2 and T_Slide.id = T_Webpage.slideId "
            oCmd.ExecuteNonQuery()


            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Convert WebPage Twitter" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                    " update T_slide set slideType = 'Webpage Twitter' from T_slide, T_Webpage where slideType = 'Webpage' and T_Webpage.type = 1 and T_Slide.id = T_Webpage.slideId "
            oCmd.ExecuteNonQuery()


            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Number of Answers" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                    "update T_Question Set " & _
                    "qanswersnotdeleted = (select count(*) from t_Answer where T_Question.id = T_Answer.questionId And isdeleted = 0 ), " & _
                    "qanswersdeleted = (select count(*) from t_Answer where T_Question.id = T_Answer.questionId And isdeleted = 1 )"
            oCmd.ExecuteNonQuery()


            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Number of Questions" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "update T_Slide Set " & _
                    "qquestionnotdeleted = (select count(*) from T_Question where T_Slide.id = T_Question.slideId And isdeleted = 0 ), " & _
                    "qquestiondeleted = (select count(*) from T_Question where T_Slide.id = T_Question.slideid And isdeleted = 1 )"
            oCmd.ExecuteNonQuery()


            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Number of NonINteractive And Video" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "update t_Slide set  " & _
                "QNonInteractive = case slideType when 'Image' Then 1 Else 0 end,  " & _
                "QVideo = case slideType when 'Video' Then 1 Else 0 end, " & _
                "QDraw  = case slideType when 'Sketch' Then 1 Else 0 end"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Update DNSDomain" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_User SET dnsdomain = right( dbo.T_User.email, len( dbo.T_User.email) - charindex ('@',dbo.T_User.email)) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Engaged User All Time" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_User SET " & _
                "sdn_QStudentEmbed         = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.embed = 1), " & _
                "sdn_QStudentHomeWork      = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 1), " & _
                "sdn_QStudentNotHomeWork   = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 0 and S.embed = 0), " & _
                "sdn_QPresentNoStore       = (SELECT Count(*) From T_Presentation T WHERE T.userId     = T_user.id And t.isDeleted = 0 And t.FromStore = 0 and t.FromGettingStarted = 0 And t.Published = 1), " & _
                "sdn_QSessionsMore5Student = (SELECT Count(*) From VI_Sessions5S S	Where S.teacher_id = T_user.id), " & _
                "sdn_EngagedUSer           = 0," & _
                "sdn_QStudentEmbed_SessValid         = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.embed = 1 and S.sessionSize >= 2), " & _
                "sdn_QStudentHomeWork_SessValid      = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 1 and S.sessionSize >= 2), " & _
                "sdn_QStudentNotHomeWork_SessValid   = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 0 AND S.embed = 0 and S.sessionSize >= 2), " & _
                "sdn_QPresentNoStore_SessValid       = (SELECT Count(*) From T_Presentation T WHERE T.userId     = T_user.id And t.isDeleted = 0 And t.FromStore = 0 and t.FromGettingStarted = 0 And t.Published = 1), " & _
                "sdn_QSessionsMore5Student_SessValid = (SELECT Count(*) From VI_Sessions5S S	Where S.teacher_id = T_user.id), " & _
                "sdn_EngagedUSer_SessValid           = 0"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "UPDATE T_User SET sdn_EngagedUSer = 1 Where sdn_QStudentHomeWork > 15 or sdn_QStudentNotHomeWork > 10 or sdn_QPresentNoStore >= 3 or sdn_QSessionsMore5Student > 5"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "UPDATE T_User SET sdn_EngagedUSer_SessValid = 1 Where sdn_QStudentHomeWork_SessValid >= 5 or sdn_QStudentNotHomeWork_SessValid >= 5"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Engaged 7" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_User SET " & _
                "sdn_QStudentEmbed7         = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.embed = 1 And DateDiff(day,session_date,GETDATE() )  <= 7), " & _
                "sdn_QStudentHomeWork7      = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 1 And DateDiff(day,session_date,GETDATE() )  <= 7), " & _
                "sdn_QStudentNotHomeWork7   = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 0 and S.embed = 0 And DateDiff(day,session_date,GETDATE() )  <= 7), " & _
                "sdn_QPresentNoStore7       = (SELECT Count(*) From T_Presentation T WHERE T.userId     = T_user.id And t.isDeleted = 0 And t.FromStore = 0 and t.FromGettingStarted = 0 And t.Published = 1 And  DateDiff(day,t.created,GETDATE() )  <= 7), " & _
                "sdn_QSessionsMore5Student7 = (SELECT Count(*) From VI_Sessions5S_7 S	Where S.teacher_id = T_user.id), " & _
                "sdn_EngagedUSer7           = 0, " & _
                "sdn_QStudentEmbed7_SessValid         = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.Embed = 1 and S.sessionSize >= 2 And DateDiff(day,session_date,GETDATE() )  <= 7), " & _
                "sdn_QStudentHomeWork7_SessValid      = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 1 and S.sessionSize >= 2 And DateDiff(day,session_date,GETDATE() )  <= 7), " & _
                "sdn_QStudentNotHomeWork7_SessValid   = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 0 and S.embed = 0 and S.sessionSize >= 2 And DateDiff(day,session_date,GETDATE() )  <= 7), " & _
                "sdn_QPresentNoStore7_SessValid       = (SELECT Count(*) From T_Presentation T WHERE T.userId     = T_user.id And t.isDeleted = 0 And t.FromStore = 0 and t.FromGettingStarted = 0 And t.Published = 1 And  DateDiff(day,t.created,GETDATE() )  <= 7), " & _
                "sdn_QSessionsMore5Student7_SessValid = (SELECT Count(*) From VI_Sessions5S_7 S	Where S.teacher_id = T_user.id), " & _
                "sdn_EngagedUSer7_SessValid           = 0"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "UPDATE T_User SET sdn_EngagedUSer7 = 1 Where sdn_QStudentHomeWork7 > 15 or sdn_QStudentNotHomeWork7 > 10 or sdn_QPresentNoStore7 >= 3 or sdn_QSessionsMore5Student7 > 5"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "UPDATE T_User SET sdn_EngagedUSer7_SessValid = 1 Where sdn_QStudentHomeWork7_SessValid >= 5 or sdn_QStudentNotHomeWork7 >= 5"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Engaged 30" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_User SET " & _
                "sdn_QStudentEmbed30         = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.embed = 1 And DateDiff(day,session_date,GETDATE() )  <= 30), " & _
                "sdn_QStudentHomeWork30      = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 1 And DateDiff(day,session_date,GETDATE() )  <= 30), " & _
                "sdn_QStudentNotHomeWork30   = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 0 And S.embed = 0  And DateDiff(day,session_date,GETDATE() )  <= 30), " & _
                "sdn_QPresentNoStore30       = (SELECT Count(*) From T_Presentation T WHERE T.userId     = T_user.id And t.isDeleted = 0 And t.FromStore = 0 and t.FromGettingStarted = 0 And t.Published = 1 And  DateDiff(day,t.created,GETDATE() )  <= 30), " & _
                "sdn_QSessionsMore5Student30 = (SELECT Count(*) From VI_Sessions5S_30 S	Where S.teacher_id = T_user.id), " & _
                "sdn_EngagedUSer30           = 0, " & _
                "sdn_QStudentEmbed30_SessValid         = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.embed = 1 and S.sessionSize >= 2 And DateDiff(day,session_date,GETDATE() )  <= 30), " & _
                "sdn_QStudentHomeWork30_SessValid      = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 1 and S.sessionSize >= 2 And DateDiff(day,session_date,GETDATE() )  <= 30), " & _
                "sdn_QStudentNotHomeWork30_SessValid   = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 0 and S.embed = 0 and S.sessionSize >= 2 And DateDiff(day,session_date,GETDATE() )  <= 30), " & _
                "sdn_QPresentNoStore30_SessValid       = (SELECT Count(*) From T_Presentation T WHERE T.userId     = T_user.id And t.isDeleted = 0 And t.FromStore = 0 and t.FromGettingStarted = 0 And t.Published = 1 And  DateDiff(day,t.created,GETDATE() )  <= 30), " & _
                "sdn_QSessionsMore5Student30_SessValid = (SELECT Count(*) From VI_Sessions5S_30 S	Where S.teacher_id = T_user.id), " & _
                "sdn_EngagedUSer30_SessValid = 0"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "UPDATE T_User SET sdn_EngagedUSer30 = 1 Where sdn_QStudentHomeWork30 > 15 or sdn_QStudentNotHomeWork30 > 10 or sdn_QPresentNoStore30 >= 3 or sdn_QSessionsMore5Student30 > 5"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "UPDATE T_User SET sdn_EngagedUSer30_SessValid = 1 Where sdn_QStudentHomeWork30_SessValid >= 5 or sdn_QStudentNotHomeWork30_SessValid >= 5"
            oCmd.ExecuteNonQuery()


            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Engaged 60" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "UPDATE T_User SET " & _
                "sdn_QStudentEmbed60         = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.embed = 1 And DateDiff(day,session_date,GETDATE() )  <= 60), " & _
                "sdn_QStudentHomeWork60      = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 1 And DateDiff(day,session_date,GETDATE() )  <= 60), " & _
                "sdn_QStudentNotHomeWork60   = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 0 and S.embed = 0 And DateDiff(day,session_date,GETDATE() )  <= 60), " & _
                "sdn_QPresentNoStore60       = (SELECT Count(*) From T_Presentation T WHERE T.userId     = T_user.id And t.isDeleted = 0 And t.FromStore = 0 and t.FromGettingStarted = 0 And t.Published = 1 And  DateDiff(day,t.created,GETDATE() )  <= 60), " & _
                "sdn_QSessionsMore5Student60 = (SELECT Count(*) From VI_Sessions5S_60 S	Where S.teacher_id = T_user.id), " & _
                "sdn_EngagedUSer60           = 0, " & _
                "sdn_QStudentEmbed60_SessValid         = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.embed = 1 and S.sessionSize >= 2 And DateDiff(day,session_date,GETDATE() )  <= 60), " & _
                "sdn_QStudentHomeWork60_SessValid      = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 1 and S.sessionSize >= 2 And DateDiff(day,session_date,GETDATE() )  <= 60), " & _
                "sdn_QStudentNotHomeWork60_SessValid   = (SELECT Count(*) FROM T_Sessions     S WHERE S.teacher_id = T_User.id And Not lead_uid is NULL And S.is_deleted = 0  AND S.homework = 0 and S.embed = 0 and S.sessionSize >= 2 And DateDiff(day,session_date,GETDATE() )  <= 60), " & _
                "sdn_QPresentNoStore60_SessValid       = (SELECT Count(*) From T_Presentation T WHERE T.userId     = T_user.id And t.isDeleted = 0 And t.FromStore = 0 and t.FromGettingStarted = 0 And t.Published = 1 And  DateDiff(day,t.created,GETDATE() )  <= 60), " & _
                "sdn_QSessionsMore5Student60_SessValid  = (SELECT Count(*) From VI_Sessions5S_60 S	Where S.teacher_id = T_user.id), " & _
                "sdn_EngagedUSer60_SessValid           = 0"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "UPDATE T_User SET sdn_EngagedUSer60 = 1 Where sdn_QStudentHomeWork60 > 15 or sdn_QStudentNotHomeWork60 > 10 or sdn_QPresentNoStore60 >= 3 or sdn_QSessionsMore5Student60 > 5"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "UPDATE T_User SET sdn_EngagedUSer60_SessValid = 1 Where sdn_QStudentHomeWork60_SessValid >= 5 or sdn_QStudentNotHomeWork60_SessValid >= 5"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Stage 10" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "update T_User Set sdn_Stage10 = 1, sdn_Stage10Date = created, sdn_Stage20 = 0, sdn_Stage30 = 0, sdn_Stage35 = 0, sdn_Stage40 = 0, sdn_Stage50 = 0"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Stage 20" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "update T_User Set sdn_Stage20Date = (SELECT MIN(session_date) From T_Sessions Where teacher_id = T_User.id And not lead_uid is null and T_Sessions.is_Deleted = 0)"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "update T_USer Set sdn_Stage20 = 1 Where not sdn_Stage20Date is null"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Stage 30" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "update T_User Set sdn_Stage30Date = (Select Gold_Upgraded__c From TSF_Contact Where T_User.sforceId = TSF_Contact.ID )"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "update T_USer Set sdn_Stage30 = 1 Where not sdn_Stage30Date is null"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Stage 35" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "update T_User Set sdn_Stage35Date = (Select School_Upgraded__c From TSF_Contact Where T_User.sforceId = TSF_Contact.ID )"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "update T_USer Set sdn_Stage35 = 1 Where not sdn_Stage35Date is null"
            oCmd.ExecuteNonQuery()
            '
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Stage 40" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "update T_User Set sdn_Stage40Date = (Select Min(created) From T_Presentation Where T_User.id = T_Presentation.userId And FromStore = 0 and FromGettingStarted = 0 and Published = 1)"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "update T_USer Set sdn_Stage40 = 1 Where not sdn_Stage40Date is null"
            oCmd.ExecuteNonQuery()
            '
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Stage 50" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "update T_User " & _
                               "Set sdn_Stage50Date = (SELECT TOP 1 MIN(session_date) as mindate From T_Sessions Where teacher_id = T_User.id And not lead_uid is null And homework = 0 and T_Sessions.is_Deleted = 0 Group by session_uid having count(*) >=5 order by mindate)"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "update T_USer Set sdn_Stage50 = 1 Where not sdn_Stage50Date is null"
            oCmd.ExecuteNonQuery()


            'Dim multisql As String
            'multisql = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TEMP_APAREOSFDC]') AND type in (N'U')) DROP TABLE [dbo].[TEMP_RUNNING_LIVE_SESSIONS];"
            'multisql = multisql + "select teacher_id, convert(date, t_sessions.session_date) as session_date, count(*) as countstudents" & _
            '                      "into(TEMP_RUNNING_LIVE_SESSIONS) " & _
            '                      "from(t_sessions) " & _
            '                      "where(T_Sessions.homework = 0) " & _
            '                      "  and   T_Sessions.is_Deleted = 0 " & _
            '                      "  and   T_sessions.lead_uid is not null " & _
            '                      "  group by teacher_id, convert(date, t_sessions.session_date); "


            '
            'ProgressBarAdd(pPgbGlobal)
            'oCmd = goConNear.CreateCommand
            'oCmd.CommandTimeout = 999999
            'plblCurrentOp.Text = "Generating Stage 50 SessValid" : Application.DoEvents()
            'sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            'oCmd.CommandText = "update T_User " & _
            '                   "Set sdn_Stage50Date_SessValid = (SELECT TOP 1 MIN(session_date) as mindate From T_Sessions Where teacher_id = T_User.id And not lead_uid is null And homework = 0 and T_Sessions.is_Deleted = 0 and T_sessions.sessionSize >= 2 Group by session_uid having count(*) >=5 order by mindate)"
            'oCmd.ExecuteNonQuery()
            'oCmd.CommandText = "update T_USer Set sdn_Stage50_SessValid = 1 Where not sdn_Stage50Date_SessValid is null"
            'oCmd.ExecuteNonQuery()

            '
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Stage 52" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "update T_User Set sdn_Stage52Date = (SELECT TOP 1 MIN(session_date) as mindate From T_Sessions Where teacher_id = T_User.id And not lead_uid is null and homework = 1 and T_Sessions.is_Deleted = 0 Group by session_uid having count(*) >=5 order by mindate)"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "update T_USer Set sdn_Stage52 = 1 Where not sdn_Stage52Date is null"
            oCmd.ExecuteNonQuery()

            'ProgressBarAdd(pPgbGlobal)
            'oCmd = goConNear.CreateCommand
            'oCmd.CommandTimeout = 999999
            'plblCurrentOp.Text = "Generating Stage 52 SessValid" : Application.DoEvents()
            'sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            'oCmd.CommandText = "update T_User " & _
            '                   "Set sdn_Stage52Date_SessValid = (SELECT TOP 1 MIN(session_date) as mindate From T_Sessions Where teacher_id = T_User.id And not lead_uid is null And homework = 1 and T_Sessions.is_Deleted = 0 and T_sessions.sessionSize >= 2 Group by session_uid having count(*) >=5 order by mindate)"
            'oCmd.ExecuteNonQuery()
            'oCmd.CommandText = "update T_USer Set sdn_Stage50_SessValid = 1 Where not sdn_Stage50Date_SessValid is null"
            'oCmd.ExecuteNonQuery()


            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Stage 55" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "update T_User Set sdn_Stage55Date =  " & _
                "  (select top 1 created From (select top 3 P1.userid, created from T_Presentation P1 inner join " & _
                "  (select P.[userId] from T_Presentation P Where P.FromStore = 0 And P.FromGettingStarted =0 and p.Published = 1 And P.isDeleted = 0 Group By P.[userId] Having Count(*) >= 3) P2 " & _
                "  ON P1.userId = P2.userId  where p1.userid = T_User.id and p1.FromStore = 0 and p1.FromGettingStarted = 0 and p1.Published = 1 And p1.isDeleted = 0 order by p1.created asc) ST55 order by created desc)"
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "update T_USer Set sdn_Stage55 = 1 Where not sdn_Stage55Date is null"
            oCmd.ExecuteNonQuery()
            '
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Stage 60" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            'oCmd.CommandText = _
            '    "update T_User Set sdn_Stage60Date = ( " & _
            '    "Select Top 1 M353.session_date " & _
            '    "From ( " & _
            '    "	Select Top 3 TS.teacher_id, TS.session_date " & _
            '    "	From T_Sessions TS Inner Join ( " & _
            '    "		select M5.teacher_id " & _
            '    "		from ( " & _
            '    "			select teacher_id, session_uid " & _
            '    "			From T_Sessions  " & _
            '    "			where is_Deleted = 0 And Not lead_uid is null And Teacher_id = T_User.id " & _
            '    "			Group by teacher_id, session_uid  " & _
            '    "			Having Count(*) >= 5) M5 " & _
            '    "		Group by M5.teacher_id  " & _
            '    "		Having Count(*) >= 3) M35 ON TS.teacher_id = M35.teacher_id  " & _
            '    "	Order By TS.session_date asc) M353 " & _
            '    "Order By M353.session_date Desc )"
            oCmd.CommandText = _
                    "update T_User Set sdn_Stage60Date = ( " & _
                    "select max(DELTOP3.session_date) " & _
                    "From  " & _
                    "(select TOP 3 TS.teacher_id, TS.session_uid, TS.session_date, count(*) as Students " & _
                    "From T_Sessions TS inner join ( " & _
                    "		select M5.teacher_id " & _
                    "		from ( " & _
                    "			select teacher_id, session_uid  " & _
                    "			From T_Sessions " & _
                    "			where is_Deleted = 0 And Not lead_uid is null And Teacher_id = T_user.id " & _
                    "			Group by teacher_id, session_uid " & _
                    "			Having Count(*) >= 5) M5 " & _
                    "		Group by M5.teacher_id " & _
                    "		Having Count(*) >= 3) TABLACONSOLOTEACHER on TS.teacher_id = TABLACONSOLOTEACHER.teacher_id " & _
                    "where is_Deleted = 0 And Not lead_uid is null And TS.Teacher_id = T_user.id " & _
                    "Group by TS.teacher_id, Ts.session_uid, TS.session_date " & _
                    "Having Count(*) >= 5) DELTOP3)"

            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "update T_USer Set sdn_Stage60 = 1 Where not sdn_Stage60Date is null"
            oCmd.ExecuteNonQuery()
            '
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Stage" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "Update T_User set " & _
                "sdn_Stage =  " & _
                "	case when sdn_Stage60 = 1 Then '60' " & _
                "		 when sdn_Stage55 = 1 Then '55' " & _
                "		 when sdn_Stage50 = 1 Then '52' " & _
                "		 when sdn_Stage50 = 1 Then '50' " & _
                "		 when sdn_Stage40 = 1 Then '40' " & _
                "		 when sdn_Stage35 = 1 Then '35' " & _
                "		 when sdn_Stage30 = 1 Then '30' " & _
                "		 when sdn_Stage20 = 1 Then '20' " & _
                "		 else '10' " & _
                "	end, " & _
                "sdn_StageDate =  " & _
                "	case when sdn_Stage60 = 1 Then sdn_Stage60Date  " & _
                "		 when sdn_Stage55 = 1 Then sdn_Stage55Date " & _
                "		 when sdn_Stage52 = 1 Then sdn_Stage52Date " & _
                "		 when sdn_Stage50 = 1 Then sdn_Stage50Date " & _
                "		 when sdn_Stage40 = 1 Then sdn_Stage40Date " & _
                "		 when sdn_Stage35 = 1 Then sdn_Stage35Date " & _
                "		 when sdn_Stage30 = 1 Then sdn_Stage30Date " & _
                "		 when sdn_Stage20 = 1 Then sdn_Stage20Date " & _
                "		 else sdn_Stage10Date " & _
                "	end"
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Engaged EngagedUserV2" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                " update T_User Set [sdn_EngagedUserV2] = [sdn_Stage60Date] where [sdn_EngagedUserV2]>[sdn_Stage60Date] or [sdn_EngagedUserV2] is null " & _
                " update T_User Set [sdn_EngagedUserV2] = [sdn_Stage55Date] where [sdn_EngagedUserV2]>[sdn_Stage55Date] or [sdn_EngagedUserV2] is null " & _
                " update T_User Set [sdn_EngagedUserV2] = [sdn_Stage52Date] where [sdn_EngagedUserV2]>[sdn_Stage52Date] or [sdn_EngagedUserV2] is null " & _
                " update T_User Set [sdn_EngagedUserV2] = [sdn_Stage50Date] where [sdn_EngagedUserV2]>[sdn_Stage50Date] or [sdn_EngagedUserV2] is null  "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Running Sum" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                "update T_Sessions  set runningsum = (select count(*) from t_sessions b where T_Sessions.teacher_id = b.teacher_id and b.is_deleted = 0 and b.sessionsize >= 2 and T_Sessions.id >= b.id  ) " & _
                " where(T_Sessions.sessionsize >= 2 And T_Sessions.is_Deleted = 0) "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Generating Engaged EngagedUserV2_SessValid" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = _
                         "update T_User Set [sdn_EngagedUserV2_SessValid] = (SELECT MIN(session_date) From T_Sessions Where teacher_id = T_User.id And not lead_uid is null and T_Sessions.is_Deleted = 0 and T_Sessions.runningsum >=5)"
            oCmd.ExecuteNonQuery()

            'Oscar Octubre/2014 sacado este calculo porque no tiene que basarse en esto tiene que basarse en la info que viene de magento
            'ProgressBarAdd(pPgbGlobal)
            'oCmd = goConNear.CreateCommand
            'oCmd.CommandTimeout = 999999
            'plblCurrentOp.Text = "Fix Bundle Prices" : Application.DoEvents()
            'sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            'oCmd.CommandText = _
            '    "update T_Presentation " & _
            '    "   set T_Presentation.price = t2.priceunit, " & _
            '    "       T_Presentation.OriginalPrice = t2.priceunit " & _
            '    "from T_presentation, (select T_Presentation.id, T_UserPresentationsBuy.price/(LEN(presentations) - LEN(REPLACE(presentations, ',', ''))+1) as priceunit from T_UserPresentationsBuy " & _
            '    "                       join T_BundlePresentation on T_UserPresentationsBuy.entityId = T_BundlePresentation.bundleId " & _
            '    "                       join T_Presentation on T_BundlePresentation.presentationId = T_Presentation.parentId " & _
            '    "                       where entityType = 'Bundle' and T_UserPresentationsBuy.userId = T_Presentation.userId " & _
            '    "                   ) t2 " & _
            '    "where(T_Presentation.id = T2.Id)"
            'oCmd.ExecuteNonQuery()


ZonaUpdates:

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates User Created...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("T_User", "created", "CR")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates User Last Loguin...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("T_User", "lastLogin", "LL")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates Engaged V2...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("T_User", "sdn_EngagedUserV2", "ENGV2")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates Engaged V2...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("T_User", "sdn_EngagedUserV2_SessValid", "ENGV2_SessValid")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates Presentation Created...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("T_Presentation", "created", "CR")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates Presentation Modified...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("T_Presentation", "modified", "MO")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates Apple Created...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("T_AppleReceipt", "created", "CR")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates User Apple Expires Date...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("T_AppleReceipt", "created", "CR")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Inherit Dates User Paypal Created...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call gp_InheritDate("T_Paypal", "created", "CR")

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Zip from IP ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            Call cIpProcessing.Generate(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable, pexError)

            'explicacion de best ubication
            'Primero seteo precision en 0
            'Luego busco todos los usuarios que me informron zip, esta en zc_zip porque lo emparege contra el mailingpostalcode que me informo en form
            'Luego busco en los fromip information
            'No estoy tomando los domain

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Calculate Best Location Step 1 ...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            oCmd.CommandText = _
               " update t_user set best_location_precision = 0 "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Calculate Best Location Step 2...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            oCmd.CommandText =
                "update t_user " & _
                "set " & _
                "	best_location_precision = 1,  " & _
                "	[best_zipcode] = [ZC_Zip], " & _
                "   [best_city] = [ZC_PrimaryCity], " & _
                "	[best_state] = [ZC_State], " & _
                "            [best_country] = COUNTRYNAME " & _
                "            from t_user " & _
                "            inner join TSF_Contact on t_user.[sforceId] = TSF_Contact.ID " & _
                "            left join Td_countrys on TSF_Contact.ZC_Country = Td_countrys.countrycode  " & _
                "where (best_location_precision = 0 or best_location_precision is null) " & _
                "and TSF_Contact.[ZC_Zip] > 0 "
            oCmd.ExecuteNonQuery()

            ProgressBarAdd(pPgbGlobal)
            plblCurrentOp.Text = "Calculate Best Location Step 3...." : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            oCmd.CommandText =
                "update t_user  " & _
                "set  " & _
                "	best_location_precision = 2,   " & _
                "	[best_zipcode] = fromip_zipcode, " & _
                "   [best_city] = fromip_city, " & _
                "	[best_state] = fromip_state, " & _
                "            [best_country] = COUNTRYNAME " & _
                "from t_user left join Td_countrys on t_user.fromip_country = Td_countrys.countrycode  " & _
                "            where (best_location_precision = 0 Or best_location_precision Is null) " & _
                "and fromip_state is not null and (fromip_state <> '' or fromip_country <> '')"
            oCmd.ExecuteNonQuery()

            plblTable.Text = "Done"
            plblCurrentOp.Text = "Done"
            pPgbParcial.Value = 0
            pPgbGlobal.Value = pPgbGlobal.Maximum
            sResulta += "End " & Now.ToString & vbCrLf

        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString
        End Try

        Call GrabarLog(eLogType.eCONTENT, sResulta)
        Return sResulta



    End Function

    Private Shared Function FieldAddText(pBizSystem As Boolean, FieldName As String) As String
        FieldAddText = IIf(pBizSystem, "if(" & FieldName & "=""""," & FieldName & ",concat(" & C_DeltaText & "," & FieldName & "))", "" & FieldName & "")
    End Function

    Private Shared Function FieldAddNumber(pBizSystem As Boolean, FieldName As String) As String

        FieldAddNumber = FieldName + IIf(pBizSystem, "+" & C_DeltaNumINI, "")

    End Function


End Class
