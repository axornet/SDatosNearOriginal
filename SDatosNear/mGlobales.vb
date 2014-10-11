Imports MySql.Data.MySqlClient
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text.RegularExpressions

Module mGlobales
    Public GC_NEARCONN As String = "Server=192.168.200.250;Database=NearPod;Trusted_Connection=True"
    Public GC_NEARCONNTEST As String = "Server=localhost;Database=NearPodTest;Trusted_Connection=True"

    Public GC_AUTHORS As String = "(133900,129918,391520,391524,391525,391532,391546,391555,391557,391588,391660,395163,414055,433114,454793,481114,485398,519824,519831,99262,519916,519922,439783,519927,519933,1152,133494,519997,519993,519980,519945,520019)"

    'Const GC_NEARCONN As String = "Server=localhost;Database=NearPodBizTest;Trusted_Connection=True"

    Public Const GC_REGEXCSV = _
        "(?:            # Match either" & Chr(10) & _
        " (?>[^"",\n]*) #  0 or more characters except comma, quote or newline" & Chr(10) & _
        "|              # or" & Chr(10) & _
        " ""            #  an opening quote" & Chr(10) & _
        " (?:           #  followed by either" & Chr(10) & _
        "  (?>[^""]*)   #   0 or more non-quote characters" & Chr(10) & _
        " |             #  or" & Chr(10) & _
        "  """"         #   an escaped quote ("""")" & Chr(10) & _
        " )*            #  any number of times" & Chr(10) & _
        " ""            #  followed by a closing quote" & Chr(10) & _
        ")              # End of alternation" & Chr(10) & _
        "(?=,|$)        # Assert that the next character is a comma (or end of line)"

    Public Const GC_BizSystem As Boolean = True
    Public Const GC_EduSystem As Boolean = False

    Public GC_LIMITRESULT As String = "" ' " LIMIT 0,100" ' "" 
    Public GC_LIMITRESULTTXT As Long = 0

    Public Process_BatSFDCin As Boolean = True
    Public Process_BatSFDCout As Boolean = True

    Public goLocal251 As Boolean = False
    Public goNearpodTest As Boolean = False

    Public noBusiness As Boolean = False

    Public goConnPhoenix As MySqlConnection
    Public goConnContent As MySqlConnection

    Public goConnNearAdmin As MySqlConnection

    Public goConnBizPhoenix As MySqlConnection
    Public goConnBizContent As MySqlConnection
    Public goMagento As MySqlConnection

    Public goConNear As SqlConnection

    Public goMainForm As frmMain
    Public goGlobalCancel As Boolean

    Private isExecuting As Boolean
    Private goException As Exception

    Public goLbl As Label

    Public Sub OpenPhoenix()
        Try
            If goLocal251 Then
                goConnPhoenix = New MySqlConnection("Server=192.168.200.251;Port=3307;Database=np_phoenix;Uid=near;Pwd=nearpass;default command timeout=999999;Connection Timeout=30;")
            Else
                goConnPhoenix = New MySqlConnection("Server=np-nr0.nearpod.com;Port=3302;Database=np_phoenix;Uid=near;Pwd=nearpass;default command timeout=9999999;")
            End If
            goConnPhoenix.Open()
        Catch ex As Exception
            Call GlobalErrorHandler(ex, "Error abriendo Phoenix")
        End Try
    End Sub


    Public Sub OpenMagento()
        Try
            If goLocal251 Then
                goMagento = New MySqlConnection("Server=192.168.200.251;Port=3311;Database=marketplace;Uid=near;Pwd=nearpass;default command timeout=999999;Connection Timeout=30;")
            Else
                goMagento = New MySqlConnection("Server=np-nr0.nearpod.com;Port=3304;Database=marketplace;Uid=near;Pwd=nearpass;default command timeout=9999999;")
            End If
            goMagento.Open()
        Catch ex As Exception
            Call GlobalErrorHandler(ex, "Error abriendo Magento")
        End Try
    End Sub

    Public Sub OpenPhoenixBIZ()
        Try
            If goLocal251 Then
                goConnBizPhoenix = New MySqlConnection("Server=192.168.200.251;Port=3308;Database=biz_hub;Uid=near;Pwd=nearpass;default command timeout=999999;Connection Timeout=30;")
            Else
                goConnBizPhoenix = New MySqlConnection("Server=np-nr0.nearpod.com;Port=3303;Database=biz_hub;Uid=near;Pwd=nearpass;default command timeout=9999999;")
            End If
            goConnBizPhoenix.Open()
        Catch ex As Exception
            Call GlobalErrorHandler(ex, "Error abriendo Phoenix")
        End Try
    End Sub


    Public Sub OpenContent()
        Try
            If goLocal251 Then
                goConnContent = New MySqlConnection("Server=192.168.200.251;Port=3306;Database=nearcontent;Uid=near;Pwd=nearpass;default command timeout=999999;Connection Timeout=30;")
            Else
                goConnContent = New MySqlConnection("Server=np-nr0.nearpod.com;Port=3301;Database=nearcontent;Uid=near;Pwd=nearpass;default command timeout=9999999;")
            End If
            goConnContent.Open()
        Catch ex As Exception
            Call GlobalErrorHandler(ex, "Error abriendo Content")
        End Try
    End Sub

    Public Sub OpenNearAdmin()
        Try
            If goLocal251 Then
                goConnNearAdmin = New MySqlConnection("Server=192.168.200.251;Port=3310;Database=nearadmin;Uid=near;Pwd=nearpass;default command timeout=999999;Connection Timeout=30;")
            End If
            goConnNearAdmin.Open()
        Catch ex As Exception
            Call GlobalErrorHandler(ex, "Error abriendo nearadmin")
        End Try
    End Sub

    Public Sub OpenContentBiz()
        Try
            If goLocal251 Then
                goConnBizContent = New MySqlConnection("Server=192.168.200.251;Port=3308;Database=biz_nearcontent;Uid=near;Pwd=nearpass;default command timeout=999999;Connection Timeout=30;")
            Else
                goConnBizContent = New MySqlConnection("Server=np-nr0.nearpod.com;Port=3303;Database=biz_nearcontent;Uid=near;Pwd=nearpass;default command timeout=9999999;")
            End If
            goConnBizContent.Open()
        Catch ex As Exception
            Call GlobalErrorHandler(ex, "Error abriendo Content")
        End Try
    End Sub

    Public Sub OpenNear()
        Try
            If goNearpodTest Then
                goConNear = New SqlConnection(GC_NEARCONNTEST)
            Else
                goConNear = New SqlConnection(GC_NEARCONN)
            End If
            goConNear.Open()
        Catch ex As Exception
            Call GlobalErrorHandler(ex, "Error abriendo Near")
        End Try
    End Sub

    Public Sub OpenConnections()
        Call OpenPhoenix()
        Call OpenPhoenixBIZ()
        Call OpenMagento()
        Call OpenContent()
        Call OpenNearAdmin()
        Call OpenContentBiz()
        Call OpenNear()
    End Sub

    Public Function GetFiles(ByVal root As String, Optional pvstrFiltro As String = "*.*") As System.Collections.Generic.IEnumerable(Of System.IO.FileInfo)
        Return From file In My.Computer.FileSystem.GetFiles _
                  (root, FileIO.SearchOption.SearchTopLevelOnly, pvstrFiltro) _
               Select New System.IO.FileInfo(file)
    End Function

    Public Function gflng_GetNumReg(pConn As IDbConnection, pvstrExpSql As String) As Long
        Dim loCmd As IDbCommand
        Dim lObj As Object
        loCmd = pConn.CreateCommand
        loCmd.CommandText = pvstrExpSql
        lObj = loCmd.ExecuteScalar
        If lObj Is Nothing Then
            Return 0
        Else
            Return CLng(lObj)
        End If
    End Function

    Public Function gfstr_ImportaBulked( _
            pConn As SqlConnection,
            pconnMySql As MySqlConnection,
            pSqlIn As String,
            pTable As String,
            pColumn As String,
            plblCurrentOp As Label,
            plblTable As Label,
            pSpecialSqlDelete As String,
            pexError As Exception) As String
        Dim lvstrExpSql As String = ""
        Dim sResulta As String
        Dim oBulkCopy As System.Data.SqlClient.SqlBulkCopy
        Dim oDrMySql As MySqlDataReader
        Dim oCmdMySql As MySqlCommand
        Dim oCmdSql As SqlCommand
        Dim oColAttrs As New Collection

        gfstr_ImportaBulked = ""
        If goGlobalCancel Or Not pexError Is Nothing Then
            Exit Function
        End If

        Try
            sResulta = "Import table " & pTable & " starting at " & Now.ToString
            plblCurrentOp.Text = "Deleting...." : My.Application.DoEvents()
            plblTable.Text = pTable
            oBulkCopy = New System.Data.SqlClient.SqlBulkCopy(pConn, SqlBulkCopyOptions.TableLock, Nothing)
            With oBulkCopy
                .DestinationTableName = pTable
                .BatchSize = 20000
                .BulkCopyTimeout = 99999999
                .NotifyAfter = 10000
            End With
            goLbl = plblCurrentOp
            AddHandler oBulkCopy.SqlRowsCopied, AddressOf OnSqlRowsCopied

            oCmdSql = pConn.CreateCommand
            With oCmdSql
                If (pSpecialSqlDelete = "") Then
                    .CommandText = "TRUNCATE TABLE " & pTable
                Else
                    .CommandText = pSpecialSqlDelete
                    .CommandTimeout = 99999999
                End If
                .ExecuteNonQuery()
                .CommandText = "SELECT name FROM sys.columns WHERE object_id = OBJECT_ID(N'" & pTable & "')"
                Dim oDrAux As SqlDataReader = .ExecuteReader
                Do While oDrAux.Read
                    oColAttrs.Add(oDrAux(0) & "")
                Loop
                oDrAux.Close()
            End With

            oCmdMySql = pconnMySql.CreateCommand
            With oCmdMySql
                .CommandText = pSqlIn
                .CommandTimeout = 99999999
            End With
            ' Abro el DataReader de MySql
            plblCurrentOp.Text = "Selecting...." : My.Application.DoEvents()
            oDrMySql = oCmdMySql.ExecuteReader(CommandBehavior.SingleResult)
            ' Hago la correspondencia de columnas (Se supone que es la misma cantidad y el mismo orden)

            With oBulkCopy.ColumnMappings
                For nCol As Integer = 0 To oDrMySql.FieldCount - 1
                    .Add(oDrMySql.GetName(nCol), GetExactColumnName(oColAttrs, oDrMySql.GetName(nCol)))
                Next
            End With
            plblCurrentOp.Text = "Importing...." : My.Application.DoEvents()
            oBulkCopy.WriteToServer(oDrMySql)
            oDrMySql.Close()
            goLbl = Nothing

            sResulta += " and Ending at " & Now.ToString & vbCrLf
            Return sResulta
        Catch ex As Exception
            gfstr_ImportaBulked = ex.ToString
            pexError = ex
            goLbl = Nothing
            Call GlobalErrorHandler(ex, "modGlobales.gfstr_ImportaBulked")
        End Try

    End Function

    Private Sub OnSqlRowsCopied(ByVal sender As Object, _
    ByVal args As SqlRowsCopiedEventArgs)
        Application.DoEvents()
        If Not goLbl Is Nothing Then
            goLbl.Text = "Importing: " & args.RowsCopied & " rows..."
        End If
    End Sub


    Private Function GetExactColumnName(pTable As Collection, ByVal pColName As String) As String
        Dim sRes As String = Nothing
        For Each oCol As String In pTable
            If oCol.Trim.ToUpper = pColName.Trim.ToUpper Then
                sRes = oCol
            End If
        Next
        If sRes Is Nothing Then
            MsgBox("No econctre " & pColName)
        End If
        Return sRes
    End Function

    Public Function gfstr_ImportaLinked( _
            pConn As SqlConnection,
            pDatabase As String,
            pSqlIn As String,
            pTable As String,
            pColumn As String,
            plblCurrentOp As Label,
            plblTable As Label,
            pSpecialSqlDelete As String,
            pexError As Exception) As String
        Dim lvstrExpSql As String = ""
        Dim oCmd As SqlCommand
        Dim oConAux As SqlConnection
        Dim sResulta As String

        If goNearpodTest Then
            oConAux = New SqlConnection(GC_NEARCONNTEST)
        Else
            oConAux = New SqlConnection(GC_NEARCONN)
        End If

        gfstr_ImportaLinked = ""
        If goGlobalCancel Or Not pexError Is Nothing Then
            Exit Function
        End If

        Try
            sResulta = "Import table " & pTable & " starting at " & Now.ToString
            oConAux.Open()
            lvstrExpSql = "INSERT INTO " & pTable & "  (" & pColumn & ") SELECT * FROM OPENQUERY(" & pDatabase & ",'" & pSqlIn & "')"
            plblCurrentOp.Text = "Importing...."
            plblTable.Text = pTable
            Application.DoEvents()
            oCmd = pConn.CreateCommand
            With oCmd
                If (pSpecialSqlDelete = "") Then
                    .CommandText = "TRUNCATE TABLE " & pTable
                    .ExecuteNonQuery()
                Else
                    .CommandText = pSpecialSqlDelete
                    .CommandTimeout = 99999999
                    .ExecuteNonQuery()
                End If
                .CommandText = lvstrExpSql
                .CommandTimeout = 99999999
                .ExecuteNonQuery()
            End With
            sResulta += " and Ending at " & Now.ToString & vbCrLf
            Return sResulta
        Catch ex As Exception
            gfstr_ImportaLinked = ex.ToString
            pexError = ex
            Call GlobalErrorHandler(ex, "modGlobales.gfstr_ImportaLinked")
        End Try

    End Function

    Public Function gflng_GetNumReg(pConn As SqlConnection, pTable As String, Optional ByVal pWhere As String = "") As Long
        Dim oCmd As SqlCommand
        Dim lRes As Long
        oCmd = pConn.CreateCommand
        oCmd.Transaction = pConn.BeginTransaction(IsolationLevel.ReadUncommitted)
        oCmd.CommandText = "SELECT COUNT(*) FROM " & pTable & IIf(pWhere <> "", " WHERE " & pWhere, "")
        lRes = oCmd.ExecuteScalar
        oCmd.Transaction.Commit()
        oCmd.Dispose()
        Return lRes
    End Function



    Public Function gfstr_Importa( _
        pConnIn As MySqlConnection, _
        pSqlIn As String, _
        pConnOut As SqlConnection, _
        pTableOut As String, _
        pTruncate As Boolean, _
        ppgbAvance As ProgressBar, _
        plblCurrentOp As Label, _
        plblTable As Label,
        pexError As Exception)

        gfstr_Importa = ""
        If goGlobalCancel Or Not pexError Is Nothing Then
            Exit Function
        End If

        Dim lodrIn As MySqlDataReader
        Dim loCmdIn As MySqlCommand
        Dim lvlngQReg As Long = 0
        Dim tblAttrs As DataTable
        Dim lvstrColumnList As String = ""
        Dim lvstrParameterList As String = ""
        Dim loCmdInsert As SqlCommand
        Dim locolInputColumns As New Dictionary(Of String, String)
        Dim lvlngTotReg As Long = ppgbAvance.Maximum
        Dim lvintNumCols As Integer
        Dim laintColumnas() As Integer
        Dim nCol As Integer = 0
        Dim sResulta As String = ""

        Try
            sResulta = "Import " & ppgbAvance.Maximum & " records into table " & pTableOut & " starting at " & Now.ToString
            plblTable.Text = pTableOut
            plblCurrentOp.Text = "Preparing records..." : Application.DoEvents()
            ppgbAvance.Value = 0
            goGlobalCancel = False
            loCmdIn = pConnIn.CreateCommand
            With loCmdIn
                .CommandText = pSqlIn
                .CommandType = CommandType.Text
            End With
            lodrIn = loCmdIn.ExecuteReader
            '
            ' Armo la lista de columnas del DataReader
            plblCurrentOp.Text = "Getting columns..." : Application.DoEvents()
            lvintNumCols = lodrIn.FieldCount
            ReDim laintColumnas(lvintNumCols)
            For I As Integer = 0 To lvintNumCols - 1
                locolInputColumns.Add(lodrIn.GetName(I).ToUpper, lodrIn.GetName(I).ToUpper)
            Next
            '
            ' Armo la instruccion para el insert
            plblCurrentOp.Text = "Building sentences..." : Application.DoEvents()
            tblAttrs = pConnOut.GetSchema("Columns", New String() {Nothing, Nothing, pTableOut, Nothing})

            For Each oRow As DataRow In tblAttrs.Rows
                If locolInputColumns.ContainsKey(oRow("COLUMN_NAME").ToString.ToUpper) Then
                    If lvstrColumnList <> "" Then
                        lvstrColumnList += ","
                        lvstrParameterList += ","
                    End If
                    lvstrColumnList += oRow("COLUMN_NAME").ToString
                    lvstrParameterList += "@" & oRow("COLUMN_NAME").ToString
                End If
            Next
            loCmdInsert = pConnOut.CreateCommand
            If pTruncate Then
                loCmdInsert.CommandText = "TRUNCATE TABLE " & pTableOut
                loCmdInsert.ExecuteNonQuery()
                loCmdInsert = pConnOut.CreateCommand
            End If
            With loCmdInsert
                .CommandText = "INSERT INTO " & pTableOut & " (" & lvstrColumnList & " ) values (" & lvstrParameterList & ")"
                .CommandType = CommandType.Text
            End With
            For Each oRow As DataRow In tblAttrs.Rows
                If locolInputColumns.ContainsKey(oRow("COLUMN_NAME").ToString.ToUpper) Then
                    Select Case oRow("DATA_TYPE").ToString.ToUpper
                        Case "INT"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.Int)
                        Case "NVARCHAR", "VARCHAR"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.NVarChar, oRow("CHARACTER_MAXIMUM_LENGTH"))
                        Case "DATETIME"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.DateTime)
                        Case "DATE"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.Date)
                        Case "NUMERIC", "FLOAT"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.Float)
                        Case "TINYINT"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.TinyInt)
                        Case "TEXT"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.Text)
                        Case "DECIMAL"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.Decimal)
                        Case Else
                            MsgBox(oRow("DATA_TYPE").ToString)
                    End Select
                    laintColumnas(nCol) = lodrIn.GetOrdinal(oRow("COLUMN_NAME").ToString)
                    nCol += 1
                End If
            Next

            plblCurrentOp.Text = "Processing..." : Application.DoEvents()
            loCmdInsert.Prepare()
            loCmdInsert.Transaction = pConnOut.BeginTransaction
            Do While lodrIn.Read And Not goGlobalCancel
                lvlngQReg += 1
                If lvlngQReg Mod 100 = 0 Then
                    plblCurrentOp.Text = "Processing..." & lvlngQReg & " of " & ppgbAvance.Maximum : Application.DoEvents()
                    If lvlngTotReg > lvlngQReg Then
                        ppgbAvance.Value = lvlngQReg : Application.DoEvents()
                    End If
                    loCmdInsert.Transaction.Commit()
                    loCmdInsert.Transaction = pConnOut.BeginTransaction
                End If
                '
                ' Insertar en resultados
                For I As Integer = 0 To lvintNumCols - 1
                    loCmdInsert.Parameters(I).Value = lodrIn(laintColumnas(I))
                Next
                loCmdInsert.ExecuteNonQuery()
            Loop
            loCmdInsert.Transaction.Commit()
            lodrIn.Close()
            sResulta += " and Ending at " & Now.ToString & vbCrLf
            Return sResulta
        Catch ex As Exception
            gfstr_Importa = ex.ToString
            Call GlobalErrorHandler(ex, "modGlobales.gfstr_Importa")
        End Try
    End Function

    Public Function Str2Null(ByVal Value As String) As Object
        If Value.Trim().Length() = 0 Then
            Return DBNull.Value
        Else
            Return Value
        End If
    End Function


    Public Function gfstr_ImportFromTxt( _
            ByVal pvstrPath As String, _
            ByVal pvstrTable As String, _
            ByVal pTruncate As Boolean, _
            ByRef pexError As Exception, _
            Optional ByVal plblEstado As Label = Nothing, _
            Optional ByVal pgbAvance As ProgressBar = Nothing, _
            Optional ByVal pstrSeparador As String = Chr(34) & "," & Chr(34)) As String
        '
        Dim lCantReg As Long = 0
        Dim lRegAct As Long = 0
        Dim loRdr As IDataReader
        Dim loCmd As IDbCommand
        Dim lCol As Integer
        Dim sr As StreamReader
        Dim sInFile As String
        Dim sLinea As String
        Dim aColTipos As New SortedList
        Dim aColumnasIn() As String
        Dim aImportarIn() As String
        Dim aValoresIn() As String
        Dim dSizeAct As Decimal
        Dim dFullSize As Decimal

        Dim tblAttrs As DataTable
        Dim lvstrColumnList As String = ""
        Dim lvstrParameterList As String = ""
        Dim loCmdInsert As SqlCommand
        Dim iSize As Integer
        Dim sAux As String
        Dim sResulta As String = ""
        Dim sLinea2 As String
        Dim lnErrores As Long

        Try
            gfstr_ImportFromTxt = ""
            If goGlobalCancel Or Not pexError Is Nothing Then
                Exit Function
            End If

            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Abriendo archivo de entrada" : Application.DoEvents()
            End If
            sInFile = pvstrPath
            '
            ' Averiguo el largo del archivo de entrada
            dFullSize = FileLenght(sInFile)
            sResulta = "Import " & dFullSize.ToString & " bytes into table " & pvstrTable & " From file " & pvstrPath & " starting at " & Now.ToString
            sr = New StreamReader(sInFile)
            '
            If Not sr.Peek >= 0 Then
                Exit Function
            End If

            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Analizando cabecera de " & pvstrTable & "..." : Application.DoEvents()
            End If
            '
            ' Leo la cabecera y la transformo en un vector de nombres de columnas
            sLinea = sr.ReadLine
            sLinea = sLinea.Replace(Chr(34), "")
            dSizeAct = sLinea.Length + 2
            aColumnasIn = sLinea.Split(",")
            ReDim aImportarIn(aColumnasIn.Length - 1)

            If Not pgbAvance Is Nothing Then
                pgbAvance.Minimum = 0
                pgbAvance.Value = 0
                pgbAvance.Maximum = 100
            End If
            '
            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Obteniendo estructura de " & pvstrTable & "..." : Application.DoEvents()
            End If
            '
            tblAttrs = goConNear.GetSchema("Columns", New String() {Nothing, Nothing, pvstrTable, Nothing})
            Dim lvstrColumn As String
            For Each oRow As DataRow In tblAttrs.Rows
                lvstrColumn = oRow("COLUMN_NAME").ToString.ToUpper
                If Array.Exists(aColumnasIn, Function(s) s.ToString.ToUpper = lvstrColumn) Then
                    aColTipos.Add(oRow("COLUMN_NAME").ToString.ToUpper, oRow("DATA_TYPE").ToString)
                    If lvstrColumnList <> "" Then
                        lvstrColumnList += ","
                        lvstrParameterList += ","
                    End If
                    lvstrColumnList += oRow("COLUMN_NAME").ToString
                    lvstrParameterList += "@" & oRow("COLUMN_NAME").ToString
                Else
                    'MsgBox(lvstrColumn)
                End If
            Next
            '
            ' Recorro el vector de columnas del txt y lo transformo en S o N para ver si tengo que importar o no
            For lCol = 0 To aColumnasIn.Length - 1
                If aColTipos.ContainsKey(UCase(aColumnasIn(lCol))) Then
                    aImportarIn(lCol) = aColTipos(UCase(aColumnasIn(lCol)))
                Else
                    aImportarIn(lCol) = "NO"
                End If
            Next

            loCmdInsert = goConNear.CreateCommand
            If pTruncate Then
                loCmdInsert.CommandText = "TRUNCATE TABLE " & pvstrTable
                loCmdInsert.ExecuteNonQuery()
                loCmdInsert = goConNear.CreateCommand
            End If
            With loCmdInsert
                .CommandText = "INSERT INTO " & pvstrTable & " (" & lvstrColumnList & " ) values (" & lvstrParameterList & ")"
                .CommandType = CommandType.Text
            End With
            For Each oRow As DataRow In tblAttrs.Rows
                lvstrColumn = oRow("COLUMN_NAME").ToString.ToUpper
                If Array.Exists(aColumnasIn, Function(s) s.ToString.ToUpper = lvstrColumn) Then
                    Select Case oRow("DATA_TYPE").ToString.ToUpper
                        Case "INT"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.Int)
                        Case "NVARCHAR", "VARCHAR", "NCHAR"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.NVarChar, oRow("CHARACTER_MAXIMUM_LENGTH"))
                        Case "DATETIME"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.DateTime)
                        Case "DATE"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.Date)
                        Case "FLOAT"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.Float)
                        Case "TINYINT"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.TinyInt)
                        Case "TEXT"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.Text)
                        Case "DECIMAL", "NUMERIC"
                            loCmdInsert.Parameters.Add("@" & oRow("COLUMN_NAME"), SqlDbType.Decimal)
                            With loCmdInsert.Parameters("@" & oRow("COLUMN_NAME"))
                                .Precision = 18
                            End With
                        Case Else
                            MsgBox("Tipo de datos desconocido: " & oRow("DATA_TYPE").ToString)
                    End Select
                End If
            Next
            '
            '
            lRegAct = 0
            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Importando " & pvstrTable & " " & lRegAct & " Regs." : Application.DoEvents()
            End If
            '
            loCmdInsert.Prepare()
            loCmdInsert.Transaction = goConNear.BeginTransaction
            Do While sr.Peek >= 0
                lRegAct += 1
                sLinea = sr.ReadLine
                sLinea = sLinea.Replace(vbTab, "")
                sLinea = sLinea.Replace(pstrSeparador, vbTab)
                Do While sLinea.Split(vbTab).Length < aColumnasIn.Length And sr.Peek >= 0
                    sLinea2 = sr.ReadLine
                    If sLinea2 <> "" Then
                        sLinea += sLinea2.Replace(vbTab, "")
                        sLinea = sLinea.Replace(pstrSeparador, vbTab)
                    End If
                Loop
                If sLinea.Split(vbTab).Length <> aColumnasIn.Length Then
                    lnErrores += 1
                End If

                dSizeAct += sLinea.Length + 2
                '
                If lRegAct Mod 100 = 0 Then
                    If Not plblEstado Is Nothing Then
                        plblEstado.Text = "Importando " & pvstrTable & " " & lRegAct & " Regs." : Application.DoEvents()
                    End If
                    If Not pgbAvance Is Nothing Then
                        pgbAvance.Value = Min(dSizeAct * 100 / dFullSize, 100)
                    End If
                    ' 
                    ' Comitear los cambios
                    loCmdInsert.Transaction.Commit()
                    loCmdInsert.Transaction = goConNear.BeginTransaction
                End If
                '
                aValoresIn = sLinea.Split(vbTab)
                If aValoresIn.Length = aColumnasIn.Length Then
                    For lCol = 0 To aColumnasIn.Length - 1
                        sAux = aValoresIn(lCol).Replace(Chr(34), "")
                        Select Case aImportarIn(lCol).ToUpper
                            Case "NO"
                                ' Ignorar, no existe ahora en la base de datos
                                sAux = sAux
                            Case "NVARCHAR", "VARCHAR", "NCHAR"
                                iSize = loCmdInsert.Parameters("@" & aColumnasIn(lCol)).Size
                                If iSize < sAux.Length Then
                                    loCmdInsert.Parameters("@" & aColumnasIn(lCol)).Value = Left(Str2Null(sAux), iSize)
                                Else
                                    loCmdInsert.Parameters("@" & aColumnasIn(lCol)).Value = Str2Null(sAux)
                                End If
                            Case "DATETIME"
                                If sAux.Contains("PDT") Or sAux.Contains("PST") Then
                                    sAux = sAux.Replace("PDT", "").Replace("PST", "")
                                    sAux = sAux.Substring(9) & " " & sAux.Substring(0, 8)
                                Else
                                    sAux = sAux.Replace("T", " ").Replace("Z", " ")
                                End If
                                loCmdInsert.Parameters("@" & aColumnasIn(lCol)).Value = Str2Null(sAux)
                            Case "DATE"
                                sAux = sAux.Replace("T", " ").Replace("Z", " ")
                                loCmdInsert.Parameters("@" & aColumnasIn(lCol)).Value = Str2Null(sAux)
                            Case "DECIMAL", "INT", "NUMERIC"
                                loCmdInsert.Parameters("@" & aColumnasIn(lCol)).Value = Val(sAux)
                            Case "TINYINT"
                                If sAux.ToUpper = "TRUE" Or sAux.ToUpper = "YES" Or sAux.ToUpper = "***** YES *****" Then
                                    loCmdInsert.Parameters("@" & aColumnasIn(lCol)).Value = 1
                                Else
                                    loCmdInsert.Parameters("@" & aColumnasIn(lCol)).Value = 0
                                End If
                            Case Else
                                MsgBox("Tipo desconocido " & aImportarIn(lCol))
                        End Select
                    Next
                    loCmdInsert.ExecuteNonQuery()
                Else
                    sResulta += vbCrLf & "Linea invalida" & sLinea & vbCrLf
                    ' No coincide la linea con la cebecera
                End If
                If goGlobalCancel Then Exit Do
            Loop
            loCmdInsert.Transaction.Commit()


            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Importando " & pvstrTable & " " & lRegAct & " Regs." : Application.DoEvents()
            End If
            If Not pgbAvance Is Nothing Then
                pgbAvance.Value = 100
            End If
            sr.Close()
            loCmd = Nothing
            loRdr = Nothing
            sResulta += " and Ending at " & Now.ToString & vbCrLf
        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf
            Call GlobalErrorHandler(ex, "modGlobales.gp_ImportFromTxt")
        End Try

        gfstr_ImportFromTxt = sResulta

    End Function

    Public Function FileLenght(ByVal pvstrFile As String) As Long
        Dim lFI As System.IO.FileInfo
        Dim lLargo As Long = 0
        Try
            lFI = New System.IO.FileInfo(pvstrFile)
            lLargo = lFI.Length
            lFI = Nothing
            Return lLargo
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Private Function Min(ByVal a As Decimal, ByVal b As Decimal) As Decimal
        If a > b Then Return b Else Return a
    End Function


    Public Enum eLogType
        eCONTENT = 1
        ePHOENIX = 2
        eSF = 3
        eSFO = 4
        eGOOGLE = 5
        eAPPLE = 6
        ePAYPAL = 7
        eERROR = 8
    End Enum

    Public Sub GrabarLog(ByVal pLogType As eLogType, ByVal sMsg As String, Optional pblnAppend As Boolean = False)
        Try
            Dim s As String = ""
            Select Case pLogType
                Case eLogType.eCONTENT
                    s = "CONTENT"
                Case eLogType.ePHOENIX
                    s = "PHOENIX"
                Case eLogType.eSF
                    s = "SF"
                Case eLogType.eSFO
                    s = "SFOUT"
                Case eLogType.eGOOGLE
                    s = "GOOGLE"
                Case eLogType.ePAYPAL
                    s = "PAYPAL"
                Case eLogType.eAPPLE
                    s = "APPLE"
                Case eLogType.eERROR
                    s = "ERROR"
            End Select
            My.Computer.FileSystem.WriteAllText(My.Settings.LogFolder & "\SDatosNear-" & s & "-" & Format(Today, "yyyy-MM-dd") & ".log", _
                Format(Now, "yyyy-MM-dd hh:mm:ss") & vbTab & sMsg & vbCrLf, pblnAppend)
        Catch ex As Exception

        End Try

    End Sub


    Public Function gfstr_BackupTable( _
        ByVal pvstrFile As String, _
        ByVal pvstrSqlExp As String, _
        ByVal pvintNumReg As Integer, _
        Optional ByVal plblEstado As Label = Nothing, _
        Optional ByVal pgbAvance As ProgressBar = Nothing) As String
        '
        Dim lRegAct As Long = 0
        Dim loRdr As IDataReader
        Dim loCmd As IDbCommand
        Dim lCantCols As Integer
        Dim lCol As Integer
        Dim sSalida As String
        Dim sw As StreamWriter
        Dim sOutFile As String
        Dim aColTipos(256) As String
        Dim dValDec As Decimal
        Dim sResulta As String = ""
        Dim sTrabajo As String

        Try
            sResulta = _
                "Exporting data to " & pvstrFile & vbCrLf & _
                "Number of reccords: " & pvintNumReg & vbCrLf & _
                "Start " & Now.ToString & vbCrLf
            goGlobalCancel = False
            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Opennig output file..." : Application.DoEvents()
            End If
            sOutFile = pvstrFile
            sw = New StreamWriter(sOutFile, False, System.Text.Encoding.Default)

            If Not pgbAvance Is Nothing Then
                pgbAvance.Minimum = 0
                pgbAvance.Value = 0
                pgbAvance.Maximum = 100
            End If
            '
            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Selecting records..." : Application.DoEvents()
            End If
            loCmd = goConNear.CreateCommand
            loCmd.CommandText = pvstrSqlExp
            loRdr = loCmd.ExecuteReader
            lCantCols = loRdr.FieldCount
            sSalida = ""
            '
            For lCol = 0 To lCantCols - 1
                If lCol > 0 Then sSalida += ","
                sSalida += """" & UCase(loRdr.GetName(lCol)) & """"
                aColTipos(lCol) = loRdr.GetFieldType(lCol).Name
            Next
            sw.WriteLine(sSalida)
            '
            lRegAct = 0
            Do While loRdr.Read
                lRegAct += 1
                If lRegAct = 192567 Then
                    'MsgBox("parar")
                End If
                If lRegAct Mod 100 = 0 Then
                    If Not plblEstado Is Nothing Then
                        plblEstado.Text = "Exporting " & lRegAct & " records" : Application.DoEvents()
                    End If
                    If Not pgbAvance Is Nothing Then
                        pgbAvance.Value = lRegAct * 100 / pvintNumReg
                    End If
                    sw.Flush()
                End If
                sSalida = ""
                For lCol = 0 To lCantCols - 1
                    If lCol > 0 Then
                        sSalida += "," & """"
                    Else
                        sSalida = """"
                    End If
                    Select Case aColTipos(lCol)
                        Case "String"
                            sTrabajo = RTrim(loRdr.Item(lCol) & "")
                            sTrabajo = sTrabajo.Replace(Chr(34), "")
                            sSalida &= sTrabajo
                        Case "DateTime"
                            If IsDBNull(loRdr.Item(lCol)) Then
                                sSalida &= ""
                            Else
                                sSalida &= Format(loRdr.Item(lCol), "yyyy-MM-dd")
                            End If
                        Case "Decimal", "Int", "Int32"
                            dValDec = Val(loRdr.Item(lCol))
                            If Int(dValDec) = dValDec Then
                                sSalida &= dValDec.ToString.Trim
                            Else
                                sSalida &= Format(dValDec, "0.0###").Trim
                            End If
                        Case "Boolean"
                            If loRdr.Item(lCol) Then
                                sSalida &= "true"
                            Else
                                sSalida &= "false"
                            End If
                        Case Else
                            MsgBox("Unknown type " & aColTipos(lCol))
                    End Select
                    sSalida &= """"
                Next
                sw.WriteLine(sSalida)
                If goGlobalCancel Then Exit Do
            Loop
            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Exported " & lRegAct & " records. Done!" : Application.DoEvents()
            End If
            If Not pgbAvance Is Nothing Then
                pgbAvance.Value = 100
            End If
            loRdr.Close()
            sw.Flush()
            sw.Close()
            loCmd = Nothing
            loRdr = Nothing
            sResulta &= _
                "End " & Now.ToString & vbCrLf
        Catch ex As Exception
            sResulta &= _
                "Failed " & ex.ToString
            Call GlobalErrorHandler(ex, "modGlobales.gp_BackupTable")
        End Try
        Return sResulta
    End Function


    Public Sub gp_InheritDate(pvstrTable As String, pvstrSourceField As String, Optional pvstr_Sufix As String = "")
        Dim oCmd As SqlCommand
        oCmd = goConNear.CreateCommand
        With oCmd
            .CommandTimeout = 9999999
            .CommandText = _
                  "UPDATE " & pvstrTable & " SET " & _
                  pvstrTable & ".DT_Date" & pvstr_Sufix & " = TD_DATES.dt_Date," & _
                  pvstrTable & ".Dt_Day" & pvstr_Sufix & " = TD_DATES.Dt_Day," & _
                  pvstrTable & ".Dt_DayOfYear" & pvstr_Sufix & " = TD_DATES.Dt_DayOfYear," & _
                  pvstrTable & ".Dt_Dow" & pvstr_Sufix & " = TD_DATES.Dt_Dow," & _
                  pvstrTable & ".Dt_Month" & pvstr_Sufix & " = TD_DATES.Dt_Month," & _
                  pvstrTable & ".Dt_Quarter" & pvstr_Sufix & " = TD_DATES.Dt_Quarter," & _
                  pvstrTable & ".Dt_Week" & pvstr_Sufix & " = TD_DATES.Dt_Week," & _
                  pvstrTable & ".Dt_WeekRange" & pvstr_Sufix & " = TD_DATES.Dt_WeekRange," & _
                  pvstrTable & ".Dt_WeekStartDay" & pvstr_Sufix & " = TD_DATES.Dt_WeekStartDay," & _
                  pvstrTable & ".Dt_Year" & pvstr_Sufix & " = TD_DATES.Dt_Year ," & _
                  pvstrTable & ".Dt_YearMonth" & pvstr_Sufix & " = TD_DATES.Dt_YearMonth," & _
                  pvstrTable & ".Dt_YearQuarter" & pvstr_Sufix & " = TD_DATES.Dt_YearQuarter " & _
                  "FROM " & pvstrTable & " INNER JOIN TD_DATES ON CAST(" & pvstrTable & "." & pvstrSourceField & " AS DATE) = TD_DATES.dt_Date "
            .ExecuteNonQuery()
        End With
    End Sub

    Public Sub gp_InheritCountry(pvstrTable As String, pvstrSourceField As String, Optional pvstr_Sufix As String = "")
        Dim oCmd, oCmd2 As SqlCommand
        oCmd = goConNear.CreateCommand
        With oCmd
            .CommandTimeout = 9999999
            .CommandText = _
                  "UPDATE " & pvstrTable & " SET " & _
                  pvstrTable & ".COUNTRYNAME" & pvstr_Sufix & " = TD_COUNTRYS.COUNTRYNAME " & _
                  "FROM " & pvstrTable & " INNER JOIN TD_COUNTRYS ON " & pvstrTable & "." & pvstrSourceField & " = TD_COUNTRYS.COUNTRYCODE "
            .ExecuteNonQuery()
        End With
        oCmd2 = goConNear.CreateCommand
        With oCmd2
            .CommandTimeout = 9999999
            .CommandText = _
                  "UPDATE " & pvstrTable & " SET " & _
                  pvstrTable & ".COUNTRYNAME" & pvstr_Sufix & " = 'N/A' " & _
                  "FROM " & pvstrTable & " WHERE " & pvstrTable & ".COUNTRYNAME" & pvstr_Sufix & " is null"
            .ExecuteNonQuery()
        End With


    End Sub


    Public Function gfstr_ImportBulkFromTxt( _
        ByVal pvstrPath As String, _
        ByVal pvstrTable As String, _
        ByVal pTruncate As Boolean, _
        ByRef pexError As Exception, _
        Optional ByVal plblEstado As Label = Nothing, _
        Optional ByVal pgbAvance As ProgressBar = Nothing, _
        Optional ByVal pstrSeparador As String = Chr(34) & "," & Chr(34), _
        Optional ByVal pblnProcesarCSV As Boolean = False, _
        Optional ByVal pastrClaves As String() = Nothing, _
        Optional ByVal pvstrDecimalPoint As String = Nothing) As String
        '
        Dim lCantReg As Long = 0
        Dim lRegAct As Long = 0
        Dim loRdr As IDataReader
        Dim loCmd As IDbCommand
        Dim lCol As Integer
        Dim sr As StreamReader
        Dim sInFile As String
        Dim sLinea As String
        Dim sLineaOrig As String
        Dim aColTipos As New SortedList
        Dim aColumnasIn() As String
        Dim aImportarIn() As String
        Dim aValoresIn() As String
        Dim dSizeAct As Decimal
        Dim dFullSize As Decimal

        Dim tblAttrs As DataTable
        Dim lvstrColumnList As String = ""
        Dim lvstrParameterList As String = ""
        Dim loCmdInsert As SqlCommand
        Dim iSize As Integer
        Dim sAux As String
        Dim sResulta As String = ""
        Dim sLinea2 As String
        Dim lnErrores As Long
        'Dim oBulkCopy As System.Data.SqlClient.SqlBulkCopy
        Dim oDataTable As New DataTable
        Dim nReg As Integer = 0
        Dim oDrAux As SqlDataReader
        Dim dValor As Double
        Dim iValue As Integer
        Dim noImporta As Integer

        Try
            gfstr_ImportBulkFromTxt = ""
            If goGlobalCancel Or Not pexError Is Nothing Then
                Exit Function
            End If

            loCmdInsert = goConNear.CreateCommand
            If pTruncate Then
                loCmdInsert.CommandText = "TRUNCATE TABLE " & pvstrTable
                loCmdInsert.ExecuteNonQuery()
                loCmdInsert = goConNear.CreateCommand
            End If

            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Abriendo archivo de entrada" : Application.DoEvents()
            End If
            sInFile = pvstrPath
            '
            ' Averiguo el largo del archivo de entrada
            dFullSize = FileLenght(sInFile)
            sResulta = "Import " & dFullSize.ToString & " bytes into table " & pvstrTable & " From file " & pvstrPath & " starting at " & Now.ToString
            sr = New StreamReader(sInFile)
            '
            If Not sr.Peek >= 0 Then
                Exit Function
            End If

            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Analizando cabecera de " & pvstrTable & "..." : Application.DoEvents()
            End If
            '
            ' Leo la cabecera y la transformo en un vector de nombres de columnas
            sLinea = sr.ReadLine
            dSizeAct = sLinea.Length + 2
            If pblnProcesarCSV Then
                aColumnasIn = DecodeCSV(sLinea, pstrSeparador)
                ' Corrijo los nombres de las columnas
                For i As Integer = 0 To aColumnasIn.Length - 1
                    aColumnasIn(i) = aColumnasIn(i).Trim.Replace(" ", "_").Replace(".", "").Replace("/", "_").ToUpper
                Next
            Else
                sLinea = sLinea.Replace(Chr(34), "")
                If pstrSeparador.Length = 1 Then
                    aColumnasIn = sLinea.Split(pstrSeparador)
                Else
                    aColumnasIn = sLinea.Split(",")
                End If
            End If
            ReDim aImportarIn(aColumnasIn.Length - 1)
            '
            If Not pgbAvance Is Nothing Then
                pgbAvance.Minimum = 0
                pgbAvance.Value = 0
                pgbAvance.Maximum = 100
            End If
            '
            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Obteniendo estructura de " & pvstrTable & "..." : Application.DoEvents()
            End If
            '
            ' Averiguo las columnas de destino que coinciden con el txt y creo una tabla en memoria
            tblAttrs = goConNear.GetSchema("Columns", New String() {Nothing, Nothing, pvstrTable, Nothing})
            Dim lvstrColumn As String
            Dim lvtypDataType As System.Type = System.Type.GetType("System.String")
            Dim lvintLenght As Integer = 0
            Dim oNuevaRow As DataRow
            Dim oCmdBuscar As SqlCommand = Nothing
            '
            ' Armo el comando de busqueda por clave primaria
            If Not pastrClaves Is Nothing And Not pTruncate Then
                oCmdBuscar = goConNear.CreateCommand
                oCmdBuscar.CommandText = ""
                For Each sKey As String In pastrClaves
                    If oCmdBuscar.CommandText <> "" Then
                        oCmdBuscar.CommandText += " AND "
                    End If
                    oCmdBuscar.CommandText += sKey + " = @" + sKey
                Next
                oCmdBuscar.CommandText = "SELECT * FROM " & pvstrTable & " WHERE " & oCmdBuscar.CommandText
                For Each sKey As String In pastrClaves
                    oCmdBuscar.Parameters.Add("@" & sKey, SqlDbType.NVarChar)
                Next
            End If

            '
            For Each oRow As DataRow In tblAttrs.Rows
                lvstrColumn = oRow("COLUMN_NAME").ToString.ToUpper
                lvintLenght = 0
                If Array.Exists(aColumnasIn, Function(s) s.ToString.ToUpper = lvstrColumn) Then
                    Select Case oRow("DATA_TYPE").ToString.ToUpper
                        Case "INT"
                            lvtypDataType = System.Type.GetType("System.Int32")
                        Case "NVARCHAR", "VARCHAR", "NCHAR"
                            lvtypDataType = System.Type.GetType("System.String")
                            lvintLenght = oRow("CHARACTER_MAXIMUM_LENGTH")
                        Case "DATETIME"
                            lvtypDataType = System.Type.GetType("System.DateTime")
                        Case "DATE"
                            lvtypDataType = System.Type.GetType("System.DateTime")
                        Case "FLOAT"
                            lvtypDataType = System.Type.GetType("System.Double")
                        Case "TINYINT"
                            lvtypDataType = System.Type.GetType("System.Byte")
                        Case "TEXT"
                            lvtypDataType = System.Type.GetType("System.String")
                            lvintLenght = oRow("")
                        Case "DECIMAL", "NUMERIC"
                            lvtypDataType = System.Type.GetType("System.Decimal")
                        Case Else
                            MsgBox("Tipo de datos desconocido: " & oRow("DATA_TYPE").ToString)
                    End Select
                    '
                    With oDataTable.Columns.Add(oRow("COLUMN_NAME").ToString, lvtypDataType)
                        If lvintLenght > 0 Then
                            .MaxLength = lvintLenght
                        End If
                        .AllowDBNull = (oRow("IS_NULLABLE").ToString.ToUpper <> "NO")
                        .ReadOnly = False
                    End With
                    '
                    aColTipos.Add(oRow("COLUMN_NAME").ToString.ToUpper, oRow("DATA_TYPE").ToString)
                Else
                    'MsgBox(lvstrColumn)
                End If
            Next
            '
            ' Recorro el vector de columnas del txt y lo transformo en S o N para ver si tengo que importar o no
            ' desde el lado del txt
            For lCol = 0 To aColumnasIn.Length - 1
                If aColTipos.ContainsKey(UCase(aColumnasIn(lCol))) Then
                    aImportarIn(lCol) = aColTipos(UCase(aColumnasIn(lCol)))
                Else
                    aImportarIn(lCol) = "NO"
                End If
            Next

            loCmdInsert = goConNear.CreateCommand
            If pTruncate Then
                loCmdInsert.CommandText = "TRUNCATE TABLE " & pvstrTable
                loCmdInsert.ExecuteNonQuery()
                loCmdInsert = goConNear.CreateCommand
            End If
            '
            lRegAct = 0
            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Importing to memory  " & pvstrTable & " " & lRegAct & " Records." : Application.DoEvents()
            End If
            '
            Do While sr.Peek >= 0
                lRegAct += 1

                If (GC_LIMITRESULTTXT > 0 And lRegAct > GC_LIMITRESULTTXT) Then
                    Exit Do
                End If

                sLinea = sr.ReadLine
                sLineaOrig = sLinea
                dSizeAct += sLinea.Length + 2
                If pblnProcesarCSV Then
                    Do
                        Do While Not ValidateCSV(sLinea, pstrSeparador) And sr.Peek >= 0
                            sLinea += sr.ReadLine
                        Loop
                        aValoresIn = DecodeCSV(sLinea, pstrSeparador)
                    Loop Until aValoresIn.Length >= aColumnasIn.Length
                Else
                    If pstrSeparador <> vbTab Then
                        sLinea = sLinea.Replace(vbTab, "")
                        sLinea = sLinea.Replace(pstrSeparador, vbTab)
                    End If
                    Do While sLinea.Split(vbTab).Length < aColumnasIn.Length And sr.Peek >= 0
                        sLinea2 = sr.ReadLine
                        dSizeAct += sLinea.Length + 2
                        If sLinea2 <> "" Then
                            sLinea += sLinea2.Replace(vbTab, "")
                            sLinea = sLinea.Replace(pstrSeparador, vbTab)
                        End If
                    Loop
                    aValoresIn = sLinea.Split(vbTab)
                    If sLinea.Split(vbTab).Length <> aColumnasIn.Length Then
                        lnErrores += 1
                    End If
                End If

                '
                If lRegAct Mod 100 = 0 Then
                    If Not plblEstado Is Nothing Then
                        plblEstado.Text = "Importing to memory " & pvstrTable & " " & lRegAct & " Records." : Application.DoEvents()
                    End If
                    If Not pgbAvance Is Nothing Then
                        pgbAvance.Value = Min(dSizeAct * 100 / dFullSize, 100)
                    End If
                End If
                '
                If aValoresIn.Length = aColumnasIn.Length And sLinea.Length >= aColumnasIn.Length Then
                    oNuevaRow = oDataTable.NewRow
                    Try
                        For lCol = 0 To aColumnasIn.Length - 1
                            sAux = aValoresIn(lCol).Replace(Chr(34), "")
                            Select Case aImportarIn(lCol).ToUpper
                                Case "NO"
                                    ' Ignorar, no existe ahora en la base de datos
                                    sAux = sAux
                                Case "NVARCHAR", "VARCHAR", "NCHAR"
                                    iSize = oDataTable.Columns(aColumnasIn(lCol)).MaxLength
                                    If iSize < sAux.Length Then
                                        oNuevaRow(aColumnasIn(lCol)) = Left(Str2Null(sAux), iSize)
                                    Else
                                        oNuevaRow(aColumnasIn(lCol)) = Str2Null(sAux)
                                    End If
                                Case "DATETIME"
                                    If sAux.Contains("PDT") Or sAux.Contains("PST") Then
                                        sAux = sAux.Replace("PDT", "").Replace("PST", "")
                                        sAux = sAux.Substring(9) & " " & sAux.Substring(0, 8)
                                    Else
                                        sAux = sAux.Replace("T", " ").Replace("Z", " ")
                                    End If
                                    oNuevaRow(aColumnasIn(lCol)) = Str2Null(sAux)
                                Case "DATE"
                                    sAux = sAux.Replace("T", " ").Replace("Z", " ")
                                    oNuevaRow(aColumnasIn(lCol)) = Str2Null(sAux)
                                Case "INT"
                                    If Not pvstrDecimalPoint Is Nothing Then
                                        If pvstrDecimalPoint = "." Then
                                            sAux = sAux.Replace(",", "") ' Borro las comas
                                        Else
                                            sAux = sAux.Replace(".", "") ' Borro los puntos
                                        End If
                                        If sAux.Contains(pvstrDecimalPoint) Then
                                            sAux = sAux.Substring(0, sAux.IndexOf(pvstrDecimalPoint))
                                        End If
                                        iValue = Val(sAux)
                                    Else
                                        If Not Integer.TryParse(sAux, iValue) Then
                                            iValue = Val(sAux.Replace(".", "").Replace(",", ""))
                                        End If
                                    End If
                                    oNuevaRow(aColumnasIn(lCol)) = iValue
                                Case "DECIMAL", "NUMERIC"
                                    dValor = 0
                                    If sAux <> "" Then
                                        If Not pvstrDecimalPoint Is Nothing Then
                                            If pvstrDecimalPoint = "." Then
                                                sAux = sAux.Replace(",", "") ' Borro las comas
                                            Else
                                                sAux = sAux.Replace(".", "") ' Borro los puntos
                                            End If
                                            If Val("1.1") <> 1.1 Then
                                                dValor = Val(sAux.Replace(".", ","))
                                            Else
                                                dValor = Val(sAux.Replace(",", "."))
                                            End If
                                        Else
                                            If Not Double.TryParse(sAux, dValor) Then
                                                If Val("1.1") <> 1.1 Then
                                                    dValor = Val(sAux.Replace(".", ","))
                                                Else
                                                    dValor = Val(sAux.Replace(",", "."))
                                                End If
                                            End If
                                        End If
                                    End If
                                    oNuevaRow(aColumnasIn(lCol)) = dValor
                                Case "TINYINT"
                                    If sAux.ToUpper = "TRUE" Or sAux.ToUpper = "YES" Or sAux.ToUpper = "***** YES *****" Then
                                        oNuevaRow(aColumnasIn(lCol)) = 1
                                    Else
                                        oNuevaRow(aColumnasIn(lCol)) = 0
                                    End If
                                Case Else
                                    MsgBox("Tipo desconocido " & aImportarIn(lCol))
                            End Select
                        Next
                        ' Si hay clave primaria, busco por la clave para no repetir
                        If Not pastrClaves Is Nothing And Not pTruncate Then
                            For Each sKey As String In pastrClaves
                                oCmdBuscar.Parameters("@" + sKey).Value = oNuevaRow(sKey)
                            Next
                            ' Si no esta la clave, lo agrego a la lista de registros a importar
                            oDrAux = oCmdBuscar.ExecuteReader
                            If Not oDrAux.Read Then
                                oDataTable.Rows.Add(oNuevaRow)
                            End If
                            oDrAux.Close()
                        Else
                            ' No hay clave informada
                            oDataTable.Rows.Add(oNuevaRow)
                        End If
                    Catch ex As Exception
                        ' Probelma de nulos probablemente
                        sResulta += vbCrLf & "Linea invalida: " & sLineaOrig & vbCrLf
                    End Try
                Else
                    sResulta += vbCrLf & "Linea invalida: " & sLineaOrig & vbCrLf
                    ' No coincide la linea con la cebecera
                End If

                If Int(lRegAct / 50000) = lRegAct / 50000 Then
                    Call GraboRegistrosDeUnSaque(pvstrTable, oDataTable, noImporta, sResulta)
                    oDataTable.Clear()
                End If

                If goGlobalCancel Then Exit Do
            Loop


            If Not plblEstado Is Nothing Then
                plblEstado.Text = "Storing in Sql Database " & pvstrTable & " " & lRegAct & " Regs." : Application.DoEvents()
            End If

            '


            '
            ' Grabo todos los registros de un saque



            'oBulkCopy = New System.Data.SqlClient.SqlBulkCopy(goConNear)
            'oBulkCopy.DestinationTableName = pvstrTable
            'oBulkCopy.BatchSize = 1000
            'Dim strMapStr As String = ""
            'With oBulkCopy.ColumnMappings
            '    For Each oCol As DataColumn In oDataTable.Columns
            '        noImporta = noImporta + 1

            '        .Add(oCol.ColumnName, oCol.ColumnName)
            '        strMapStr += oCol.ColumnName & vbTab

            '    Next
            'End With
            'Try
            '    oBulkCopy.WriteToServer(oDataTable)
            'Catch ex As SqlException When ex.Number = 2601 Or ex.Number = 2627 ' Duplicate Key
            '    sResulta += " DUPLICATED FILE "
            'End Try

            Call GraboRegistrosDeUnSaque(pvstrTable, oDataTable, noImporta, sResulta)
            oDataTable = Nothing

            If Not pgbAvance Is Nothing Then
                pgbAvance.Value = 100
            End If
            sr.Close()
            loCmd = Nothing
            loRdr = Nothing
            sResulta += " and Ending at " & Now.ToString & vbCrLf
        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf
            Call GlobalErrorHandler(ex, "modGlobales.gp_ImportFromTxt")
        End Try

        gfstr_ImportBulkFromTxt = sResulta

    End Function

    Public Sub GraboRegistrosDeUnSaque(pvstrTable As String, oDataTable As Object, ByRef noImporta As Integer, ByRef sResulta As String)

        Dim oBulkCopy As System.Data.SqlClient.SqlBulkCopy
        
        oBulkCopy = New System.Data.SqlClient.SqlBulkCopy(goConNear)
        oBulkCopy.DestinationTableName = pvstrTable
        oBulkCopy.BatchSize = 1000
        Dim strMapStr As String = ""
        With oBulkCopy.ColumnMappings
            For Each oCol As DataColumn In oDataTable.Columns
                noImporta = noImporta + 1

                .Add(oCol.ColumnName, oCol.ColumnName)
                strMapStr += oCol.ColumnName & vbTab

            Next
        End With
        Try
            oBulkCopy.WriteToServer(oDataTable)
        Catch ex As SqlException When ex.Number = 2601 Or ex.Number = 2627 ' Duplicate Key
            sResulta += " DUPLICATED FILE "
        End Try

    End Sub

    Public Sub ProgressBarAdd(pb As ProgressBar)
        If pb.Maximum = pb.Value Then
            pb.Maximum += 1
        End If
        pb.Value += 1
    End Sub

    Public Function ValidateCSV(ByVal strLine As String, Optional ByVal strDelim As String = "") As Boolean
        Dim strPattern As String
        Dim objMatch As Match

        ' build a pattern
        If strDelim = vbTab Then
            strPattern = "^" ' anchor to start of the string
            strPattern += "(?:""(?<value>(?:""""|[^""\f\r])*)""|(?<value>[^\t\f\r""]*))"
            strPattern += "(?:\t(?:[ ,]*""(?<value>(?:""""|[^""\f\r])*)""|(?<value>[^\t\f\r""]*)))*"
            strPattern += "$" ' anchor to the end of the string
        ElseIf strDelim = "," Then
            strPattern = "^" ' anchor to start of the string
            strPattern += "(?:""(?<value>(?:""""|[^""\f\r])*)""|(?<value>[^,\f\r""]*))"
            strPattern += "(?:,(?:[ \t]*""(?<value>(?:""""|[^""\f\r])*)""|(?<value>[^,\f\r""]*)))*"
            strPattern += "$" ' anchor to the end of the string
        Else
            Return False
        End If

        ' get the match
        objMatch = Regex.Match(strLine, strPattern)

        ' if RegEx match was ok
        Return objMatch.Success
    End Function

    Public Function DecodeCSV(ByVal strLine As String, Optional ByVal strDelim As String = "") As String()

        Dim strPattern As String
        Dim objMatch As Match

        ' build a pattern
        If strDelim = vbTab Then
            strPattern = "^" ' anchor to start of the string
            strPattern += "(?:""(?<value>(?:""""|[^""\f\r])*)""|(?<value>[^\t\f\r""]*))"
            strPattern += "(?:\t(?:[ ,]*""(?<value>(?:""""|[^""\f\r])*)""|(?<value>[^\t\f\r""]*)))*"
            strPattern += "$" ' anchor to the end of the string
        ElseIf strDelim = "," Then
            strPattern = "^" ' anchor to start of the string
            strPattern += "(?:""(?<value>(?:""""|[^""\f\r])*)""|(?<value>[^,\f\r""]*))"
            strPattern += "(?:,(?:[ \t]*""(?<value>(?:""""|[^""\f\r])*)""|(?<value>[^,\f\r""]*)))*"
            strPattern += "$" ' anchor to the end of the string
        Else
            Throw New ApplicationException("Bad Delimiter: " & strDelim)
        End If

        ' get the match
        objMatch = Regex.Match(strLine, strPattern)

        ' if RegEx match was ok
        If objMatch.Success Then
            Dim objGroup As Group = objMatch.Groups("value")
            Dim intCount As Integer = objGroup.Captures.Count
            Dim arrOutput(intCount - 1) As String

            ' transfer data to array
            For i As Integer = 0 To intCount - 1
                Dim objCapture As Capture = objGroup.Captures.Item(i)
                arrOutput(i) = objCapture.Value

                ' replace double-escaped quotes
                arrOutput(i) = arrOutput(i).Replace("""""", """")
            Next

            ' return the array
            Return arrOutput
        Else
            Throw New ApplicationException("Bad CSV line: " & strLine)
        End If

    End Function

End Module
