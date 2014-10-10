Imports System.Data.SqlClient

Public Class cTest
    Public Shared Function Test1() As String

        Dim sResulta As String = ""

        Dim oCmd As SqlCommand
        Dim oConAux As SqlConnection
        Dim pDatabase As String
        Dim Psql As String
        Dim lodrIn As IDataReader
        Dim lvintNumCols As Integer

        If goNearpodTest Then
            oConAux = New SqlConnection(GC_NEARCONNTEST)
        Else
            oConAux = New SqlConnection(GC_NEARCONN)
        End If

        pDatabase = "[PHOENIX]"
        Psql = "select * from lead where id = 4040"
        oCmd = oConAux.CreateCommand
        With oCmd
            .CommandText = "SELECT * FROM OPENQUERY(" & pDatabase & ",'" & Psql & "')"
            .CommandType = CommandType.Text
        End With
        lodrIn = oCmd.ExecuteReader
        lodrIn.Close()
        Return sResulta


    End Function

        

End Class
