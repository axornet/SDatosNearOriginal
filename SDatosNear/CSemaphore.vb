Imports System.Data.SqlClient

Public Class CSemaphore

    Private Sub lp_TestSmaphore()
        Dim oCmd As SqlCommand
        If gflng_GetNumReg(goConNear, "t_semaphore", "sf_id=1") = 0 Then
            oCmd = goConNear.CreateCommand
            With oCmd
                .CommandTimeout = 9999999
                .CommandText = "INSERT INTO [T_Semaphore] ([sf_id],[sf_Begin],[sf_End]) VALUES(1,null,null)"
                .ExecuteNonQuery()
            End With
        End If
    End Sub

    Public Sub New()
        Call lp_TestSmaphore()
    End Sub

    Public Sub BeginProcess()
        Dim oCmd As SqlCommand
        Call lp_TestSmaphore()
        oCmd = goConNear.CreateCommand
        With oCmd
            .CommandTimeout = 9999999
            .CommandText = "UPDATE [T_Semaphore] SET  [sf_Begin] = CURRENT_TIMESTAMP, [sf_End] = Null Where [sf_id] = 1"
            .ExecuteNonQuery()
        End With
    End Sub

    Public Sub EndProcess()
        Dim oCmd As SqlCommand
        Call lp_TestSmaphore()
        oCmd = goConNear.CreateCommand
        With oCmd
            .CommandTimeout = 9999999
            .CommandText = "UPDATE [T_Semaphore] SET  [sf_End] = CURRENT_TIMESTAMP Where [sf_id] = 1"
            .ExecuteNonQuery()
        End With
    End Sub

    Public Function InProcess() As Boolean
        Return gflng_GetNumReg(goConNear, "t_semaphore", "sf_id = 1 and not sf_Begin is null and sf_End is null ") = 1
    End Function

    Public Function Terminated() As Boolean
        Return gflng_GetNumReg(goConNear, "t_semaphore", "sf_id = 1 and not sf_End is null ") = 1
    End Function

End Class
