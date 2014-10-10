Public Class frmMain

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim Cadena
        Dim pos

        goMainForm = Me
        
        For i = 0 To System.Environment.CommandLine.Split("/").Length - 1
            If Trim(System.Environment.CommandLine.Split("/")(i)) = "LOCAL251" Then
                goLocal251 = True
                Me.Text = Me.Text & " FROM: 251"
            End If
        Next

        For i = 0 To System.Environment.CommandLine.Split("/").Length - 1
            If Trim(System.Environment.CommandLine.Split("/")(i)) = "NEARPODTEST" Then
                goNearpodTest = True
                Me.Text = Me.Text & " To: TEST DATABASE"
            End If
        Next

        For i = 0 To System.Environment.CommandLine.Split("/").Length - 1
            If Trim(System.Environment.CommandLine.Split("/")(i)) = "NOBUSINESS" Then
                noBusiness = True
                Me.Text = Me.Text & " NO BUSINESS"
            End If
        Next

        For i = 0 To System.Environment.CommandLine.Split("/").Length - 1
            If Trim(System.Environment.CommandLine.Split("/")(i)) = "NOBATSFDC" Then
                Process_BatSFDCin = False
                Process_BatSFDCout = False
                Me.Text = Me.Text & " NO SALESFORCE "
            End If
        Next

        For i = 0 To System.Environment.CommandLine.Split("/").Length - 1
            If Trim(System.Environment.CommandLine.Split("/")(i)) = "NOBATSFDCIN" Then
                Process_BatSFDCin = False
                Me.Text = Me.Text & " NO SALESFORCE "
            End If
        Next

        For i = 0 To System.Environment.CommandLine.Split("/").Length - 1
            If Trim(System.Environment.CommandLine.Split("/")(i)) = "NOBATSFDCOUT" Then
                Process_BatSFDCout = False
                Me.Text = Me.Text & " NO SALESFORCE "
            End If
        Next


        For i = 0 To System.Environment.CommandLine.Split("/").Length - 1
            Cadena = Trim(System.Environment.CommandLine.Split("/")(i))
            pos = Cadena.indexof("LIMIT")
            If pos >= 0 Then
                GC_LIMITRESULT = " LIMIT 0," + Trim(Mid(Cadena, pos + 6)) + " "
                Me.Text = Me.Text & " LIMIT:  " + Trim(Mid(Cadena, pos + 6))
                GC_LIMITRESULTTXT = Val(Trim(Mid(Cadena, pos + 6)))
            End If
        Next

        Call OpenConnections()

        If System.Environment.CommandLine.Split("/").Length > 1 Then
            If Trim(System.Environment.CommandLine.Split("/")(1)) = "AUTORUN" Then
                Dim f As New frmImportAll
                f.StartCancelStop(True)
                End
            End If
        End If

    End Sub


    Private Sub cmdImportPhoenix()
        Dim f As New frmImportPhoenix
        f.ShowDialog()
    End Sub

    Private Sub cmdImportBizPhoenix()
        Dim f As New frmImportBizPhoenix
        f.ShowDialog()
    End Sub

    Private Sub cmdImportContent()
        Dim f As New frmImportContent
        f.ShowDialog()
    End Sub

    Private Sub cmdImportBizContent()
        Dim f As New frmImportBizContent
        f.ShowDialog()
    End Sub

    Private Sub cmdImportSF()
        Dim f As New frmImportSF
        f.ShowDialog()
    End Sub

    Private Sub cmdImportGoogle()
        Dim f As New frmImportGoogle
        f.ShowDialog()
    End Sub

    Private Sub cmdImportApple()
        Dim f As New frmImportApple
        f.ShowDialog()
    End Sub

    Private Sub cmdImportPaypal()
        Dim f As New frmImportPayPal
        f.ShowDialog()
    End Sub

    Private Sub cmdImportALL()
        Dim f As New frmImportAll
        f.ShowDialog()
    End Sub

    Private Sub cmdTruncateAll()
        Dim f As New frmTruncateAll
        f.ShowDialog()
    End Sub

    Private Sub cmdExportSF()
        Dim f As New frmExportSF
        f.ShowDialog()
    End Sub

    Private Sub cmdTransacSDNEAR()
        Dim f As New frmTransacSDNEAR
        f.ShowDialog()
    End Sub

    Private Sub cmdApareo()
        Dim f As New frmApareao
        f.ShowDialog()
    End Sub

    Private Sub cmdTest1()
        cTest.Test1()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub ImportPhoenixToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ImportPhoenixToolStripMenuItem.Click
        Call cmdImportPhoenix()
    End Sub

    Private Sub ImportContentToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ImportContentToolStripMenuItem.Click
        Call cmdImportContent()
    End Sub

    Private Sub ImportSFToolStripMenuItem_Click(sender As Object, e As System.EventArgs) Handles ImportSFToolStripMenuItem.Click
        Call cmdImportSF()
    End Sub

    Private Sub ImportAllToolStripMenuItem1_lick(sender As System.Object, e As System.EventArgs) Handles ImportAllToolStripMenuItem.Click
        Call cmdImportALL()
    End Sub

    Private Sub ExportAllToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExportAllToolStripMenuItem.Click
        Call cmdExportSF()
    End Sub

    Private Sub ImportBizPhoenixToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ImportPhoenixToolStripMenuItem.Click
        Call cmdImportPhoenix()
    End Sub

    Private Sub ImportBizPhoenixToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles ImportBizPhoenixToolStripMenuItem.Click
        Call cmdImportBizPhoenix()
    End Sub

    Private Sub ImportBizContentToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles ImportBizContentToolStripMenuItem.Click
        Call cmdImportBizContent()
    End Sub

    Private Sub TruncateALLToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TruncateALLToolStripMenuItem.Click
        Call cmdTruncateAll()
    End Sub

    
    Private Sub ApareoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ApareoToolStripMenuItem.Click
        Call cmdApareo()
    End Sub

    Private Sub Test1_Click(sender As System.Object, e As System.EventArgs) Handles Test1.Click
        Call cmdTest1()
    End Sub

    Private Sub TransacSDNEAR_Click(sender As System.Object, e As System.EventArgs) Handles TransacSDNEAR.Click
        Call cmdTransacSDNEAR()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Call cmdImportGoogle()
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem2.Click
        Call cmdImportApple()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem3.Click
        Call cmdImportPayPal()
    End Sub
End Class