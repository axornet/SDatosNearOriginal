<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportPhoenix
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.grpStatus = New System.Windows.Forms.GroupBox()
        Me.lblTable = New System.Windows.Forms.Label()
        Me.lblCurrentOp = New System.Windows.Forms.Label()
        Me.pgbGlobal = New System.Windows.Forms.ProgressBar()
        Me.pgbCurrent = New System.Windows.Forms.ProgressBar()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdStartCancelExit = New System.Windows.Forms.Button()
        Me.txtResultado = New System.Windows.Forms.TextBox()
        Me.grpStatus.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpStatus
        '
        Me.grpStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpStatus.Controls.Add(Me.lblTable)
        Me.grpStatus.Controls.Add(Me.lblCurrentOp)
        Me.grpStatus.Controls.Add(Me.pgbGlobal)
        Me.grpStatus.Controls.Add(Me.pgbCurrent)
        Me.grpStatus.Controls.Add(Me.Label2)
        Me.grpStatus.Controls.Add(Me.Label1)
        Me.grpStatus.Location = New System.Drawing.Point(12, 12)
        Me.grpStatus.Name = "grpStatus"
        Me.grpStatus.Size = New System.Drawing.Size(643, 144)
        Me.grpStatus.TabIndex = 1
        Me.grpStatus.TabStop = False
        Me.grpStatus.Text = "Process Status"
        '
        'lblTable
        '
        Me.lblTable.AutoSize = True
        Me.lblTable.Location = New System.Drawing.Point(86, 87)
        Me.lblTable.Name = "lblTable"
        Me.lblTable.Size = New System.Drawing.Size(24, 13)
        Me.lblTable.TabIndex = 5
        Me.lblTable.Text = "Idle"
        '
        'lblCurrentOp
        '
        Me.lblCurrentOp.AutoSize = True
        Me.lblCurrentOp.Location = New System.Drawing.Point(86, 113)
        Me.lblCurrentOp.Name = "lblCurrentOp"
        Me.lblCurrentOp.Size = New System.Drawing.Size(24, 13)
        Me.lblCurrentOp.TabIndex = 4
        Me.lblCurrentOp.Text = "Idle"
        '
        'pgbGlobal
        '
        Me.pgbGlobal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pgbGlobal.Location = New System.Drawing.Point(89, 49)
        Me.pgbGlobal.Name = "pgbGlobal"
        Me.pgbGlobal.Size = New System.Drawing.Size(522, 23)
        Me.pgbGlobal.TabIndex = 3
        '
        'pgbCurrent
        '
        Me.pgbCurrent.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pgbCurrent.Location = New System.Drawing.Point(89, 20)
        Me.pgbCurrent.Name = "pgbCurrent"
        Me.pgbCurrent.Size = New System.Drawing.Size(522, 23)
        Me.pgbCurrent.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(22, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Global:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(22, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Current:"
        '
        'cmdStartCancelExit
        '
        Me.cmdStartCancelExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStartCancelExit.Location = New System.Drawing.Point(570, 322)
        Me.cmdStartCancelExit.Name = "cmdStartCancelExit"
        Me.cmdStartCancelExit.Size = New System.Drawing.Size(84, 36)
        Me.cmdStartCancelExit.TabIndex = 2
        Me.cmdStartCancelExit.Text = "Start"
        Me.cmdStartCancelExit.UseVisualStyleBackColor = True
        '
        'txtResultado
        '
        Me.txtResultado.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtResultado.Location = New System.Drawing.Point(13, 162)
        Me.txtResultado.Multiline = True
        Me.txtResultado.Name = "txtResultado"
        Me.txtResultado.ReadOnly = True
        Me.txtResultado.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtResultado.Size = New System.Drawing.Size(642, 154)
        Me.txtResultado.TabIndex = 8
        '
        'frmImportPhoenix
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(667, 364)
        Me.Controls.Add(Me.txtResultado)
        Me.Controls.Add(Me.cmdStartCancelExit)
        Me.Controls.Add(Me.grpStatus)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmImportPhoenix"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Import Phoenix"
        Me.grpStatus.ResumeLayout(False)
        Me.grpStatus.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grpStatus As System.Windows.Forms.GroupBox
    Friend WithEvents pgbGlobal As System.Windows.Forms.ProgressBar
    Friend WithEvents pgbCurrent As System.Windows.Forms.ProgressBar
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdStartCancelExit As System.Windows.Forms.Button
    Friend WithEvents lblCurrentOp As System.Windows.Forms.Label
    Friend WithEvents lblTable As System.Windows.Forms.Label
    Friend WithEvents txtResultado As System.Windows.Forms.TextBox
End Class
