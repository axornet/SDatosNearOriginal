
Public Class frmFatalError
    Inherits System.Windows.Forms.Form

    Private mvblnDetails As Boolean

#Region " Código generado por el Diseñador de Windows Forms "

    Public Sub New()
        MyBase.New()

        mvblnDetails = False

        'El Diseñador de Windows Forms requiere esta llamada.
        InitializeComponent()

        SetSize()

        'Agregar cualquier inicialización después de la llamada a InitializeComponent()

    End Sub

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms requiere el siguiente procedimiento
    'Puede modificarse utilizando el Diseñador de Windows Forms. 
    'No lo modifique con el editor de código.
    Friend WithEvents cmdAccept As System.Windows.Forms.Button
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents txtDetail As System.Windows.Forms.TextBox
    Friend WithEvents chkDetails As System.Windows.Forms.CheckBox
    Friend WithEvents imgIconCritical As System.Windows.Forms.PictureBox
    Friend WithEvents imgIconWarning As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFatalError))
        Me.imgIconCritical = New System.Windows.Forms.PictureBox
        Me.lblMessage = New System.Windows.Forms.Label
        Me.cmdAccept = New System.Windows.Forms.Button
        Me.txtDetail = New System.Windows.Forms.TextBox
        Me.chkDetails = New System.Windows.Forms.CheckBox
        Me.imgIconWarning = New System.Windows.Forms.PictureBox
        Me.SuspendLayout()
        '
        'imgIconCritical
        '
        Me.imgIconCritical.AccessibleDescription = resources.GetString("imgIconCritical.AccessibleDescription")
        Me.imgIconCritical.AccessibleName = resources.GetString("imgIconCritical.AccessibleName")
        Me.imgIconCritical.Anchor = CType(resources.GetObject("imgIconCritical.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.imgIconCritical.BackgroundImage = CType(resources.GetObject("imgIconCritical.BackgroundImage"), System.Drawing.Image)
        Me.imgIconCritical.Dock = CType(resources.GetObject("imgIconCritical.Dock"), System.Windows.Forms.DockStyle)
        Me.imgIconCritical.Enabled = CType(resources.GetObject("imgIconCritical.Enabled"), Boolean)
        Me.imgIconCritical.Font = CType(resources.GetObject("imgIconCritical.Font"), System.Drawing.Font)
        Me.imgIconCritical.Image = CType(resources.GetObject("imgIconCritical.Image"), System.Drawing.Image)
        Me.imgIconCritical.ImeMode = CType(resources.GetObject("imgIconCritical.ImeMode"), System.Windows.Forms.ImeMode)
        Me.imgIconCritical.Location = CType(resources.GetObject("imgIconCritical.Location"), System.Drawing.Point)
        Me.imgIconCritical.Name = "imgIconCritical"
        Me.imgIconCritical.RightToLeft = CType(resources.GetObject("imgIconCritical.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.imgIconCritical.Size = CType(resources.GetObject("imgIconCritical.Size"), System.Drawing.Size)
        Me.imgIconCritical.SizeMode = CType(resources.GetObject("imgIconCritical.SizeMode"), System.Windows.Forms.PictureBoxSizeMode)
        Me.imgIconCritical.TabIndex = CType(resources.GetObject("imgIconCritical.TabIndex"), Integer)
        Me.imgIconCritical.TabStop = False
        Me.imgIconCritical.Text = resources.GetString("imgIconCritical.Text")
        Me.imgIconCritical.Visible = CType(resources.GetObject("imgIconCritical.Visible"), Boolean)
        '
        'lblMessage
        '
        Me.lblMessage.AccessibleDescription = resources.GetString("lblMessage.AccessibleDescription")
        Me.lblMessage.AccessibleName = resources.GetString("lblMessage.AccessibleName")
        Me.lblMessage.Anchor = CType(resources.GetObject("lblMessage.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.lblMessage.AutoSize = CType(resources.GetObject("lblMessage.AutoSize"), Boolean)
        Me.lblMessage.Dock = CType(resources.GetObject("lblMessage.Dock"), System.Windows.Forms.DockStyle)
        Me.lblMessage.Enabled = CType(resources.GetObject("lblMessage.Enabled"), Boolean)
        Me.lblMessage.Font = CType(resources.GetObject("lblMessage.Font"), System.Drawing.Font)
        Me.lblMessage.Image = CType(resources.GetObject("lblMessage.Image"), System.Drawing.Image)
        Me.lblMessage.ImageAlign = CType(resources.GetObject("lblMessage.ImageAlign"), System.Drawing.ContentAlignment)
        Me.lblMessage.ImageIndex = CType(resources.GetObject("lblMessage.ImageIndex"), Integer)
        Me.lblMessage.ImeMode = CType(resources.GetObject("lblMessage.ImeMode"), System.Windows.Forms.ImeMode)
        Me.lblMessage.Location = CType(resources.GetObject("lblMessage.Location"), System.Drawing.Point)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.RightToLeft = CType(resources.GetObject("lblMessage.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.lblMessage.Size = CType(resources.GetObject("lblMessage.Size"), System.Drawing.Size)
        Me.lblMessage.TabIndex = CType(resources.GetObject("lblMessage.TabIndex"), Integer)
        Me.lblMessage.Text = resources.GetString("lblMessage.Text")
        Me.lblMessage.TextAlign = CType(resources.GetObject("lblMessage.TextAlign"), System.Drawing.ContentAlignment)
        Me.lblMessage.Visible = CType(resources.GetObject("lblMessage.Visible"), Boolean)
        '
        'cmdAccept
        '
        Me.cmdAccept.AccessibleDescription = resources.GetString("cmdAccept.AccessibleDescription")
        Me.cmdAccept.AccessibleName = resources.GetString("cmdAccept.AccessibleName")
        Me.cmdAccept.Anchor = CType(resources.GetObject("cmdAccept.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.cmdAccept.BackgroundImage = CType(resources.GetObject("cmdAccept.BackgroundImage"), System.Drawing.Image)
        Me.cmdAccept.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdAccept.Dock = CType(resources.GetObject("cmdAccept.Dock"), System.Windows.Forms.DockStyle)
        Me.cmdAccept.Enabled = CType(resources.GetObject("cmdAccept.Enabled"), Boolean)
        Me.cmdAccept.FlatStyle = CType(resources.GetObject("cmdAccept.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.cmdAccept.Font = CType(resources.GetObject("cmdAccept.Font"), System.Drawing.Font)
        Me.cmdAccept.Image = CType(resources.GetObject("cmdAccept.Image"), System.Drawing.Image)
        Me.cmdAccept.ImageAlign = CType(resources.GetObject("cmdAccept.ImageAlign"), System.Drawing.ContentAlignment)
        Me.cmdAccept.ImageIndex = CType(resources.GetObject("cmdAccept.ImageIndex"), Integer)
        Me.cmdAccept.ImeMode = CType(resources.GetObject("cmdAccept.ImeMode"), System.Windows.Forms.ImeMode)
        Me.cmdAccept.Location = CType(resources.GetObject("cmdAccept.Location"), System.Drawing.Point)
        Me.cmdAccept.Name = "cmdAccept"
        Me.cmdAccept.RightToLeft = CType(resources.GetObject("cmdAccept.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.cmdAccept.Size = CType(resources.GetObject("cmdAccept.Size"), System.Drawing.Size)
        Me.cmdAccept.TabIndex = CType(resources.GetObject("cmdAccept.TabIndex"), Integer)
        Me.cmdAccept.Text = resources.GetString("cmdAccept.Text")
        Me.cmdAccept.TextAlign = CType(resources.GetObject("cmdAccept.TextAlign"), System.Drawing.ContentAlignment)
        Me.cmdAccept.Visible = CType(resources.GetObject("cmdAccept.Visible"), Boolean)
        '
        'txtDetail
        '
        Me.txtDetail.AccessibleDescription = resources.GetString("txtDetail.AccessibleDescription")
        Me.txtDetail.AccessibleName = resources.GetString("txtDetail.AccessibleName")
        Me.txtDetail.Anchor = CType(resources.GetObject("txtDetail.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtDetail.AutoSize = CType(resources.GetObject("txtDetail.AutoSize"), Boolean)
        Me.txtDetail.BackgroundImage = CType(resources.GetObject("txtDetail.BackgroundImage"), System.Drawing.Image)
        Me.txtDetail.Dock = CType(resources.GetObject("txtDetail.Dock"), System.Windows.Forms.DockStyle)
        Me.txtDetail.Enabled = CType(resources.GetObject("txtDetail.Enabled"), Boolean)
        Me.txtDetail.Font = CType(resources.GetObject("txtDetail.Font"), System.Drawing.Font)
        Me.txtDetail.ImeMode = CType(resources.GetObject("txtDetail.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtDetail.Location = CType(resources.GetObject("txtDetail.Location"), System.Drawing.Point)
        Me.txtDetail.MaxLength = CType(resources.GetObject("txtDetail.MaxLength"), Integer)
        Me.txtDetail.Multiline = CType(resources.GetObject("txtDetail.Multiline"), Boolean)
        Me.txtDetail.Name = "txtDetail"
        Me.txtDetail.PasswordChar = CType(resources.GetObject("txtDetail.PasswordChar"), Char)
        Me.txtDetail.RightToLeft = CType(resources.GetObject("txtDetail.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtDetail.ScrollBars = CType(resources.GetObject("txtDetail.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtDetail.Size = CType(resources.GetObject("txtDetail.Size"), System.Drawing.Size)
        Me.txtDetail.TabIndex = CType(resources.GetObject("txtDetail.TabIndex"), Integer)
        Me.txtDetail.Text = resources.GetString("txtDetail.Text")
        Me.txtDetail.TextAlign = CType(resources.GetObject("txtDetail.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtDetail.Visible = CType(resources.GetObject("txtDetail.Visible"), Boolean)
        Me.txtDetail.WordWrap = CType(resources.GetObject("txtDetail.WordWrap"), Boolean)
        '
        'chkDetails
        '
        Me.chkDetails.AccessibleDescription = resources.GetString("chkDetails.AccessibleDescription")
        Me.chkDetails.AccessibleName = resources.GetString("chkDetails.AccessibleName")
        Me.chkDetails.Anchor = CType(resources.GetObject("chkDetails.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.chkDetails.Appearance = CType(resources.GetObject("chkDetails.Appearance"), System.Windows.Forms.Appearance)
        Me.chkDetails.BackgroundImage = CType(resources.GetObject("chkDetails.BackgroundImage"), System.Drawing.Image)
        Me.chkDetails.CheckAlign = CType(resources.GetObject("chkDetails.CheckAlign"), System.Drawing.ContentAlignment)
        Me.chkDetails.Dock = CType(resources.GetObject("chkDetails.Dock"), System.Windows.Forms.DockStyle)
        Me.chkDetails.Enabled = CType(resources.GetObject("chkDetails.Enabled"), Boolean)
        Me.chkDetails.FlatStyle = CType(resources.GetObject("chkDetails.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.chkDetails.Font = CType(resources.GetObject("chkDetails.Font"), System.Drawing.Font)
        Me.chkDetails.Image = CType(resources.GetObject("chkDetails.Image"), System.Drawing.Image)
        Me.chkDetails.ImageAlign = CType(resources.GetObject("chkDetails.ImageAlign"), System.Drawing.ContentAlignment)
        Me.chkDetails.ImageIndex = CType(resources.GetObject("chkDetails.ImageIndex"), Integer)
        Me.chkDetails.ImeMode = CType(resources.GetObject("chkDetails.ImeMode"), System.Windows.Forms.ImeMode)
        Me.chkDetails.Location = CType(resources.GetObject("chkDetails.Location"), System.Drawing.Point)
        Me.chkDetails.Name = "chkDetails"
        Me.chkDetails.RightToLeft = CType(resources.GetObject("chkDetails.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.chkDetails.Size = CType(resources.GetObject("chkDetails.Size"), System.Drawing.Size)
        Me.chkDetails.TabIndex = CType(resources.GetObject("chkDetails.TabIndex"), Integer)
        Me.chkDetails.Text = resources.GetString("chkDetails.Text")
        Me.chkDetails.TextAlign = CType(resources.GetObject("chkDetails.TextAlign"), System.Drawing.ContentAlignment)
        Me.chkDetails.Visible = CType(resources.GetObject("chkDetails.Visible"), Boolean)
        '
        'imgIconWarning
        '
        Me.imgIconWarning.AccessibleDescription = resources.GetString("imgIconWarning.AccessibleDescription")
        Me.imgIconWarning.AccessibleName = resources.GetString("imgIconWarning.AccessibleName")
        Me.imgIconWarning.Anchor = CType(resources.GetObject("imgIconWarning.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.imgIconWarning.BackgroundImage = CType(resources.GetObject("imgIconWarning.BackgroundImage"), System.Drawing.Image)
        Me.imgIconWarning.Dock = CType(resources.GetObject("imgIconWarning.Dock"), System.Windows.Forms.DockStyle)
        Me.imgIconWarning.Enabled = CType(resources.GetObject("imgIconWarning.Enabled"), Boolean)
        Me.imgIconWarning.Font = CType(resources.GetObject("imgIconWarning.Font"), System.Drawing.Font)
        Me.imgIconWarning.Image = CType(resources.GetObject("imgIconWarning.Image"), System.Drawing.Image)
        Me.imgIconWarning.ImeMode = CType(resources.GetObject("imgIconWarning.ImeMode"), System.Windows.Forms.ImeMode)
        Me.imgIconWarning.Location = CType(resources.GetObject("imgIconWarning.Location"), System.Drawing.Point)
        Me.imgIconWarning.Name = "imgIconWarning"
        Me.imgIconWarning.RightToLeft = CType(resources.GetObject("imgIconWarning.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.imgIconWarning.Size = CType(resources.GetObject("imgIconWarning.Size"), System.Drawing.Size)
        Me.imgIconWarning.SizeMode = CType(resources.GetObject("imgIconWarning.SizeMode"), System.Windows.Forms.PictureBoxSizeMode)
        Me.imgIconWarning.TabIndex = CType(resources.GetObject("imgIconWarning.TabIndex"), Integer)
        Me.imgIconWarning.TabStop = False
        Me.imgIconWarning.Text = resources.GetString("imgIconWarning.Text")
        Me.imgIconWarning.Visible = CType(resources.GetObject("imgIconWarning.Visible"), Boolean)
        '
        'frmFatalError
        '
        Me.AcceptButton = Me.cmdAccept
        Me.AccessibleDescription = resources.GetString("$this.AccessibleDescription")
        Me.AccessibleName = resources.GetString("$this.AccessibleName")
        Me.AutoScaleBaseSize = CType(resources.GetObject("$this.AutoScaleBaseSize"), System.Drawing.Size)
        Me.AutoScroll = CType(resources.GetObject("$this.AutoScroll"), Boolean)
        Me.AutoScrollMargin = CType(resources.GetObject("$this.AutoScrollMargin"), System.Drawing.Size)
        Me.AutoScrollMinSize = CType(resources.GetObject("$this.AutoScrollMinSize"), System.Drawing.Size)
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.CancelButton = Me.cmdAccept
        Me.ClientSize = CType(resources.GetObject("$this.ClientSize"), System.Drawing.Size)
        Me.ControlBox = False
        Me.Controls.Add(Me.chkDetails)
        Me.Controls.Add(Me.txtDetail)
        Me.Controls.Add(Me.cmdAccept)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.imgIconCritical)
        Me.Controls.Add(Me.imgIconWarning)
        Me.Enabled = CType(resources.GetObject("$this.Enabled"), Boolean)
        Me.Font = CType(resources.GetObject("$this.Font"), System.Drawing.Font)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.ImeMode = CType(resources.GetObject("$this.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Location = CType(resources.GetObject("$this.Location"), System.Drawing.Point)
        Me.MaximumSize = CType(resources.GetObject("$this.MaximumSize"), System.Drawing.Size)
        Me.MinimumSize = CType(resources.GetObject("$this.MinimumSize"), System.Drawing.Size)
        Me.Name = "frmFatalError"
        Me.RightToLeft = CType(resources.GetObject("$this.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.StartPosition = CType(resources.GetObject("$this.StartPosition"), System.Windows.Forms.FormStartPosition)
        Me.Text = resources.GetString("$this.Text")
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region


    Public Sub ShowCritical(ByVal ex As Exception, Optional ByVal pvstr_Mensaje As String = "")
        Me.Text = "Error fatal"
        imgIconCritical.Visible = True
        imgIconWarning.Visible = False
        ShowMe(ex, pvstr_Mensaje)
    End Sub

    Public Sub ShowWarning(ByVal ex As Exception, Optional ByVal pvstr_Mensaje As String = "")
        Me.Text = "Atencion!"
        imgIconCritical.Visible = False
        imgIconWarning.Visible = True
        ShowMe(ex, pvstr_Mensaje)
    End Sub

    Private Sub ShowMe(ByVal ex As Exception, Optional ByVal pvstr_Mensaje As String = "")
        If Not pvstr_Mensaje Is Nothing AndAlso pvstr_Mensaje.Length() > 0 Then
            lblMessage.Text = pvstr_Mensaje
        Else
            lblMessage.Text = ex.Message
        End If

        Dim loEx As Exception = ex
        Dim lvstrBuff As String = ""
        Do While Not (loEx Is Nothing)
            lvstrBuff &= "Error message: " & loEx.Message & vbCrLf & _
                         "Exception: " & TypeName(loEx) & vbCrLf & _
                         "Source: " & loEx.Source & vbCrLf & _
                         New String("-", 80) & vbCrLf & _
                         loEx.StackTrace & vbCrLf
            loEx = loEx.InnerException
            If Not (loEx Is Nothing) Then
                lvstrBuff &= vbCrLf & New String("=", 80) & vbCrLf & vbCrLf & "SOURCE EXCEPTION:" & vbCrLf & vbCrLf
            End If
        Loop

        txtDetail.Text = lvstrBuff
        ShowDialog()
    End Sub


    Private Sub SetSize()
        Dim lvintScreenWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
        If mvblnDetails Then
            Me.MinimumSize = New Size(300, 240)
            Me.MaximumSize = New Size(0, 0)
            Me.Height = 400
        Else
            Me.MinimumSize = New Size(300, 120)
            Me.MaximumSize = New Size(lvintScreenWidth, 120)
            Me.Height = 104
        End If
    End Sub

    Private Sub chkDetails_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDetails.CheckedChanged
        mvblnDetails = chkDetails.Checked
        SetSize()
    End Sub

    Private Sub cmdAccept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAccept.Click
        Me.Close()
    End Sub

End Class
