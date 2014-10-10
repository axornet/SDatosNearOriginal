Public Class frmSplash
    Inherits System.Windows.Forms.Form

#Region " Código generado por el Diseñador de Windows Forms "

    Public Sub New()
        MyBase.New()

        'El Diseñador de Windows Forms requiere esta llamada.
        InitializeComponent()

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
    Friend WithEvents lblFecha As System.Windows.Forms.Label
    Friend WithEvents lblAppVersion As System.Windows.Forms.Label
    Friend WithEvents lblAppName As System.Windows.Forms.Label
    Friend WithEvents cmdCerrar As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSplash))
        Me.lblFecha = New System.Windows.Forms.Label()
        Me.lblAppVersion = New System.Windows.Forms.Label()
        Me.lblAppName = New System.Windows.Forms.Label()
        Me.cmdCerrar = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblFecha
        '
        resources.ApplyResources(Me.lblFecha, "lblFecha")
        Me.lblFecha.ForeColor = System.Drawing.Color.Maroon
        Me.lblFecha.Name = "lblFecha"
        '
        'lblAppVersion
        '
        resources.ApplyResources(Me.lblAppVersion, "lblAppVersion")
        Me.lblAppVersion.ForeColor = System.Drawing.Color.Maroon
        Me.lblAppVersion.Name = "lblAppVersion"
        '
        'lblAppName
        '
        resources.ApplyResources(Me.lblAppName, "lblAppName")
        Me.lblAppName.ForeColor = System.Drawing.Color.Maroon
        Me.lblAppName.Name = "lblAppName"
        '
        'cmdCerrar
        '
        Me.cmdCerrar.BackColor = System.Drawing.Color.Khaki
        Me.cmdCerrar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        resources.ApplyResources(Me.cmdCerrar, "cmdCerrar")
        Me.cmdCerrar.Name = "cmdCerrar"
        Me.cmdCerrar.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        resources.ApplyResources(Me.PictureBox1, "PictureBox1")
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        resources.ApplyResources(Me.PictureBox2, "PictureBox2")
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.TabStop = False
        '
        'frmSplash
        '
        Me.AcceptButton = Me.cmdCerrar
        resources.ApplyResources(Me, "$this")
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.cmdCerrar
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdCerrar)
        Me.Controls.Add(Me.lblFecha)
        Me.Controls.Add(Me.lblAppVersion)
        Me.Controls.Add(Me.lblAppName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Name = "frmSplash"
        Me.ShowInTaskbar = False
        Me.TopMost = True
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Shared mofrmSplash As frmSplash

    Public Shared Sub ShowAbout()
        Dim loFrm As New frmSplash
        loFrm.cmdCerrar.Visible = True
        loFrm.ShowDialog(goMainForm)
    End Sub

    Public Shared Sub ShowSplash()
        If mofrmSplash Is Nothing Then mofrmSplash = New frmSplash
        mofrmSplash.cmdCerrar.Visible = False
        mofrmSplash.Show()
        Application.DoEvents()
    End Sub

    Public Shared Sub HideSplash()
        If Not mofrmSplash Is Nothing Then
            mofrmSplash.Close()
            mofrmSplash = Nothing
            Application.DoEvents()
        End If
    End Sub

    Public Shared Function ExistsSplash() As Boolean
        Return Not (mofrmSplash Is Nothing)
    End Function

    Private Sub frmSplash_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim lvdatDate As DateTime
        lblAppName.Text = "Datamart Services"
        lblAppVersion.Text = "Version 1.0"
        lvdatDate = FileDateTime(System.Reflection.Assembly.GetExecutingAssembly.Location)
        lblFecha.Text = String.Format("{0:d}  -  {0:T}", lvdatDate)
    End Sub


    Private Sub frmSplash_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        'Dibujo un borde 3D al formulario
        System.Windows.Forms.ControlPaint.DrawBorder3D( _
                System.Drawing.Graphics.FromHwnd(Me.Handle), _
                New System.Drawing.Rectangle(0, 0, Me.Width, Me.Height))
    End Sub

    Private Sub frmSplash_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Dim i As Single
        For i = 1 To 0 Step -0.1
            Me.Opacity = i
            Application.DoEvents()
            System.Threading.Thread.Sleep(100)
        Next
    End Sub

    Private Sub cmdCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCerrar.Click
        Me.Close()
    End Sub

End Class
