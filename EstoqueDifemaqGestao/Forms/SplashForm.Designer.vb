<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SplashForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.pnlMain = New System.Windows.Forms.Panel()
        Me.pnlHeader = New System.Windows.Forms.Panel()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblSubtitle = New System.Windows.Forms.Label()
        Me.pbLogo = New System.Windows.Forms.PictureBox()
        Me.pnlContent = New System.Windows.Forms.Panel()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.progressBar = New System.Windows.Forms.ProgressBar()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.pnlFooter = New System.Windows.Forms.Panel()
        Me.lblCopyright = New System.Windows.Forms.Label()
        Me.pnlMain.SuspendLayout()
        Me.pnlHeader.SuspendLayout()
        CType(Me.pbLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlContent.SuspendLayout()
        Me.pnlFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlMain
        '
        Me.pnlMain.BackColor = System.Drawing.Color.White
        Me.pnlMain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlMain.Controls.Add(Me.pnlHeader)
        Me.pnlMain.Controls.Add(Me.pnlContent)
        Me.pnlMain.Controls.Add(Me.pnlFooter)
        Me.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMain.Location = New System.Drawing.Point(0, 0)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Size = New System.Drawing.Size(500, 280)
        Me.pnlMain.TabIndex = 0
        '
        'pnlHeader
        '
        Me.pnlHeader.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(134, Byte), Integer), CType(CType(171, Byte), Integer))
        Me.pnlHeader.Controls.Add(Me.lblTitle)
        Me.pnlHeader.Controls.Add(Me.lblSubtitle)
        Me.pnlHeader.Controls.Add(Me.pbLogo)
        Me.pnlHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlHeader.Name = "pnlHeader"
        Me.pnlHeader.Size = New System.Drawing.Size(498, 100)
        Me.pnlHeader.TabIndex = 0
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Segoe UI", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.White
        Me.lblTitle.Location = New System.Drawing.Point(110, 20)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(191, 30)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Gestão de Estoque"
        '
        'lblSubtitle
        '
        Me.lblSubtitle.AutoSize = True
        Me.lblSubtitle.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSubtitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(230, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblSubtitle.Location = New System.Drawing.Point(112, 55)
        Me.lblSubtitle.Name = "lblSubtitle"
        Me.lblSubtitle.Size = New System.Drawing.Size(118, 20)
        Me.lblSubtitle.TabIndex = 1
        Me.lblSubtitle.Text = "DIFEMAQ LTDA"
        '
        'pbLogo
        '
        Me.pbLogo.BackColor = System.Drawing.Color.FromArgb(CType(CType(80, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.pbLogo.Location = New System.Drawing.Point(20, 20)
        Me.pbLogo.Name = "pbLogo"
        Me.pbLogo.Size = New System.Drawing.Size(70, 60)
        Me.pbLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pbLogo.TabIndex = 2
        Me.pbLogo.TabStop = False
        '
        'pnlContent
        '
        Me.pnlContent.BackColor = System.Drawing.Color.White
        Me.pnlContent.Controls.Add(Me.lblStatus)
        Me.pnlContent.Controls.Add(Me.progressBar)
        Me.pnlContent.Controls.Add(Me.lblVersion)
        Me.pnlContent.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlContent.Location = New System.Drawing.Point(0, 100)
        Me.pnlContent.Name = "pnlContent"
        Me.pnlContent.Size = New System.Drawing.Size(498, 135)
        Me.pnlContent.TabIndex = 1
        '
        'lblStatus
        '
        Me.lblStatus.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblStatus.Location = New System.Drawing.Point(20, 30)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(460, 25)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Iniciando sistema..."
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'progressBar
        '
        Me.progressBar.Location = New System.Drawing.Point(40, 70)
        Me.progressBar.Maximum = 100
        Me.progressBar.Name = "progressBar"
        Me.progressBar.Size = New System.Drawing.Size(420, 18)
        Me.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.progressBar.TabIndex = 1
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.Color.Gray
        Me.lblVersion.Location = New System.Drawing.Point(20, 105)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(64, 13)
        Me.lblVersion.TabIndex = 2
        Me.lblVersion.Text = "Versão 1.0.0"
        '
        'pnlFooter
        '
        Me.pnlFooter.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.pnlFooter.Controls.Add(Me.lblCopyright)
        Me.pnlFooter.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFooter.Location = New System.Drawing.Point(0, 235)
        Me.pnlFooter.Name = "pnlFooter"
        Me.pnlFooter.Size = New System.Drawing.Size(498, 45)
        Me.pnlFooter.TabIndex = 2
        '
        'lblCopyright
        '
        Me.lblCopyright.AutoSize = True
        Me.lblCopyright.Font = New System.Drawing.Font("Segoe UI", 7.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCopyright.ForeColor = System.Drawing.Color.Gray
        Me.lblCopyright.Location = New System.Drawing.Point(20, 15)
        Me.lblCopyright.Name = "lblCopyright"
        Me.lblCopyright.Size = New System.Drawing.Size(206, 13)
        Me.lblCopyright.TabIndex = 0
        Me.lblCopyright.Text = "© 2024 Difemaq Ltda. Todos os direitos reservados."
        '
        'SplashForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(500, 280)
        Me.Controls.Add(Me.pnlMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "SplashForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Carregando..."
        Me.TopMost = True
        Me.pnlMain.ResumeLayout(False)
        Me.pnlHeader.ResumeLayout(False)
        Me.pnlHeader.PerformLayout()
        CType(Me.pbLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlContent.ResumeLayout(False)
        Me.pnlContent.PerformLayout()
        Me.pnlFooter.ResumeLayout(False)
        Me.pnlFooter.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlMain As Panel
    Friend WithEvents pnlHeader As Panel
    Friend WithEvents lblTitle As Label
    Friend WithEvents lblSubtitle As Label
    Friend WithEvents pbLogo As PictureBox
    Friend WithEvents pnlContent As Panel
    Friend WithEvents lblStatus As Label
    Friend WithEvents progressBar As ProgressBar
    Friend WithEvents lblVersion As Label
    Friend WithEvents pnlFooter As Panel
    Friend WithEvents lblCopyright As Label
End Class