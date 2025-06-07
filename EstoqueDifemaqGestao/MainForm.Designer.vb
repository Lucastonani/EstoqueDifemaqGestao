<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private Sub InitializeComponent()
        Me.pnlTitulo = New System.Windows.Forms.Panel()
        Me.lblTitulo = New System.Windows.Forms.Label()
        Me.pnlControlesJanela = New System.Windows.Forms.Panel()
        Me.btnFechar = New System.Windows.Forms.Button()
        Me.btnMaximizar = New System.Windows.Forms.Button()
        Me.btnMinimizar = New System.Windows.Forms.Button()
        Me.pnlMenu = New System.Windows.Forms.Panel()
        Me.btnConfiguracoes = New System.Windows.Forms.Button()
        Me.btnRelatorios = New System.Windows.Forms.Button()
        Me.btnReposicaoEstoque = New System.Windows.Forms.Button()
        Me.pnlConteudo = New System.Windows.Forms.Panel()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.lblStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblDataHora = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pnlTitulo.SuspendLayout()
        Me.pnlControlesJanela.SuspendLayout()
        Me.pnlMenu.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()

        ' pnlTitulo
        Me.pnlTitulo.BackColor = System.Drawing.Color.FromArgb(46, 134, 171)
        Me.pnlTitulo.Controls.Add(Me.lblTitulo)
        Me.pnlTitulo.Controls.Add(Me.pnlControlesJanela)
        Me.pnlTitulo.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTitulo.Location = New System.Drawing.Point(0, 0)
        Me.pnlTitulo.Name = "pnlTitulo"
        Me.pnlTitulo.Size = New System.Drawing.Size(1400, 50)
        Me.pnlTitulo.TabIndex = 0

        ' lblTitulo
        Me.lblTitulo.AutoSize = True
        Me.lblTitulo.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitulo.ForeColor = System.Drawing.Color.White
        Me.lblTitulo.Location = New System.Drawing.Point(15, 12)
        Me.lblTitulo.Name = "lblTitulo"
        Me.lblTitulo.Size = New System.Drawing.Size(302, 25)
        Me.lblTitulo.TabIndex = 0
        Me.lblTitulo.Text = "Gestão de Estoque - Difemaq"

        ' pnlControlesJanela
        Me.pnlControlesJanela.Controls.Add(Me.btnFechar)
        Me.pnlControlesJanela.Controls.Add(Me.btnMaximizar)
        Me.pnlControlesJanela.Controls.Add(Me.btnMinimizar)
        Me.pnlControlesJanela.Dock = System.Windows.Forms.DockStyle.Right
        Me.pnlControlesJanela.Location = New System.Drawing.Point(1265, 0)
        Me.pnlControlesJanela.Name = "pnlControlesJanela"
        Me.pnlControlesJanela.Size = New System.Drawing.Size(135, 50)
        Me.pnlControlesJanela.TabIndex = 1

        ' btnFechar
        Me.btnFechar.BackColor = System.Drawing.Color.Transparent
        Me.btnFechar.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnFechar.FlatAppearance.BorderSize = 0
        Me.btnFechar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(232, 17, 35)
        Me.btnFechar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(232, 17, 35)
        Me.btnFechar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFechar.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.btnFechar.ForeColor = System.Drawing.Color.White
        Me.btnFechar.Location = New System.Drawing.Point(90, 0)
        Me.btnFechar.Name = "btnFechar"
        Me.btnFechar.Size = New System.Drawing.Size(45, 50)
        Me.btnFechar.TabIndex = 2
        Me.btnFechar.Text = "✕"
        Me.btnFechar.UseVisualStyleBackColor = False

        ' btnMaximizar
        Me.btnMaximizar.BackColor = System.Drawing.Color.Transparent
        Me.btnMaximizar.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnMaximizar.FlatAppearance.BorderSize = 0
        Me.btnMaximizar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(80, 150, 190)
        Me.btnMaximizar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(80, 150, 190)
        Me.btnMaximizar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMaximizar.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.btnMaximizar.ForeColor = System.Drawing.Color.White
        Me.btnMaximizar.Location = New System.Drawing.Point(45, 0)
        Me.btnMaximizar.Name = "btnMaximizar"
        Me.btnMaximizar.Size = New System.Drawing.Size(45, 50)
        Me.btnMaximizar.TabIndex = 1
        Me.btnMaximizar.Text = "□"
        Me.btnMaximizar.UseVisualStyleBackColor = False

        ' btnMinimizar
        Me.btnMinimizar.BackColor = System.Drawing.Color.Transparent
        Me.btnMinimizar.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnMinimizar.FlatAppearance.BorderSize = 0
        Me.btnMinimizar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(80, 150, 190)
        Me.btnMinimizar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(80, 150, 190)
        Me.btnMinimizar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMinimizar.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.btnMinimizar.ForeColor = System.Drawing.Color.White
        Me.btnMinimizar.Location = New System.Drawing.Point(0, 0)
        Me.btnMinimizar.Name = "btnMinimizar"
        Me.btnMinimizar.Size = New System.Drawing.Size(45, 50)
        Me.btnMinimizar.TabIndex = 0
        Me.btnMinimizar.Text = "_"
        Me.btnMinimizar.UseVisualStyleBackColor = False

        ' pnlMenu
        Me.pnlMenu.BackColor = System.Drawing.Color.FromArgb(240, 240, 240)
        Me.pnlMenu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlMenu.Controls.Add(Me.btnConfiguracoes)
        Me.pnlMenu.Controls.Add(Me.btnRelatorios)
        Me.pnlMenu.Controls.Add(Me.btnReposicaoEstoque)
        Me.pnlMenu.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMenu.Location = New System.Drawing.Point(0, 50)
        Me.pnlMenu.Name = "pnlMenu"
        Me.pnlMenu.Size = New System.Drawing.Size(1400, 60)
        Me.pnlMenu.TabIndex = 1

        ' btnReposicaoEstoque
        Me.btnReposicaoEstoque.BackColor = System.Drawing.Color.Transparent
        Me.btnReposicaoEstoque.FlatAppearance.BorderSize = 0
        Me.btnReposicaoEstoque.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnReposicaoEstoque.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnReposicaoEstoque.Location = New System.Drawing.Point(15, 10)
        Me.btnReposicaoEstoque.Name = "btnReposicaoEstoque"
        Me.btnReposicaoEstoque.Size = New System.Drawing.Size(150, 40)
        Me.btnReposicaoEstoque.TabIndex = 0
        Me.btnReposicaoEstoque.Text = "Reposição Estoque"
        Me.btnReposicaoEstoque.UseVisualStyleBackColor = False

        ' btnRelatorios
        Me.btnRelatorios.BackColor = System.Drawing.Color.Transparent
        Me.btnRelatorios.FlatAppearance.BorderSize = 0
        Me.btnRelatorios.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRelatorios.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnRelatorios.Location = New System.Drawing.Point(180, 10)
        Me.btnRelatorios.Name = "btnRelatorios"
        Me.btnRelatorios.Size = New System.Drawing.Size(100, 40)
        Me.btnRelatorios.TabIndex = 1
        Me.btnRelatorios.Text = "Relatórios"
        Me.btnRelatorios.UseVisualStyleBackColor = False

        ' btnConfiguracoes
        Me.btnConfiguracoes.BackColor = System.Drawing.Color.Transparent
        Me.btnConfiguracoes.FlatAppearance.BorderSize = 0
        Me.btnConfiguracoes.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnConfiguracoes.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnConfiguracoes.Location = New System.Drawing.Point(295, 10)
        Me.btnConfiguracoes.Name = "btnConfiguracoes"
        Me.btnConfiguracoes.Size = New System.Drawing.Size(120, 40)
        Me.btnConfiguracoes.TabIndex = 2
        Me.btnConfiguracoes.Text = "Configurações"
        Me.btnConfiguracoes.UseVisualStyleBackColor = False

        ' pnlConteudo
        Me.pnlConteudo.BackColor = System.Drawing.Color.White
        Me.pnlConteudo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlConteudo.Location = New System.Drawing.Point(0, 110)
        Me.pnlConteudo.Name = "pnlConteudo"
        Me.pnlConteudo.Size = New System.Drawing.Size(1400, 740)
        Me.pnlConteudo.TabIndex = 2

        ' StatusStrip
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblStatus, Me.lblDataHora})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 828)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(1400, 22)
        Me.StatusStrip.TabIndex = 3

        ' lblStatus
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(42, 17)
        Me.lblStatus.Text = "Pronto"

        ' lblDataHora
        Me.lblDataHora.Name = "lblDataHora"
        Me.lblDataHora.Size = New System.Drawing.Size(1343, 17)
        Me.lblDataHora.Spring = True
        Me.lblDataHora.Text = "00/00/0000 00:00:00"
        Me.lblDataHora.TextAlign = System.Drawing.ContentAlignment.MiddleRight

        ' MainForm
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1400, 850)
        Me.Controls.Add(Me.pnlConteudo)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.pnlMenu)
        Me.Controls.Add(Me.pnlTitulo)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "MainForm"
        Me.Text = "Gestão de Estoque - Difemaq"
        Me.pnlTitulo.ResumeLayout(False)
        Me.pnlTitulo.PerformLayout()
        Me.pnlControlesJanela.ResumeLayout(False)
        Me.pnlMenu.ResumeLayout(False)
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents pnlTitulo As Panel
    Friend WithEvents lblTitulo As Label
    Friend WithEvents pnlControlesJanela As Panel
    Friend WithEvents btnFechar As Button
    Friend WithEvents btnMaximizar As Button
    Friend WithEvents btnMinimizar As Button
    Friend WithEvents pnlMenu As Panel
    Friend WithEvents btnReposicaoEstoque As Button
    Friend WithEvents btnRelatorios As Button
    Friend WithEvents btnConfiguracoes As Button
    Friend WithEvents pnlConteudo As Panel
    Friend WithEvents StatusStrip As StatusStrip
    Friend WithEvents lblStatus As ToolStripStatusLabel
    Friend WithEvents lblDataHora As ToolStripStatusLabel
End Class