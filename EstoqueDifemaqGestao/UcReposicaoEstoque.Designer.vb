<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class UcReposicaoEstoque
    Inherits System.Windows.Forms.UserControl

    'UserControl substitui o descarte para limpar a lista de componentes.
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

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.pnlPrincipal = New System.Windows.Forms.Panel()
        Me.splitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.pnlEsquerda = New System.Windows.Forms.Panel()
        Me.grpProdutos = New System.Windows.Forms.GroupBox()
        Me.dgvProdutos = New System.Windows.Forms.DataGridView()
        Me.pnlFiltros = New System.Windows.Forms.Panel()
        Me.txtFiltro = New System.Windows.Forms.TextBox()
        Me.lblFiltro = New System.Windows.Forms.Label()
        Me.btnAtualizar = New System.Windows.Forms.Button()
        Me.pnlDireita = New System.Windows.Forms.Panel()
        Me.splitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.pnlSuperior = New System.Windows.Forms.Panel()
        Me.splitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.grpImagem = New System.Windows.Forms.GroupBox()
        Me.pbProduto = New System.Windows.Forms.PictureBox()
        Me.grpEstoque = New System.Windows.Forms.GroupBox()
        Me.dgvEstoque = New System.Windows.Forms.DataGridView()
        Me.pnlInferior = New System.Windows.Forms.Panel()
        Me.splitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.grpCompras = New System.Windows.Forms.GroupBox()
        Me.dgvCompras = New System.Windows.Forms.DataGridView()
        Me.grpVendas = New System.Windows.Forms.GroupBox()
        Me.dgvVendas = New System.Windows.Forms.DataGridView()
        Me.pnlPrincipal.SuspendLayout()
        CType(Me.splitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitContainer1.Panel1.SuspendLayout()
        Me.splitContainer1.Panel2.SuspendLayout()
        Me.splitContainer1.SuspendLayout()
        Me.pnlEsquerda.SuspendLayout()
        Me.grpProdutos.SuspendLayout()
        CType(Me.dgvProdutos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFiltros.SuspendLayout()
        Me.pnlDireita.SuspendLayout()
        CType(Me.splitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitContainer2.Panel1.SuspendLayout()
        Me.splitContainer2.Panel2.SuspendLayout()
        Me.splitContainer2.SuspendLayout()
        Me.pnlSuperior.SuspendLayout()
        CType(Me.splitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitContainer3.Panel1.SuspendLayout()
        Me.splitContainer3.Panel2.SuspendLayout()
        Me.splitContainer3.SuspendLayout()
        Me.grpImagem.SuspendLayout()
        CType(Me.pbProduto, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpEstoque.SuspendLayout()
        CType(Me.dgvEstoque, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlInferior.SuspendLayout()
        CType(Me.splitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitContainer4.Panel1.SuspendLayout()
        Me.splitContainer4.Panel2.SuspendLayout()
        Me.splitContainer4.SuspendLayout()
        Me.grpCompras.SuspendLayout()
        CType(Me.dgvCompras, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpVendas.SuspendLayout()
        CType(Me.dgvVendas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        
        '
        'pnlPrincipal
        '
        Me.pnlPrincipal.Controls.Add(Me.splitContainer1)
        Me.pnlPrincipal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlPrincipal.Location = New System.Drawing.Point(0, 0)
        Me.pnlPrincipal.Name = "pnlPrincipal"
        Me.pnlPrincipal.Padding = New System.Windows.Forms.Padding(10)
        Me.pnlPrincipal.Size = New System.Drawing.Size(1150, 702)
        Me.pnlPrincipal.TabIndex = 0
        
        '
        'splitContainer1
        '
        Me.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.splitContainer1.Location = New System.Drawing.Point(10, 10)
        Me.splitContainer1.Name = "splitContainer1"
        Me.splitContainer1.Panel1MinSize = 350
        Me.splitContainer1.Panel2MinSize = 500
        '
        'splitContainer1.Panel1
        '
        Me.splitContainer1.Panel1.Controls.Add(Me.pnlEsquerda)
        '
        'splitContainer1.Panel2
        '
        Me.splitContainer1.Panel2.Controls.Add(Me.pnlDireita)
        Me.splitContainer1.Size = New System.Drawing.Size(1130, 682)
        Me.splitContainer1.SplitterDistance = 400
        Me.splitContainer1.SplitterWidth = 8
        Me.splitContainer1.TabIndex = 0
        
        '
        'pnlEsquerda
        '
        Me.pnlEsquerda.Controls.Add(Me.grpProdutos)
        Me.pnlEsquerda.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlEsquerda.Location = New System.Drawing.Point(0, 0)
        Me.pnlEsquerda.Name = "pnlEsquerda"
        Me.pnlEsquerda.Size = New System.Drawing.Size(400, 682)
        Me.pnlEsquerda.TabIndex = 0
        
        '
        'grpProdutos
        '
        Me.grpProdutos.Controls.Add(Me.dgvProdutos)
        Me.grpProdutos.Controls.Add(Me.pnlFiltros)
        Me.grpProdutos.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpProdutos.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpProdutos.Location = New System.Drawing.Point(0, 0)
        Me.grpProdutos.Name = "grpProdutos"
        Me.grpProdutos.Padding = New System.Windows.Forms.Padding(8)
        Me.grpProdutos.Size = New System.Drawing.Size(400, 682)
        Me.grpProdutos.TabIndex = 0
        Me.grpProdutos.TabStop = False
        Me.grpProdutos.Text = "📦 Lista de Produtos"
        
        '
        'dgvProdutos
        '
        Me.dgvProdutos.AllowUserToAddRows = False
        Me.dgvProdutos.AllowUserToDeleteRows = False
        Me.dgvProdutos.BackgroundColor = System.Drawing.Color.White
        Me.dgvProdutos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgvProdutos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProdutos.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvProdutos.Location = New System.Drawing.Point(8, 73)
        Me.dgvProdutos.MultiSelect = False
        Me.dgvProdutos.Name = "dgvProdutos"
        Me.dgvProdutos.ReadOnly = True
        Me.dgvProdutos.RowHeadersVisible = False
        Me.dgvProdutos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvProdutos.Size = New System.Drawing.Size(384, 601)
        Me.dgvProdutos.TabIndex = 1
        
        '
        'pnlFiltros
        '
        Me.pnlFiltros.Controls.Add(Me.txtFiltro)
        Me.pnlFiltros.Controls.Add(Me.lblFiltro)
        Me.pnlFiltros.Controls.Add(Me.btnAtualizar)
        Me.pnlFiltros.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFiltros.Location = New System.Drawing.Point(8, 26)
        Me.pnlFiltros.Name = "pnlFiltros"
        Me.pnlFiltros.Padding = New System.Windows.Forms.Padding(0, 8, 0, 8)
        Me.pnlFiltros.Size = New System.Drawing.Size(384, 47)
        Me.pnlFiltros.TabIndex = 0
        
        '
        'txtFiltro
        '
        Me.txtFiltro.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFiltro.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtFiltro.Location = New System.Drawing.Point(50, 12)
        Me.txtFiltro.Name = "txtFiltro"
        Me.txtFiltro.Size = New System.Drawing.Size(240, 23)
        Me.txtFiltro.TabIndex = 0
        
        '
        'lblFiltro
        '
        Me.lblFiltro.AutoSize = True
        Me.lblFiltro.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblFiltro.Location = New System.Drawing.Point(3, 15)
        Me.lblFiltro.Name = "lblFiltro"
        Me.lblFiltro.Size = New System.Drawing.Size(39, 15)
        Me.lblFiltro.TabIndex = 1
        Me.lblFiltro.Text = "Filtro:"
        
        '
        'btnAtualizar
        '
        Me.btnAtualizar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAtualizar.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(134, Byte), Integer), CType(CType(171, Byte), Integer))
        Me.btnAtualizar.FlatAppearance.BorderSize = 0
        Me.btnAtualizar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAtualizar.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Bold)
        Me.btnAtualizar.ForeColor = System.Drawing.Color.White
        Me.btnAtualizar.Location = New System.Drawing.Point(300, 10)
        Me.btnAtualizar.Name = "btnAtualizar"
        Me.btnAtualizar.Size = New System.Drawing.Size(80, 27)
        Me.btnAtualizar.TabIndex = 2
        Me.btnAtualizar.Text = "🔄 Atualizar"
        Me.btnAtualizar.UseVisualStyleBackColor = False
        
        '
        'pnlDireita
        '
        Me.pnlDireita.Controls.Add(Me.splitContainer2)
        Me.pnlDireita.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlDireita.Location = New System.Drawing.Point(0, 0)
        Me.pnlDireita.Name = "pnlDireita"
        Me.pnlDireita.Size = New System.Drawing.Size(722, 682)
        Me.pnlDireita.TabIndex = 0
        
        '
        'splitContainer2
        '
        Me.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.splitContainer2.Name = "splitContainer2"
        Me.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.splitContainer2.Panel1MinSize = 250
        Me.splitContainer2.Panel2MinSize = 200
        '
        'splitContainer2.Panel1
        '
        Me.splitContainer2.Panel1.Controls.Add(Me.pnlSuperior)
        '
        'splitContainer2.Panel2
        '
        Me.splitContainer2.Panel2.Controls.Add(Me.pnlInferior)
        Me.splitContainer2.Size = New System.Drawing.Size(722, 682)
        Me.splitContainer2.SplitterDistance = 341
        Me.splitContainer2.SplitterWidth = 8
        Me.splitContainer2.TabIndex = 0
        
        '
        'pnlSuperior
        '
        Me.pnlSuperior.Controls.Add(Me.splitContainer3)
        Me.pnlSuperior.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlSuperior.Location = New System.Drawing.Point(0, 0)
        Me.pnlSuperior.Name = "pnlSuperior"
        Me.pnlSuperior.Size = New System.Drawing.Size(722, 341)
        Me.pnlSuperior.TabIndex = 0
        
        '
        'splitContainer3
        '
        Me.splitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.splitContainer3.Name = "splitContainer3"
        Me.splitContainer3.Panel1MinSize = 250
        Me.splitContainer3.Panel2MinSize = 250
        '
        'splitContainer3.Panel1
        '
        Me.splitContainer3.Panel1.Controls.Add(Me.grpImagem)
        '
        'splitContainer3.Panel2
        '
        Me.splitContainer3.Panel2.Controls.Add(Me.grpEstoque)
        Me.splitContainer3.Size = New System.Drawing.Size(722, 341)
        Me.splitContainer3.SplitterDistance = 355
        Me.splitContainer3.SplitterWidth = 8
        Me.splitContainer3.TabIndex = 0
        
        '
        'grpImagem
        '
        Me.grpImagem.Controls.Add(Me.pbProduto)
        Me.grpImagem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpImagem.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpImagem.Location = New System.Drawing.Point(0, 0)
        Me.grpImagem.Name = "grpImagem"
        Me.grpImagem.Padding = New System.Windows.Forms.Padding(8)
        Me.grpImagem.Size = New System.Drawing.Size(355, 341)
        Me.grpImagem.TabIndex = 0
        Me.grpImagem.TabStop = False
        Me.grpImagem.Text = "🖼️ Imagem do Produto"
        
        '
        'pbProduto
        '
        Me.pbProduto.BackColor = System.Drawing.Color.White
        Me.pbProduto.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pbProduto.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pbProduto.Location = New System.Drawing.Point(8, 26)
        Me.pbProduto.Name = "pbProduto"
        Me.pbProduto.Size = New System.Drawing.Size(339, 307)
        Me.pbProduto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pbProduto.TabIndex = 0
        Me.pbProduto.TabStop = False
        
        '
        'grpEstoque
        '
        Me.grpEstoque.Controls.Add(Me.dgvEstoque)
        Me.grpEstoque.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpEstoque.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpEstoque.Location = New System.Drawing.Point(0, 0)
        Me.grpEstoque.Name = "grpEstoque"
        Me.grpEstoque.Padding = New System.Windows.Forms.Padding(8)
        Me.grpEstoque.Size = New System.Drawing.Size(359, 341)
        Me.grpEstoque.TabIndex = 0
        Me.grpEstoque.TabStop = False
        Me.grpEstoque.Text = "📊 Estoque Atual"
        
        '
        'dgvEstoque
        '
        Me.dgvEstoque.AllowUserToAddRows = False
        Me.dgvEstoque.AllowUserToDeleteRows = False
        Me.dgvEstoque.BackgroundColor = System.Drawing.Color.White
        Me.dgvEstoque.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgvEstoque.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvEstoque.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEstoque.Location = New System.Drawing.Point(8, 26)
        Me.dgvEstoque.MultiSelect = False
        Me.dgvEstoque.Name = "dgvEstoque"
        Me.dgvEstoque.ReadOnly = True
        Me.dgvEstoque.RowHeadersVisible = False
        Me.dgvEstoque.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvEstoque.Size = New System.Drawing.Size(343, 307)
        Me.dgvEstoque.TabIndex = 0
        
        '
        'pnlInferior
        '
        Me.pnlInferior.Controls.Add(Me.splitContainer4)
        Me.pnlInferior.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlInferior.Location = New System.Drawing.Point(0, 0)
        Me.pnlInferior.Name = "pnlInferior"
        Me.pnlInferior.Size = New System.Drawing.Size(722, 333)
        Me.pnlInferior.TabIndex = 0
        
        '
        'splitContainer4
        '
        Me.splitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.splitContainer4.Name = "splitContainer4"
        Me.splitContainer4.Panel1MinSize = 250
        Me.splitContainer4.Panel2MinSize = 250
        '
        'splitContainer4.Panel1
        '
        Me.splitContainer4.Panel1.Controls.Add(Me.grpCompras)
        '
        'splitContainer4.Panel2
        '
        Me.splitContainer4.Panel2.Controls.Add(Me.grpVendas)
        Me.splitContainer4.Size = New System.Drawing.Size(722, 333)
        Me.splitContainer4.SplitterDistance = 355
        Me.splitContainer4.SplitterWidth = 8
        Me.splitContainer4.TabIndex = 0
        
        '
        'grpCompras
        '
        Me.grpCompras.Controls.Add(Me.dgvCompras)
        Me.grpCompras.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpCompras.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpCompras.Location = New System.Drawing.Point(0, 0)
        Me.grpCompras.Name = "grpCompras"
        Me.grpCompras.Padding = New System.Windows.Forms.Padding(8)
        Me.grpCompras.Size = New System.Drawing.Size(355, 333)
        Me.grpCompras.TabIndex = 0
        Me.grpCompras.TabStop = False
        Me.grpCompras.Text = "📈 Compras"
        
        '
        'dgvCompras
        '
        Me.dgvCompras.AllowUserToAddRows = False
        Me.dgvCompras.AllowUserToDeleteRows = False
        Me.dgvCompras.BackgroundColor = System.Drawing.Color.White
        Me.dgvCompras.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgvCompras.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCompras.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCompras.Location = New System.Drawing.Point(8, 26)
        Me.dgvCompras.MultiSelect = False
        Me.dgvCompras.Name = "dgvCompras"
        Me.dgvCompras.ReadOnly = True
        Me.dgvCompras.RowHeadersVisible = False
        Me.dgvCompras.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvCompras.Size = New System.Drawing.Size(339, 299)
        Me.dgvCompras.TabIndex = 0
        
        '
        'grpVendas
        '
        Me.grpVendas.Controls.Add(Me.dgvVendas)
        Me.grpVendas.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpVendas.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpVendas.Location = New System.Drawing.Point(0, 0)
        Me.grpVendas.Name = "grpVendas"
        Me.grpVendas.Padding = New System.Windows.Forms.Padding(8)
        Me.grpVendas.Size = New System.Drawing.Size(359, 333)
        Me.grpVendas.TabIndex = 0
        Me.grpVendas.TabStop = False
        Me.grpVendas.Text = "📉 Vendas"
        
        '
        'dgvVendas
        '
        Me.dgvVendas.AllowUserToAddRows = False
        Me.dgvVendas.AllowUserToDeleteRows = False
        Me.dgvVendas.BackgroundColor = System.Drawing.Color.White
        Me.dgvVendas.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgvVendas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvVendas.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvVendas.Location = New System.Drawing.Point(8, 26)
        Me.dgvVendas.MultiSelect = False
        Me.dgvVendas.Name = "dgvVendas"
        Me.dgvVendas.ReadOnly = True
        Me.dgvVendas.RowHeadersVisible = False
        Me.dgvVendas.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvVendas.Size = New System.Drawing.Size(343, 299)
        Me.dgvVendas.TabIndex = 0
        
        '
        'UcReposicaoEstoque
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.pnlPrincipal)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Name = "UcReposicaoEstoque"
        Me.Size = New System.Drawing.Size(1150, 702)
        
        ' Finalizar configuração
        Me.pnlPrincipal.ResumeLayout(False)
        CType(Me.splitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitContainer1.Panel1.ResumeLayout(False)
        Me.splitContainer1.Panel2.ResumeLayout(False)
        Me.splitContainer1.ResumeLayout(False)
        Me.pnlEsquerda.ResumeLayout(False)
        Me.grpProdutos.ResumeLayout(False)
        CType(Me.dgvProdutos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFiltros.ResumeLayout(False)
        Me.pnlFiltros.PerformLayout()
        Me.pnlDireita.ResumeLayout(False)
        CType(Me.splitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitContainer2.Panel1.ResumeLayout(False)
        Me.splitContainer2.Panel2.ResumeLayout(False)
        Me.splitContainer2.ResumeLayout(False)
        Me.pnlSuperior.ResumeLayout(False)
        CType(Me.splitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitContainer3.Panel1.ResumeLayout(False)
        Me.splitContainer3.Panel2.ResumeLayout(False)
        Me.splitContainer3.ResumeLayout(False)
        Me.grpImagem.ResumeLayout(False)
        CType(Me.pbProduto, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpEstoque.ResumeLayout(False)
        CType(Me.dgvEstoque, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlInferior.ResumeLayout(False)
        CType(Me.splitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitContainer4.Panel1.ResumeLayout(False)
        Me.splitContainer4.Panel2.ResumeLayout(False)
        Me.splitContainer4.ResumeLayout(False)
        Me.grpCompras.ResumeLayout(False)
        CType(Me.dgvCompras, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpVendas.ResumeLayout(False)
        CType(Me.dgvVendas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlPrincipal As Panel
    Friend WithEvents splitContainer1 As SplitContainer
    Friend WithEvents pnlEsquerda As Panel
    Friend WithEvents grpProdutos As GroupBox
    Friend WithEvents dgvProdutos As DataGridView
    Friend WithEvents pnlFiltros As Panel
    Friend WithEvents txtFiltro As TextBox
    Friend WithEvents lblFiltro As Label
    Friend WithEvents btnAtualizar As Button
    Friend WithEvents pnlDireita As Panel
    Friend WithEvents splitContainer2 As SplitContainer
    Friend WithEvents pnlSuperior As Panel
    Friend WithEvents splitContainer3 As SplitContainer
    Friend WithEvents grpImagem As GroupBox
    Friend WithEvents pbProduto As PictureBox
    Friend WithEvents grpEstoque As GroupBox
    Friend WithEvents dgvEstoque As DataGridView
    Friend WithEvents pnlInferior As Panel
    Friend WithEvents splitContainer4 As SplitContainer
    Friend WithEvents grpCompras As GroupBox
    Friend WithEvents dgvCompras As DataGridView
    Friend WithEvents grpVendas As GroupBox
    Friend WithEvents dgvVendas As DataGridView
End Class