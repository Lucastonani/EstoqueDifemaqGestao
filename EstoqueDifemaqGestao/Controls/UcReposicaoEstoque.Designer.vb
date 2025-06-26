<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class UcReposicaoEstoque
    Inherits System.Windows.Forms.UserControl

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTA: O procedimento a seguir é exigido pelo Windows Form Designer
    'Pode ser modificado usando o Windows Form Designer.  
    'Não o modifique usando o editor de códigos.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend2 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series2 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.pnlPrincipal = New System.Windows.Forms.Panel()
        Me.splitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.pnlEsquerda = New System.Windows.Forms.Panel()
        Me.grpProdutos = New System.Windows.Forms.GroupBox()
        Me.dgvProdutos = New System.Windows.Forms.DataGridView()
        Me.pnlFiltros = New System.Windows.Forms.Panel()
        Me.txtFiltro = New System.Windows.Forms.TextBox()
        Me.lblFiltro = New System.Windows.Forms.Label()
        Me.btnAtualizar = New System.Windows.Forms.Button()
        Me.btnAtualizarPowerQuery = New System.Windows.Forms.Button()
        Me.lblUltimaAtualizacao = New System.Windows.Forms.Label()
        Me.grpInclusaoManual = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblCodigoManual = New System.Windows.Forms.Label()
        Me.txtCodigoManual = New System.Windows.Forms.TextBox()
        Me.lblDescricaoManual = New System.Windows.Forms.Label()
        Me.txtDescricaoManual = New System.Windows.Forms.TextBox()
        Me.lblFabricanteManual = New System.Windows.Forms.Label()
        Me.txtFabricanteManual = New System.Windows.Forms.TextBox()
        Me.lblQtdInicialManual = New System.Windows.Forms.Label()
        Me.txtQtdInicialManual = New System.Windows.Forms.TextBox()
        Me.lblDataManual = New System.Windows.Forms.Label()
        Me.txtDataManual = New System.Windows.Forms.TextBox()
        Me.btnIncluirProdutoManual = New System.Windows.Forms.Button()
        Me.pnlDireita = New System.Windows.Forms.Panel()
        Me.splitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.pnlSuperior = New System.Windows.Forms.Panel()
        Me.splitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.pnlImagemEAplicacao = New System.Windows.Forms.Panel()
        Me.grpImagem = New System.Windows.Forms.GroupBox()
        Me.pbProduto = New System.Windows.Forms.PictureBox()
        Me.grpAplicacao = New System.Windows.Forms.GroupBox()
        Me.txtAplicacaoProduto = New System.Windows.Forms.TextBox()
        Me.pnlDireitaSuperior = New System.Windows.Forms.Panel()
        Me.grpEstoque = New System.Windows.Forms.GroupBox()
        Me.dgvEstoque = New System.Windows.Forms.DataGridView()
        Me.grpPedidoProduto = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblEstoqueMinimo = New System.Windows.Forms.Label()
        Me.txtEstoqueMinimo = New System.Windows.Forms.TextBox()
        Me.lblEstoqueMaximo = New System.Windows.Forms.Label()
        Me.txtEstoqueMaximo = New System.Windows.Forms.TextBox()
        Me.lblQtdPedir = New System.Windows.Forms.Label()
        Me.txtQtdPedir = New System.Windows.Forms.TextBox()
        Me.lblLoja = New System.Windows.Forms.Label()
        Me.comboLoja = New System.Windows.Forms.ComboBox()
        Me.lblCliente = New System.Windows.Forms.Label()
        Me.txtCliente = New System.Windows.Forms.TextBox()
        Me.lblTelefone = New System.Windows.Forms.Label()
        Me.txtTelefone = New System.Windows.Forms.MaskedTextBox()
        Me.btnProximoProduto = New System.Windows.Forms.Button()
        Me.pnlInferior = New System.Windows.Forms.Panel()
        Me.splitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.grpCompras = New System.Windows.Forms.GroupBox()
        Me.chartComprasMensais = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.grpVendas = New System.Windows.Forms.GroupBox()
        Me.chartVendasMensais = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.pnlBotoesSessao = New System.Windows.Forms.Panel()
        Me.btnDescartarPedido = New System.Windows.Forms.Button()
        Me.btnFinalizarPedido = New System.Windows.Forms.Button()
        Me.btnNovoPedido = New System.Windows.Forms.Button()
        Me.pnlPrincipal.SuspendLayout()
        CType(Me.splitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitContainer1.Panel1.SuspendLayout()
        Me.splitContainer1.Panel2.SuspendLayout()
        Me.splitContainer1.SuspendLayout()
        Me.pnlEsquerda.SuspendLayout()
        Me.grpProdutos.SuspendLayout()
        CType(Me.dgvProdutos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFiltros.SuspendLayout()
        Me.grpInclusaoManual.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
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
        Me.pnlImagemEAplicacao.SuspendLayout()
        Me.grpImagem.SuspendLayout()
        CType(Me.pbProduto, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAplicacao.SuspendLayout()
        Me.pnlDireitaSuperior.SuspendLayout()
        Me.grpEstoque.SuspendLayout()
        CType(Me.dgvEstoque, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpPedidoProduto.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.pnlInferior.SuspendLayout()
        CType(Me.splitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitContainer4.Panel1.SuspendLayout()
        Me.splitContainer4.Panel2.SuspendLayout()
        Me.splitContainer4.SuspendLayout()
        Me.grpCompras.SuspendLayout()
        CType(Me.chartComprasMensais, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpVendas.SuspendLayout()
        CType(Me.chartVendasMensais, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlBotoesSessao.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlPrincipal
        '
        Me.pnlPrincipal.Controls.Add(Me.splitContainer1)
        Me.pnlPrincipal.Controls.Add(Me.pnlBotoesSessao)
        Me.pnlPrincipal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlPrincipal.Location = New System.Drawing.Point(0, 0)
        Me.pnlPrincipal.Name = "pnlPrincipal"
        Me.pnlPrincipal.Padding = New System.Windows.Forms.Padding(10)
        Me.pnlPrincipal.Size = New System.Drawing.Size(1200, 800)
        Me.pnlPrincipal.TabIndex = 0
        '
        'splitContainer1
        '
        Me.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.splitContainer1.Location = New System.Drawing.Point(10, 10)
        Me.splitContainer1.Name = "splitContainer1"
        '
        'splitContainer1.Panel1
        '
        Me.splitContainer1.Panel1.Controls.Add(Me.pnlEsquerda)
        Me.splitContainer1.Panel1MinSize = 400
        '
        'splitContainer1.Panel2
        '
        Me.splitContainer1.Panel2.Controls.Add(Me.pnlDireita)
        Me.splitContainer1.Panel2MinSize = 600
        Me.splitContainer1.Size = New System.Drawing.Size(1180, 720)
        Me.splitContainer1.SplitterDistance = 450
        Me.splitContainer1.SplitterWidth = 8
        Me.splitContainer1.TabIndex = 0
        '
        'pnlEsquerda
        '
        Me.pnlEsquerda.Controls.Add(Me.grpProdutos)
        Me.pnlEsquerda.Controls.Add(Me.grpInclusaoManual)
        Me.pnlEsquerda.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlEsquerda.Location = New System.Drawing.Point(0, 0)
        Me.pnlEsquerda.Name = "pnlEsquerda"
        Me.pnlEsquerda.Size = New System.Drawing.Size(450, 720)
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
        Me.grpProdutos.Size = New System.Drawing.Size(450, 520)
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
        Me.dgvProdutos.Location = New System.Drawing.Point(8, 66)
        Me.dgvProdutos.MultiSelect = False
        Me.dgvProdutos.Name = "dgvProdutos"
        Me.dgvProdutos.ReadOnly = True
        Me.dgvProdutos.RowHeadersVisible = False
        Me.dgvProdutos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvProdutos.Size = New System.Drawing.Size(434, 446)
        Me.dgvProdutos.TabIndex = 1
        '
        'pnlFiltros
        '
        Me.pnlFiltros.Controls.Add(Me.txtFiltro)
        Me.pnlFiltros.Controls.Add(Me.lblFiltro)
        Me.pnlFiltros.Controls.Add(Me.btnAtualizar)
        Me.pnlFiltros.Controls.Add(Me.btnAtualizarPowerQuery)
        Me.pnlFiltros.Controls.Add(Me.lblUltimaAtualizacao)
        Me.pnlFiltros.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFiltros.Location = New System.Drawing.Point(8, 26)
        Me.pnlFiltros.Name = "pnlFiltros"
        Me.pnlFiltros.Size = New System.Drawing.Size(434, 70)
        Me.pnlFiltros.TabIndex = 0
        '
        'txtFiltro
        '
        Me.txtFiltro.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFiltro.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtFiltro.Location = New System.Drawing.Point(50, 9)
        Me.txtFiltro.Name = "txtFiltro"
        Me.txtFiltro.Size = New System.Drawing.Size(194, 23)
        Me.txtFiltro.TabIndex = 1
        '
        'lblFiltro
        '
        Me.lblFiltro.AutoSize = True
        Me.lblFiltro.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblFiltro.Location = New System.Drawing.Point(3, 12)
        Me.lblFiltro.Name = "lblFiltro"
        Me.lblFiltro.Size = New System.Drawing.Size(40, 15)
        Me.lblFiltro.TabIndex = 0
        Me.lblFiltro.Text = "Filtrar:"
        '
        'btnAtualizar
        '
        Me.btnAtualizar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAtualizar.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(123, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnAtualizar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAtualizar.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnAtualizar.ForeColor = System.Drawing.Color.White
        Me.btnAtualizar.Location = New System.Drawing.Point(250, 7)
        Me.btnAtualizar.Name = "btnAtualizar"
        Me.btnAtualizar.Size = New System.Drawing.Size(70, 27)
        Me.btnAtualizar.TabIndex = 2
        Me.btnAtualizar.Text = "Filtrar"
        Me.btnAtualizar.UseVisualStyleBackColor = False
        '
        'btnAtualizarPowerQuery
        '
        Me.btnAtualizarPowerQuery.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAtualizarPowerQuery.BackColor = System.Drawing.Color.FromArgb(CType(CType(40, Byte), Integer), CType(CType(167, Byte), Integer), CType(CType(69, Byte), Integer))
        Me.btnAtualizarPowerQuery.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAtualizarPowerQuery.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnAtualizarPowerQuery.ForeColor = System.Drawing.Color.White
        Me.btnAtualizarPowerQuery.Location = New System.Drawing.Point(326, 7)
        Me.btnAtualizarPowerQuery.Name = "btnAtualizarPowerQuery"
        Me.btnAtualizarPowerQuery.Size = New System.Drawing.Size(105, 27)
        Me.btnAtualizarPowerQuery.TabIndex = 3
        Me.btnAtualizarPowerQuery.Text = "📊 Atualizar"
        Me.btnAtualizarPowerQuery.UseVisualStyleBackColor = False
        '
        'lblUltimaAtualizacao
        '
        Me.lblUltimaAtualizacao.AutoSize = True
        Me.lblUltimaAtualizacao.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Italic)
        Me.lblUltimaAtualizacao.ForeColor = System.Drawing.Color.Gray
        Me.lblUltimaAtualizacao.Location = New System.Drawing.Point(3, 45)
        Me.lblUltimaAtualizacao.Name = "lblUltimaAtualizacao"
        Me.lblUltimaAtualizacao.Size = New System.Drawing.Size(120, 13)
        Me.lblUltimaAtualizacao.TabIndex = 4
        Me.lblUltimaAtualizacao.Text = "Última atualização: Nunca"
        '
        'grpInclusaoManual
        '
        Me.grpInclusaoManual.Controls.Add(Me.TableLayoutPanel1)
        Me.grpInclusaoManual.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.grpInclusaoManual.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpInclusaoManual.Location = New System.Drawing.Point(0, 520)
        Me.grpInclusaoManual.Name = "grpInclusaoManual"
        Me.grpInclusaoManual.Size = New System.Drawing.Size(450, 200)
        Me.grpInclusaoManual.TabIndex = 1
        Me.grpInclusaoManual.TabStop = False
        Me.grpInclusaoManual.Text = "➕ Inclusão Manual"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.lblCodigoManual, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtCodigoManual, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.lblDescricaoManual, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtDescricaoManual, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.lblFabricanteManual, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtFabricanteManual, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.lblQtdInicialManual, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtQtdInicialManual, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.lblDataManual, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.txtDataManual, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.btnIncluirProdutoManual, 1, 5)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 21)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.Padding = New System.Windows.Forms.Padding(5)
        Me.TableLayoutPanel1.RowCount = 6
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(444, 176)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'lblCodigoManual
        '
        Me.lblCodigoManual.AutoSize = True
        Me.lblCodigoManual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCodigoManual.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblCodigoManual.Location = New System.Drawing.Point(8, 5)
        Me.lblCodigoManual.Name = "lblCodigoManual"
        Me.lblCodigoManual.Size = New System.Drawing.Size(124, 28)
        Me.lblCodigoManual.TabIndex = 0
        Me.lblCodigoManual.Text = "Código:"
        Me.lblCodigoManual.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCodigoManual
        '
        Me.txtCodigoManual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtCodigoManual.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtCodigoManual.Location = New System.Drawing.Point(138, 8)
        Me.txtCodigoManual.MaxLength = 5
        Me.txtCodigoManual.Name = "txtCodigoManual"
        Me.txtCodigoManual.Size = New System.Drawing.Size(298, 23)
        Me.txtCodigoManual.TabIndex = 1
        '
        'lblDescricaoManual
        '
        Me.lblDescricaoManual.AutoSize = True
        Me.lblDescricaoManual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblDescricaoManual.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblDescricaoManual.Location = New System.Drawing.Point(8, 33)
        Me.lblDescricaoManual.Name = "lblDescricaoManual"
        Me.lblDescricaoManual.Size = New System.Drawing.Size(124, 28)
        Me.lblDescricaoManual.TabIndex = 2
        Me.lblDescricaoManual.Text = "Descrição:"
        Me.lblDescricaoManual.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDescricaoManual
        '
        Me.txtDescricaoManual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtDescricaoManual.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtDescricaoManual.Location = New System.Drawing.Point(138, 36)
        Me.txtDescricaoManual.Name = "txtDescricaoManual"
        Me.txtDescricaoManual.Size = New System.Drawing.Size(298, 23)
        Me.txtDescricaoManual.TabIndex = 3
        '
        'lblFabricanteManual
        '
        Me.lblFabricanteManual.AutoSize = True
        Me.lblFabricanteManual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblFabricanteManual.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblFabricanteManual.Location = New System.Drawing.Point(8, 61)
        Me.lblFabricanteManual.Name = "lblFabricanteManual"
        Me.lblFabricanteManual.Size = New System.Drawing.Size(124, 28)
        Me.lblFabricanteManual.TabIndex = 4
        Me.lblFabricanteManual.Text = "Fabricante:"
        Me.lblFabricanteManual.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtFabricanteManual
        '
        Me.txtFabricanteManual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtFabricanteManual.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtFabricanteManual.Location = New System.Drawing.Point(138, 64)
        Me.txtFabricanteManual.Name = "txtFabricanteManual"
        Me.txtFabricanteManual.Size = New System.Drawing.Size(298, 23)
        Me.txtFabricanteManual.TabIndex = 5
        '
        'lblQtdInicialManual
        '
        Me.lblQtdInicialManual.AutoSize = True
        Me.lblQtdInicialManual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblQtdInicialManual.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblQtdInicialManual.Location = New System.Drawing.Point(8, 89)
        Me.lblQtdInicialManual.Name = "lblQtdInicialManual"
        Me.lblQtdInicialManual.Size = New System.Drawing.Size(124, 28)
        Me.lblQtdInicialManual.TabIndex = 6
        Me.lblQtdInicialManual.Text = "Quantidade Inicial:"
        Me.lblQtdInicialManual.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtQtdInicialManual
        '
        Me.txtQtdInicialManual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtQtdInicialManual.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtQtdInicialManual.Location = New System.Drawing.Point(138, 92)
        Me.txtQtdInicialManual.Name = "txtQtdInicialManual"
        Me.txtQtdInicialManual.Size = New System.Drawing.Size(298, 23)
        Me.txtQtdInicialManual.TabIndex = 7
        '
        'lblDataManual
        '
        Me.lblDataManual.AutoSize = True
        Me.lblDataManual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblDataManual.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblDataManual.Location = New System.Drawing.Point(8, 117)
        Me.lblDataManual.Name = "lblDataManual"
        Me.lblDataManual.Size = New System.Drawing.Size(124, 28)
        Me.lblDataManual.TabIndex = 8
        Me.lblDataManual.Text = "Data:"
        Me.lblDataManual.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDataManual
        '
        Me.txtDataManual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtDataManual.Enabled = False
        Me.txtDataManual.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtDataManual.Location = New System.Drawing.Point(138, 120)
        Me.txtDataManual.Name = "txtDataManual"
        Me.txtDataManual.ReadOnly = True
        Me.txtDataManual.Size = New System.Drawing.Size(298, 23)
        Me.txtDataManual.TabIndex = 9
        '
        'btnIncluirProdutoManual
        '
        Me.btnIncluirProdutoManual.BackColor = System.Drawing.Color.FromArgb(CType(CType(40, Byte), Integer), CType(CType(167, Byte), Integer), CType(CType(69, Byte), Integer))
        Me.btnIncluirProdutoManual.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnIncluirProdutoManual.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnIncluirProdutoManual.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnIncluirProdutoManual.ForeColor = System.Drawing.Color.White
        Me.btnIncluirProdutoManual.Location = New System.Drawing.Point(336, 148)
        Me.btnIncluirProdutoManual.Name = "btnIncluirProdutoManual"
        Me.btnIncluirProdutoManual.Size = New System.Drawing.Size(100, 20)
        Me.btnIncluirProdutoManual.TabIndex = 10
        Me.btnIncluirProdutoManual.Text = "Incluir"
        Me.btnIncluirProdutoManual.UseVisualStyleBackColor = False
        '
        'pnlDireita
        '
        Me.pnlDireita.Controls.Add(Me.splitContainer2)
        Me.pnlDireita.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlDireita.Location = New System.Drawing.Point(0, 0)
        Me.pnlDireita.Name = "pnlDireita"
        Me.pnlDireita.Size = New System.Drawing.Size(722, 720)
        Me.pnlDireita.TabIndex = 0
        '
        'splitContainer2
        '
        Me.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.splitContainer2.Name = "splitContainer2"
        Me.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitContainer2.Panel1
        '
        Me.splitContainer2.Panel1.Controls.Add(Me.pnlSuperior)
        Me.splitContainer2.Panel1MinSize = 300
        '
        'splitContainer2.Panel2
        '
        Me.splitContainer2.Panel2.Controls.Add(Me.pnlInferior)
        Me.splitContainer2.Panel2MinSize = 300
        Me.splitContainer2.Size = New System.Drawing.Size(722, 720)
        Me.splitContainer2.SplitterDistance = 360
        Me.splitContainer2.SplitterWidth = 8
        Me.splitContainer2.TabIndex = 0
        '
        'pnlSuperior
        '
        Me.pnlSuperior.Controls.Add(Me.splitContainer3)
        Me.pnlSuperior.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlSuperior.Location = New System.Drawing.Point(0, 0)
        Me.pnlSuperior.Name = "pnlSuperior"
        Me.pnlSuperior.Size = New System.Drawing.Size(722, 360)
        Me.pnlSuperior.TabIndex = 0
        '
        'splitContainer3
        '
        Me.splitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.splitContainer3.Name = "splitContainer3"
        '
        'splitContainer3.Panel1
        '
        Me.splitContainer3.Panel1.Controls.Add(Me.pnlImagemEAplicacao)
        Me.splitContainer3.Panel1MinSize = 250
        '
        'splitContainer3.Panel2
        '
        Me.splitContainer3.Panel2.Controls.Add(Me.pnlDireitaSuperior)
        Me.splitContainer3.Panel2MinSize = 350
        Me.splitContainer3.Size = New System.Drawing.Size(722, 360)
        Me.splitContainer3.SplitterDistance = 300
        Me.splitContainer3.SplitterWidth = 8
        Me.splitContainer3.TabIndex = 0
        '
        'pnlImagemEAplicacao
        '
        Me.pnlImagemEAplicacao.Controls.Add(Me.grpImagem)
        Me.pnlImagemEAplicacao.Controls.Add(Me.grpAplicacao)
        Me.pnlImagemEAplicacao.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlImagemEAplicacao.Location = New System.Drawing.Point(0, 0)
        Me.pnlImagemEAplicacao.Name = "pnlImagemEAplicacao"
        Me.pnlImagemEAplicacao.Size = New System.Drawing.Size(300, 360)
        Me.pnlImagemEAplicacao.TabIndex = 0
        '
        'grpImagem
        '
        Me.grpImagem.Controls.Add(Me.pbProduto)
        Me.grpImagem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpImagem.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpImagem.Location = New System.Drawing.Point(0, 0)
        Me.grpImagem.Name = "grpImagem"
        Me.grpImagem.Padding = New System.Windows.Forms.Padding(8)
        Me.grpImagem.Size = New System.Drawing.Size(300, 290)
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
        Me.pbProduto.Size = New System.Drawing.Size(284, 256)
        Me.pbProduto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pbProduto.TabIndex = 0
        Me.pbProduto.TabStop = False
        '
        'grpAplicacao
        '
        Me.grpAplicacao.Controls.Add(Me.txtAplicacaoProduto)
        Me.grpAplicacao.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.grpAplicacao.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpAplicacao.Location = New System.Drawing.Point(0, 290)
        Me.grpAplicacao.Name = "grpAplicacao"
        Me.grpAplicacao.Padding = New System.Windows.Forms.Padding(8)
        Me.grpAplicacao.Size = New System.Drawing.Size(300, 70)
        Me.grpAplicacao.TabIndex = 1
        Me.grpAplicacao.TabStop = False
        Me.grpAplicacao.Text = "📋 Aplicação do Produto"
        '
        'txtAplicacaoProduto
        '
        Me.txtAplicacaoProduto.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtAplicacaoProduto.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtAplicacaoProduto.Location = New System.Drawing.Point(8, 26)
        Me.txtAplicacaoProduto.Multiline = True
        Me.txtAplicacaoProduto.Name = "txtAplicacaoProduto"
        Me.txtAplicacaoProduto.ReadOnly = True
        Me.txtAplicacaoProduto.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtAplicacaoProduto.Size = New System.Drawing.Size(284, 36)
        Me.txtAplicacaoProduto.TabIndex = 0
        '
        'pnlDireitaSuperior
        '
        Me.pnlDireitaSuperior.Controls.Add(Me.grpEstoque)
        Me.pnlDireitaSuperior.Controls.Add(Me.grpPedidoProduto)
        Me.pnlDireitaSuperior.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlDireitaSuperior.Location = New System.Drawing.Point(0, 0)
        Me.pnlDireitaSuperior.Name = "pnlDireitaSuperior"
        Me.pnlDireitaSuperior.Size = New System.Drawing.Size(414, 360)
        Me.pnlDireitaSuperior.TabIndex = 0
        '
        'grpEstoque
        '
        Me.grpEstoque.Controls.Add(Me.dgvEstoque)
        Me.grpEstoque.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpEstoque.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpEstoque.Location = New System.Drawing.Point(0, 0)
        Me.grpEstoque.Name = "grpEstoque"
        Me.grpEstoque.Padding = New System.Windows.Forms.Padding(8)
        Me.grpEstoque.Size = New System.Drawing.Size(414, 160)
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
        Me.dgvEstoque.Size = New System.Drawing.Size(398, 126)
        Me.dgvEstoque.TabIndex = 0
        '
        'grpPedidoProduto
        '
        Me.grpPedidoProduto.Controls.Add(Me.TableLayoutPanel2)
        Me.grpPedidoProduto.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.grpPedidoProduto.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpPedidoProduto.Location = New System.Drawing.Point(0, 160)
        Me.grpPedidoProduto.Name = "grpPedidoProduto"
        Me.grpPedidoProduto.Size = New System.Drawing.Size(414, 200)
        Me.grpPedidoProduto.TabIndex = 1
        Me.grpPedidoProduto.TabStop = False
        Me.grpPedidoProduto.Text = "📝 Pedido do Produto Selecionado"
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 4
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.lblEstoqueMinimo, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.txtEstoqueMinimo, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.lblEstoqueMaximo, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.txtEstoqueMaximo, 3, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.lblQtdPedir, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.txtQtdPedir, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblLoja, 2, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.comboLoja, 3, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblCliente, 0, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.txtCliente, 1, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.lblTelefone, 2, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.txtTelefone, 3, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.btnProximoProduto, 2, 3)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(3, 21)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.Padding = New System.Windows.Forms.Padding(5)
        Me.TableLayoutPanel2.RowCount = 4
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(408, 176)
        Me.TableLayoutPanel2.TabIndex = 0
        '
        'lblEstoqueMinimo
        '
        Me.lblEstoqueMinimo.AutoSize = True
        Me.lblEstoqueMinimo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblEstoqueMinimo.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblEstoqueMinimo.Location = New System.Drawing.Point(8, 5)
        Me.lblEstoqueMinimo.Name = "lblEstoqueMinimo"
        Me.lblEstoqueMinimo.Size = New System.Drawing.Size(93, 35)
        Me.lblEstoqueMinimo.TabIndex = 0
        Me.lblEstoqueMinimo.Text = "Estoque Mín:"
        Me.lblEstoqueMinimo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEstoqueMinimo
        '
        Me.txtEstoqueMinimo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtEstoqueMinimo.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtEstoqueMinimo.Location = New System.Drawing.Point(107, 8)
        Me.txtEstoqueMinimo.Name = "txtEstoqueMinimo"
        Me.txtEstoqueMinimo.ReadOnly = True
        Me.txtEstoqueMinimo.Size = New System.Drawing.Size(93, 23)
        Me.txtEstoqueMinimo.TabIndex = 1
        Me.txtEstoqueMinimo.Text = "0"
        '
        'lblEstoqueMaximo
        '
        Me.lblEstoqueMaximo.AutoSize = True
        Me.lblEstoqueMaximo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblEstoqueMaximo.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblEstoqueMaximo.Location = New System.Drawing.Point(206, 5)
        Me.lblEstoqueMaximo.Name = "lblEstoqueMaximo"
        Me.lblEstoqueMaximo.Size = New System.Drawing.Size(93, 35)
        Me.lblEstoqueMaximo.TabIndex = 2
        Me.lblEstoqueMaximo.Text = "Estoque Máx:"
        Me.lblEstoqueMaximo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEstoqueMaximo
        '
        Me.txtEstoqueMaximo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtEstoqueMaximo.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtEstoqueMaximo.Location = New System.Drawing.Point(305, 8)
        Me.txtEstoqueMaximo.Name = "txtEstoqueMaximo"
        Me.txtEstoqueMaximo.ReadOnly = True
        Me.txtEstoqueMaximo.Size = New System.Drawing.Size(95, 23)
        Me.txtEstoqueMaximo.TabIndex = 3
        Me.txtEstoqueMaximo.Text = "0"
        '
        'lblQtdPedir
        '
        Me.lblQtdPedir.AutoSize = True
        Me.lblQtdPedir.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblQtdPedir.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblQtdPedir.Location = New System.Drawing.Point(8, 40)
        Me.lblQtdPedir.Name = "lblQtdPedir"
        Me.lblQtdPedir.Size = New System.Drawing.Size(93, 35)
        Me.lblQtdPedir.TabIndex = 4
        Me.lblQtdPedir.Text = "Qtd a Pedir:"
        Me.lblQtdPedir.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtQtdPedir
        '
        Me.txtQtdPedir.BackColor = System.Drawing.Color.LightYellow
        Me.txtQtdPedir.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtQtdPedir.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.txtQtdPedir.Location = New System.Drawing.Point(107, 43)
        Me.txtQtdPedir.Name = "txtQtdPedir"
        Me.txtQtdPedir.Size = New System.Drawing.Size(93, 23)
        Me.txtQtdPedir.TabIndex = 5
        Me.txtQtdPedir.Text = "0"
        '
        'lblLoja
        '
        Me.lblLoja.AutoSize = True
        Me.lblLoja.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblLoja.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblLoja.Location = New System.Drawing.Point(206, 40)
        Me.lblLoja.Name = "lblLoja"
        Me.lblLoja.Size = New System.Drawing.Size(93, 35)
        Me.lblLoja.TabIndex = 6
        Me.lblLoja.Text = "Loja Destino:"
        Me.lblLoja.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'comboLoja
        '
        Me.comboLoja.Dock = System.Windows.Forms.DockStyle.Fill
        Me.comboLoja.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.comboLoja.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.comboLoja.FormattingEnabled = True
        Me.comboLoja.Location = New System.Drawing.Point(305, 43)
        Me.comboLoja.Name = "comboLoja"
        Me.comboLoja.Size = New System.Drawing.Size(95, 23)
        Me.comboLoja.TabIndex = 7
        '
        'lblCliente
        '
        Me.lblCliente.AutoSize = True
        Me.lblCliente.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCliente.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblCliente.Location = New System.Drawing.Point(8, 75)
        Me.lblCliente.Name = "lblCliente"
        Me.lblCliente.Size = New System.Drawing.Size(93, 35)
        Me.lblCliente.TabIndex = 8
        Me.lblCliente.Text = "Cliente:"
        Me.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCliente
        '
        Me.txtCliente.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtCliente.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtCliente.Location = New System.Drawing.Point(107, 78)
        Me.txtCliente.Name = "txtCliente"
        Me.txtCliente.Size = New System.Drawing.Size(93, 23)
        Me.txtCliente.TabIndex = 9
        '
        'lblTelefone
        '
        Me.lblTelefone.AutoSize = True
        Me.lblTelefone.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTelefone.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblTelefone.Location = New System.Drawing.Point(206, 75)
        Me.lblTelefone.Name = "lblTelefone"
        Me.lblTelefone.Size = New System.Drawing.Size(93, 35)
        Me.lblTelefone.TabIndex = 10
        Me.lblTelefone.Text = "Telefone:"
        Me.lblTelefone.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtTelefone
        '
        Me.txtTelefone.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtTelefone.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtTelefone.Location = New System.Drawing.Point(305, 78)
        Me.txtTelefone.Mask = "(00)00000-0000"
        Me.txtTelefone.Name = "txtTelefone"
        Me.txtTelefone.Size = New System.Drawing.Size(95, 23)
        Me.txtTelefone.TabIndex = 11
        '
        'btnProximoProduto
        '
        Me.btnProximoProduto.BackColor = System.Drawing.Color.FromArgb(CType(CType(52, Byte), Integer), CType(CType(152, Byte), Integer), CType(CType(219, Byte), Integer))
        Me.TableLayoutPanel2.SetColumnSpan(Me.btnProximoProduto, 2)
        Me.btnProximoProduto.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnProximoProduto.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnProximoProduto.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnProximoProduto.ForeColor = System.Drawing.Color.White
        Me.btnProximoProduto.Location = New System.Drawing.Point(250, 113)
        Me.btnProximoProduto.Name = "btnProximoProduto"
        Me.btnProximoProduto.Size = New System.Drawing.Size(150, 55)
        Me.btnProximoProduto.TabIndex = 12
        Me.btnProximoProduto.Text = "Próximo Produto ➡️"
        Me.btnProximoProduto.UseVisualStyleBackColor = False
        '
        'pnlInferior
        '
        Me.pnlInferior.Controls.Add(Me.splitContainer4)
        Me.pnlInferior.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlInferior.Location = New System.Drawing.Point(0, 0)
        Me.pnlInferior.Name = "pnlInferior"
        Me.pnlInferior.Size = New System.Drawing.Size(722, 352)
        Me.pnlInferior.TabIndex = 0
        '
        'splitContainer4
        '
        Me.splitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.splitContainer4.Name = "splitContainer4"
        '
        'splitContainer4.Panel1
        '
        Me.splitContainer4.Panel1.Controls.Add(Me.grpCompras)
        Me.splitContainer4.Panel1MinSize = 300
        '
        'splitContainer4.Panel2
        '
        Me.splitContainer4.Panel2.Controls.Add(Me.grpVendas)
        Me.splitContainer4.Panel2MinSize = 300
        Me.splitContainer4.Size = New System.Drawing.Size(722, 352)
        Me.splitContainer4.SplitterDistance = 355
        Me.splitContainer4.SplitterWidth = 8
        Me.splitContainer4.TabIndex = 0
        '
        'grpCompras
        '
        Me.grpCompras.Controls.Add(Me.chartComprasMensais)
        Me.grpCompras.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpCompras.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpCompras.Location = New System.Drawing.Point(0, 0)
        Me.grpCompras.Name = "grpCompras"
        Me.grpCompras.Padding = New System.Windows.Forms.Padding(8)
        Me.grpCompras.Size = New System.Drawing.Size(355, 352)
        Me.grpCompras.TabIndex = 0
        Me.grpCompras.TabStop = False
        Me.grpCompras.Text = "📈 Histórico de Compras (24 meses)"
        '
        'chartComprasMensais
        '
        Legend1.Name = "Legend1"
        Me.chartComprasMensais.Legends.Add(Legend1)
        Me.chartComprasMensais.Location = New System.Drawing.Point(8, 26)
        Me.chartComprasMensais.Name = "chartComprasMensais"
        Series1.BorderWidth = 3
        Series1.ChartArea = "ChartArea1"
        Series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        Series1.Color = System.Drawing.Color.FromArgb(CType(CType(52, Byte), Integer), CType(CType(152, Byte), Integer), CType(CType(219, Byte), Integer))
        Series1.Legend = "Legend1"
        Series1.MarkerBorderColor = System.Drawing.Color.FromArgb(CType(CType(41, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(185, Byte), Integer))
        Series1.MarkerColor = System.Drawing.Color.FromArgb(CType(CType(52, Byte), Integer), CType(CType(152, Byte), Integer), CType(CType(219, Byte), Integer))
        Series1.MarkerSize = 8
        Series1.MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Circle
        Series1.Name = "Compras"
        Me.chartComprasMensais.Series.Add(Series1)
        Me.chartComprasMensais.Size = New System.Drawing.Size(339, 318)
        Me.chartComprasMensais.TabIndex = 0
        Me.chartComprasMensais.Text = "Chart1"
        '
        'grpVendas
        '
        Me.grpVendas.Controls.Add(Me.chartVendasMensais)
        Me.grpVendas.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpVendas.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpVendas.Location = New System.Drawing.Point(0, 0)
        Me.grpVendas.Name = "grpVendas"
        Me.grpVendas.Padding = New System.Windows.Forms.Padding(8)
        Me.grpVendas.Size = New System.Drawing.Size(359, 352)
        Me.grpVendas.TabIndex = 0
        Me.grpVendas.TabStop = False
        Me.grpVendas.Text = "📉 Histórico de Vendas (24 meses)"
        '
        'chartVendasMensais
        '
        Me.chartVendasMensais.BackColor = System.Drawing.Color.WhiteSmoke
        Me.chartVendasMensais.BorderlineColor = System.Drawing.Color.LightGray
        Me.chartVendasMensais.BorderlineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Solid
        ChartArea1.AxisX.IntervalAutoMode = System.Windows.Forms.DataVisualization.Charting.IntervalAutoMode.VariableCount
        ChartArea1.AxisX.LabelStyle.Angle = -45
        ChartArea1.AxisX.MajorGrid.LineColor = System.Drawing.Color.LightGray
        ChartArea1.AxisX.MajorGrid.LineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Dot
        ChartArea1.AxisY.LabelStyle.Format = "N0"
        ChartArea1.AxisY.MajorGrid.LineColor = System.Drawing.Color.LightGray
        ChartArea1.AxisY.MajorGrid.LineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Dot
        ChartArea1.BackColor = System.Drawing.Color.White
        ChartArea1.Name = "ChartArea1"
        Me.chartVendasMensais.ChartAreas.Add(ChartArea1)
        Me.chartVendasMensais.Dock = System.Windows.Forms.DockStyle.Fill
        Legend2.Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Bottom
        Legend2.Name = "Legend1"
        Me.chartVendasMensais.Legends.Add(Legend2)
        Me.chartVendasMensais.Location = New System.Drawing.Point(8, 26)
        Me.chartVendasMensais.Name = "chartVendasMensais"
        Series2.BorderWidth = 3
        Series2.ChartArea = "ChartArea1"
        Series2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        Series2.Color = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(76, Byte), Integer), CType(CType(60, Byte), Integer))
        Series2.Legend = "Legend1"
        Series2.MarkerBorderColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(57, Byte), Integer), CType(CType(43, Byte), Integer))
        Series2.MarkerColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(76, Byte), Integer), CType(CType(60, Byte), Integer))
        Series2.MarkerSize = 8
        Series2.MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Circle
        Series2.Name = "Vendas"
        Me.chartVendasMensais.Series.Add(Series2)
        Me.chartVendasMensais.Size = New System.Drawing.Size(343, 318)
        Me.chartVendasMensais.TabIndex = 0
        Me.chartVendasMensais.Text = "Chart2"
        '
        'pnlBotoesSessao
        '
        Me.pnlBotoesSessao.Controls.Add(Me.btnDescartarPedido)
        Me.pnlBotoesSessao.Controls.Add(Me.btnFinalizarPedido)
        Me.pnlBotoesSessao.Controls.Add(Me.btnNovoPedido)
        Me.pnlBotoesSessao.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBotoesSessao.Location = New System.Drawing.Point(10, 730)
        Me.pnlBotoesSessao.Name = "pnlBotoesSessao"
        Me.pnlBotoesSessao.Size = New System.Drawing.Size(1180, 60)
        Me.pnlBotoesSessao.TabIndex = 1
        '
        'btnDescartarPedido
        '
        Me.btnDescartarPedido.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDescartarPedido.BackColor = System.Drawing.Color.FromArgb(CType(CType(108, Byte), Integer), CType(CType(117, Byte), Integer), CType(CType(125, Byte), Integer))
        Me.btnDescartarPedido.FlatAppearance.BorderSize = 0
        Me.btnDescartarPedido.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDescartarPedido.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnDescartarPedido.ForeColor = System.Drawing.Color.White
        Me.btnDescartarPedido.Location = New System.Drawing.Point(1030, 10)
        Me.btnDescartarPedido.Name = "btnDescartarPedido"
        Me.btnDescartarPedido.Size = New System.Drawing.Size(150, 40)
        Me.btnDescartarPedido.TabIndex = 2
        Me.btnDescartarPedido.Text = "🗑️ Descartar Pedido"
        Me.btnDescartarPedido.UseVisualStyleBackColor = False
        '
        'btnFinalizarPedido
        '
        Me.btnFinalizarPedido.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFinalizarPedido.BackColor = System.Drawing.Color.FromArgb(CType(CType(220, Byte), Integer), CType(CType(53, Byte), Integer), CType(CType(69, Byte), Integer))
        Me.btnFinalizarPedido.FlatAppearance.BorderSize = 0
        Me.btnFinalizarPedido.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFinalizarPedido.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnFinalizarPedido.ForeColor = System.Drawing.Color.White
        Me.btnFinalizarPedido.Location = New System.Drawing.Point(874, 10)
        Me.btnFinalizarPedido.Name = "btnFinalizarPedido"
        Me.btnFinalizarPedido.Size = New System.Drawing.Size(150, 40)
        Me.btnFinalizarPedido.TabIndex = 1
        Me.btnFinalizarPedido.Text = "✅ Finalizar Pedido"
        Me.btnFinalizarPedido.UseVisualStyleBackColor = False
        '
        'btnNovoPedido
        '
        Me.btnNovoPedido.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNovoPedido.BackColor = System.Drawing.Color.FromArgb(CType(CType(40, Byte), Integer), CType(CType(167, Byte), Integer), CType(CType(69, Byte), Integer))
        Me.btnNovoPedido.FlatAppearance.BorderSize = 0
        Me.btnNovoPedido.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNovoPedido.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnNovoPedido.ForeColor = System.Drawing.Color.White
        Me.btnNovoPedido.Location = New System.Drawing.Point(718, 10)
        Me.btnNovoPedido.Name = "btnNovoPedido"
        Me.btnNovoPedido.Size = New System.Drawing.Size(150, 40)
        Me.btnNovoPedido.TabIndex = 0
        Me.btnNovoPedido.Text = "➕ Novo Pedido"
        Me.btnNovoPedido.UseVisualStyleBackColor = False
        '
        'UcReposicaoEstoque
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.pnlPrincipal)
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!)
        Me.Name = "UcReposicaoEstoque"
        Me.Size = New System.Drawing.Size(1200, 800)
        Me.pnlPrincipal.ResumeLayout(False)
        Me.splitContainer1.Panel1.ResumeLayout(False)
        Me.splitContainer1.Panel2.ResumeLayout(False)
        CType(Me.splitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitContainer1.ResumeLayout(False)
        Me.pnlEsquerda.ResumeLayout(False)
        Me.grpProdutos.ResumeLayout(False)
        CType(Me.dgvProdutos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFiltros.ResumeLayout(False)
        Me.pnlFiltros.PerformLayout()
        Me.grpInclusaoManual.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.pnlDireita.ResumeLayout(False)
        Me.splitContainer2.Panel1.ResumeLayout(False)
        Me.splitContainer2.Panel2.ResumeLayout(False)
        CType(Me.splitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitContainer2.ResumeLayout(False)
        Me.pnlSuperior.ResumeLayout(False)
        Me.splitContainer3.Panel1.ResumeLayout(False)
        Me.splitContainer3.Panel2.ResumeLayout(False)
        CType(Me.splitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitContainer3.ResumeLayout(False)
        Me.pnlImagemEAplicacao.ResumeLayout(False)
        Me.grpImagem.ResumeLayout(False)
        CType(Me.pbProduto, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAplicacao.ResumeLayout(False)
        Me.grpAplicacao.PerformLayout()
        Me.pnlDireitaSuperior.ResumeLayout(False)
        Me.grpEstoque.ResumeLayout(False)
        CType(Me.dgvEstoque, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpPedidoProduto.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.pnlInferior.ResumeLayout(False)
        Me.splitContainer4.Panel1.ResumeLayout(False)
        Me.splitContainer4.Panel2.ResumeLayout(False)
        CType(Me.splitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitContainer4.ResumeLayout(False)
        Me.grpCompras.ResumeLayout(False)
        CType(Me.chartComprasMensais, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpVendas.ResumeLayout(False)
        CType(Me.chartVendasMensais, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlBotoesSessao.ResumeLayout(False)
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
    Friend WithEvents btnAtualizarPowerQuery As Button
    Friend WithEvents lblUltimaAtualizacao As Label
    Friend WithEvents pnlDireita As Panel
    Friend WithEvents splitContainer2 As SplitContainer
    Friend WithEvents pnlSuperior As Panel
    Friend WithEvents splitContainer3 As SplitContainer
    Friend WithEvents pnlImagemEAplicacao As Panel
    Friend WithEvents grpImagem As GroupBox
    Friend WithEvents pbProduto As PictureBox
    Friend WithEvents grpAplicacao As GroupBox
    Friend WithEvents txtAplicacaoProduto As TextBox
    Friend WithEvents pnlDireitaSuperior As Panel
    Friend WithEvents grpEstoque As GroupBox
    Friend WithEvents dgvEstoque As DataGridView
    Friend WithEvents grpPedidoProduto As GroupBox
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents lblEstoqueMinimo As Label
    Friend WithEvents txtEstoqueMinimo As TextBox
    Friend WithEvents lblEstoqueMaximo As Label
    Friend WithEvents txtEstoqueMaximo As TextBox
    Friend WithEvents lblQtdPedir As Label
    Friend WithEvents txtQtdPedir As TextBox
    Friend WithEvents lblLoja As Label
    Friend WithEvents comboLoja As ComboBox
    Friend WithEvents lblCliente As Label
    Friend WithEvents txtCliente As TextBox
    Friend WithEvents lblTelefone As Label
    Friend WithEvents txtTelefone As MaskedTextBox
    Friend WithEvents btnProximoProduto As Button
    Friend WithEvents pnlInferior As Panel
    Friend WithEvents splitContainer4 As SplitContainer
    Friend WithEvents grpCompras As GroupBox
    Friend WithEvents chartComprasMensais As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents grpVendas As GroupBox
    Friend WithEvents chartVendasMensais As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents grpInclusaoManual As GroupBox
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents lblCodigoManual As Label
    Friend WithEvents txtCodigoManual As TextBox
    Friend WithEvents lblDescricaoManual As Label
    Friend WithEvents txtDescricaoManual As TextBox
    Friend WithEvents lblFabricanteManual As Label
    Friend WithEvents txtFabricanteManual As TextBox
    Friend WithEvents lblQtdInicialManual As Label
    Friend WithEvents txtQtdInicialManual As TextBox
    Friend WithEvents lblDataManual As Label
    Friend WithEvents txtDataManual As TextBox
    Friend WithEvents btnIncluirProdutoManual As Button
    Friend WithEvents pnlBotoesSessao As Panel
    Friend WithEvents btnNovoPedido As Button
    Friend WithEvents btnFinalizarPedido As Button
    Friend WithEvents btnDescartarPedido As Button
End Class