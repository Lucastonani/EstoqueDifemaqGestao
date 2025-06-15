Imports System.Drawing
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports WinFormsApp = System.Windows.Forms.Application

Public Class UcReposicaoEstoque
    Private powerQueryManager As PowerQueryManager
    Private produtoSelecionado As String = String.Empty
    Private debounceTimer As System.Windows.Forms.Timer
    Private filtroTimer As System.Windows.Forms.Timer
    Private filtroAtual As String = String.Empty
    Private dadosProdutosOriginais As System.Data.DataTable
    Private isCarregandoImagem As Boolean = False
    Private dadosCarregados As Boolean = False
    Private carregamentoInicial As Boolean = True

    ' Cache de dados para melhor performance
    Private cacheEstoque As New Dictionary(Of String, System.Data.DataTable)
    Private cacheCompras As New Dictionary(Of String, System.Data.DataTable)
    Private cacheVendas As New Dictionary(Of String, System.Data.DataTable)
    Private cacheValido As DateTime = DateTime.MinValue
    Private Const CACHE_TIMEOUT_MINUTES As Integer = 5

    Public Sub New()
        InitializeComponent()
        ConfigurarComponentes()
        ' NÃO inicializar dados no construtor - fazer lazy loading
    End Sub

    Private Sub ConfigurarComponentes()
        Try
            ' Configurações básicas e rápidas dos DataGridViews
            ConfigurarDataGridViewsBasico()

            ' Configurar PictureBox
            ConfigurarPictureBox()

            ' Configurar timers
            ConfigurarTimers()

            ' Configurar eventos
            ConfigurarEventos()

            ' Configurar controles de filtro
            ConfigurarFiltros()

            ' Mostrar mensagem inicial
            AtualizarStatus("Clique em 'Atualizar' para carregar os dados")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarComponentes")
            MessageBox.Show($"Erro ao configurar componentes: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ConfigurarDataGridViewsBasico()
        Try
            ' Configuração básica e rápida - sem estilização pesada
            For Each dgv As DataGridView In {dgvProdutos, dgvEstoque, dgvCompras, dgvVendas}
                With dgv
                    .AllowUserToAddRows = False
                    .AllowUserToDeleteRows = False
                    .ReadOnly = True
                    .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                    .MultiSelect = False
                    .RowHeadersVisible = False
                    .BackgroundColor = Color.White
                    .BorderStyle = BorderStyle.Fixed3D
                    .EnableHeadersVisualStyles = False
                    .AutoGenerateColumns = True
                    .VirtualMode = False ' Manter false para simplicidade inicial
                    .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                End With
            Next

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarDataGridViewsBasico")
        End Try
    End Sub

    Private Sub ConfigurarTimers()
        ' Timers otimizados com intervalos menores
        debounceTimer = New System.Windows.Forms.Timer()
        debounceTimer.Interval = 100 ' Reduzido de 150ms para 100ms
        AddHandler debounceTimer.Tick, AddressOf DebounceTimer_Tick

        filtroTimer = New System.Windows.Forms.Timer()
        filtroTimer.Interval = 200 ' Reduzido de 300ms para 200ms
        AddHandler filtroTimer.Tick, AddressOf FiltroTimer_Tick
    End Sub

    Private Sub ConfigurarFiltros()
        txtFiltro.Text = ""
        txtFiltro.ForeColor = Color.Black
    End Sub

    Private Sub ConfigurarPictureBox()
        With pbProduto
            .SizeMode = PictureBoxSizeMode.Zoom
            .BorderStyle = BorderStyle.FixedSingle
            .BackColor = Color.White
            .BackgroundImage = Nothing
            .BackgroundImageLayout = ImageLayout.Center
        End With
    End Sub

    ' Lazy loading - só carrega dados quando necessário
    Private Sub InicializarDados()
        Try
            If dadosCarregados Then Return

            ' Obter workbook de forma otimizada
            Dim workbookObj = ObterWorkbookOtimizado()
            If workbookObj Is Nothing Then
                MessageBox.Show("Não foi possível acessar o workbook do Excel.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            powerQueryManager = New PowerQueryManager(workbookObj)
            Me.Cursor = Cursors.WaitCursor

            ' Carregar produtos em background
            Task.Run(Sub()
                         Try
                             Me.Invoke(Sub() CarregarProdutos())
                             dadosCarregados = True
                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, "UcReposicaoEstoque.InicializarDados.Background")
                             Me.Invoke(Sub()
                                           MessageBox.Show($"Erro ao carregar dados: {ex.Message}", "Erro",
                                                         MessageBoxButtons.OK, MessageBoxIcon.Error)
                                       End Sub)
                         Finally
                             Me.Invoke(Sub() Me.Cursor = Cursors.Default)
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.InicializarDados")
            MessageBox.Show($"Erro ao inicializar dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function ObterWorkbookOtimizado() As Microsoft.Office.Interop.Excel.Workbook
        Try
            ' Método otimizado para obter workbook
            If TypeOf Globals.ThisWorkbook Is ThisWorkbook Then
                Dim thisWb As ThisWorkbook = CType(Globals.ThisWorkbook, ThisWorkbook)
                Return thisWb.ObterWorkbook()
            End If

            Return Nothing

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ObterWorkbookOtimizado")
            Return Nothing
        End Try
    End Function

    Private Sub BtnAtualizar_Click(sender As Object, e As EventArgs)
        Try
            ' Lazy loading - inicializar se necessário
            If Not dadosCarregados Then
                InicializarDados()
                Return
            End If

            btnAtualizar.Enabled = False
            btnAtualizar.Text = "🔄 Atualizando..."
            Me.Cursor = Cursors.WaitCursor

            ' Invalidar cache
            InvalidarCache()

            ' Atualizar em background
            Task.Run(Sub()
                         Try
                             AtualizarDadosPowerQuery()

                             Me.Invoke(Sub()
                                           CarregarProdutos()
                                           MessageBox.Show("Dados atualizados com sucesso!", "Sucesso",
                                                         MessageBoxButtons.OK, MessageBoxIcon.Information)
                                       End Sub)

                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, "UcReposicaoEstoque.BtnAtualizar_Click.Background")
                             Me.Invoke(Sub()
                                           MessageBox.Show($"Erro ao atualizar dados: {ex.Message}", "Erro",
                                                         MessageBoxButtons.OK, MessageBoxIcon.Error)
                                       End Sub)
                         Finally
                             Me.Invoke(Sub()
                                           btnAtualizar.Enabled = True
                                           btnAtualizar.Text = "🔄 Atualizar"
                                           Me.Cursor = Cursors.Default
                                       End Sub)
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.BtnAtualizar_Click")
            MessageBox.Show($"Erro ao atualizar dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnAtualizar.Enabled = True
            btnAtualizar.Text = "🔄 Atualizar"
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub InvalidarCache()
        cacheEstoque.Clear()
        cacheCompras.Clear()
        cacheVendas.Clear()
        cacheValido = DateTime.MinValue
    End Sub

    Private Function CacheEstaValido() As Boolean
        Return DateTime.Now.Subtract(cacheValido).TotalMinutes < CACHE_TIMEOUT_MINUTES
    End Function

    Private Sub AtualizarDadosPowerQuery()
        Try
            If powerQueryManager IsNot Nothing Then
                powerQueryManager.AtualizarTodasConsultas()
                LogErros.RegistrarInfo("Power Query atualizado com sucesso", "UcReposicaoEstoque.AtualizarDadosPowerQuery")
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.AtualizarDadosPowerQuery")
            Throw New Exception($"Erro ao atualizar Power Query: {ex.Message}")
        End Try
    End Sub

    Private Sub CarregarProdutos()
        Try
            If powerQueryManager Is Nothing Then
                MessageBox.Show("PowerQueryManager não está inicializado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Me.Cursor = Cursors.WaitCursor
            dgvProdutos.SuspendLayout()

            Try
                Dim tabelaProdutos As ListObject = powerQueryManager.ObterTabela(ConfiguracaoApp.TABELA_PRODUTOS)
                If tabelaProdutos Is Nothing Then
                    MessageBox.Show($"Tabela '{ConfiguracaoApp.TABELA_PRODUTOS}' não encontrada!", "Aviso",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                ' Converter para DataTable e armazenar
                dadosProdutosOriginais = DataHelper.ConvertListObjectToDataTable(tabelaProdutos)

                ' Aplicar filtro se existir
                AplicarFiltro()

                ' Configurar colunas otimizado
                ConfigurarColunasProdutosOtimizado()

                ' Aplicar estilização apenas após carregar dados
                dgvProdutos.EstilizarDataGridView()

                ' Selecionar primeiro produto se existir
                If dgvProdutos.Rows.Count > 0 Then
                    dgvProdutos.Rows(0).Selected = True
                End If

                AtualizarStatus($"Produtos carregados: {dgvProdutos.Rows.Count} registros")

            Finally
                dgvProdutos.ResumeLayout(True)
                Me.Cursor = Cursors.Default
            End Try

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.CarregarProdutos")
            MessageBox.Show($"Erro ao carregar produtos: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub ConfigurarColunasProdutosOtimizado()
        Try
            If dgvProdutos.Columns.Count >= 7 Then
                dgvProdutos.ConfigurarColunas(
                    New ColumnConfig(0, "Código", 80),
                    New ColumnConfig(1, "Descrição", 250),
                    New ColumnConfig(2, "Fabricante", 150),
                    New ColumnConfig(3, "Tipo", 100),
                    New ColumnConfig(4, "Custo", 80, True, DataGridViewContentAlignment.MiddleRight, "C2"),
                    New ColumnConfig(5, "Curva", 60),
                    New ColumnConfig(6, "Estoque", 80, True, DataGridViewContentAlignment.MiddleRight, "N2")
                )
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarColunasProdutosOtimizado")
        End Try
    End Sub

    ' Carregamento otimizado com cache
    Private Sub CarregarDadosProdutoOtimizado(codigoProduto As String)
        Try
            LimparDadosSecundarios()

            ' Usar cursor apenas se necessário
            Dim usarCursor = Not CacheEstaValido()
            If usarCursor Then Me.Cursor = Cursors.WaitCursor

            ' Suspender layouts para performance
            dgvEstoque.SuspendLayout()
            dgvCompras.SuspendLayout()
            dgvVendas.SuspendLayout()

            Try
                ' Carregar dados com cache
                Dim estoqueData = CarregarDadosComCache("estoque", codigoProduto, ConfiguracaoApp.TABELA_ESTOQUE)
                dgvEstoque.DataSource = estoqueData
                ConfigurarColunasEstoqueOtimizado()
                grpEstoque.Text = $"📊 Estoque Atual ({estoqueData.Rows.Count} registros)"

                Dim comprasData = CarregarDadosComCache("compras", codigoProduto, ConfiguracaoApp.TABELA_COMPRAS)
                dgvCompras.DataSource = comprasData
                ConfigurarColunasComprasOtimizado()
                grpCompras.Text = $"📈 Compras ({comprasData.Rows.Count} registros)"

                Dim vendasData = CarregarDadosComCache("vendas", codigoProduto, ConfiguracaoApp.TABELA_VENDAS)
                dgvVendas.DataSource = vendasData
                ConfigurarColunasVendasOtimizado()
                grpVendas.Text = $"📉 Vendas ({vendasData.Rows.Count} registros)"

                ' Aplicar estilização apenas se necessário
                If usarCursor Then
                    dgvEstoque.EstilizarDataGridView()
                    dgvCompras.EstilizarDataGridView()
                    dgvVendas.EstilizarDataGridView()
                End If

                ' Carregar imagem de forma assíncrona
                CarregarImagemProdutoAsync(codigoProduto)

            Finally
                dgvEstoque.ResumeLayout(True)
                dgvCompras.ResumeLayout(True)
                dgvVendas.ResumeLayout(True)
                If usarCursor Then Me.Cursor = Cursors.Default
            End Try

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.CarregarDadosProdutoOtimizado")
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function CarregarDadosComCache(tipo As String, codigoProduto As String, nomeTabela As String) As System.Data.DataTable
        Try
            Dim chaveCache = $"{tipo}_{codigoProduto}"

            ' Verificar cache
            If CacheEstaValido() Then
                Select Case tipo
                    Case "estoque"
                        If cacheEstoque.ContainsKey(chaveCache) Then
                            Return cacheEstoque(chaveCache)
                        End If
                    Case "compras"
                        If cacheCompras.ContainsKey(chaveCache) Then
                            Return cacheCompras(chaveCache)
                        End If
                    Case "vendas"
                        If cacheVendas.ContainsKey(chaveCache) Then
                            Return cacheVendas(chaveCache)
                        End If
                End Select
            End If

            ' Carregar dados
            Dim dados = CarregarDadosFiltrados(nomeTabela, codigoProduto)

            ' Armazenar no cache
            Select Case tipo
                Case "estoque"
                    cacheEstoque(chaveCache) = dados
                Case "compras"
                    cacheCompras(chaveCache) = dados
                Case "vendas"
                    cacheVendas(chaveCache) = dados
            End Select

            ' Atualizar timestamp do cache
            cacheValido = DateTime.Now

            Return dados

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"UcReposicaoEstoque.CarregarDadosComCache({tipo}, {codigoProduto})")
            Return New System.Data.DataTable()
        End Try
    End Function

    Private Function CarregarDadosFiltrados(nomeTabela As String, codigoProduto As String) As System.Data.DataTable
        Try
            If powerQueryManager Is Nothing Then
                Return New System.Data.DataTable()
            End If

            Dim tabela As ListObject = powerQueryManager.ObterTabela(nomeTabela)
            If tabela Is Nothing Then
                Return New System.Data.DataTable()
            End If

            Dim dataTable As System.Data.DataTable = DataHelper.ConvertListObjectToDataTable(tabela)

            ' Filtrar por código do produto de forma otimizada
            If dataTable.Columns.Count > 0 Then
                Return DataHelper.FiltrarDataTable(dataTable, dataTable.Columns(0).ColumnName, codigoProduto)
            End If

            Return dataTable

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"UcReposicaoEstoque.CarregarDadosFiltrados({nomeTabela}, {codigoProduto})")
            Return New System.Data.DataTable()
        End Try
    End Function

    ' Carregamento assíncrono de imagem otimizado
    Private Sub CarregarImagemProdutoAsync(codigoProduto As String)
        If isCarregandoImagem Then Return

        Try
            isCarregandoImagem = True
            grpImagem.Text = "🖼️ Imagem do Produto - Carregando..."

            Task.Run(Sub()
                         Try
                             Dim imagemEncontrada As Boolean = False
                             Dim imagemCarregada As Image = Nothing

                             ' Procurar imagem com diferentes extensões
                             For Each extensao As String In ConfiguracaoApp.EXTENSOES_IMAGEM
                                 Dim caminhoImagem = Path.Combine(ConfiguracaoApp.CAMINHO_IMAGENS, $"{codigoProduto}{extensao}")

                                 If File.Exists(caminhoImagem) Then
                                     Try
                                         Dim fileInfo As New FileInfo(caminhoImagem)
                                         If fileInfo.Length <= ConfiguracaoApp.TAMANHO_MAXIMO_IMAGEM Then
                                             Using fs As New FileStream(caminhoImagem, FileMode.Open, FileAccess.Read)
                                                 imagemCarregada = Image.FromStream(fs)
                                             End Using
                                             imagemEncontrada = True
                                             Exit For
                                         End If
                                     Catch
                                         Continue For
                                     End Try
                                 End If
                             Next

                             ' Atualizar UI no thread principal
                             Me.Invoke(Sub()
                                           If pbProduto.Image IsNot Nothing Then
                                               pbProduto.Image.Dispose()
                                               pbProduto.Image = Nothing
                                           End If

                                           If imagemEncontrada AndAlso imagemCarregada IsNot Nothing Then
                                               pbProduto.Image = imagemCarregada
                                               grpImagem.Text = "🖼️ Imagem do Produto"
                                           Else
                                               grpImagem.Text = "🖼️ Imagem do Produto - Não disponível"
                                           End If
                                       End Sub)

                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, $"UcReposicaoEstoque.CarregarImagemProdutoAsync({codigoProduto})")
                             Me.Invoke(Sub()
                                           pbProduto.Image = Nothing
                                           grpImagem.Text = "🖼️ Imagem do Produto - Erro"
                                       End Sub)
                         Finally
                             isCarregandoImagem = False
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"UcReposicaoEstoque.CarregarImagemProdutoAsync({codigoProduto}) - Outer")
            isCarregandoImagem = False
            grpImagem.Text = "🖼️ Imagem do Produto - Erro"
        End Try
    End Sub

    ' Resto dos métodos de configuração otimizados...
    Private Sub ConfigurarColunasEstoqueOtimizado()
        Try
            If dgvEstoque.DataSource IsNot Nothing AndAlso dgvEstoque.Columns.Count >= 9 Then
                dgvEstoque.ConfigurarColunas(
                    New ColumnConfig(0, "Código", 70),
                    New ColumnConfig(1, "Loja", 100),
                    New ColumnConfig(2, "Disponível", 80, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(3, "Pré-Venda", 80, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(4, "Conta Cliente", 90, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(5, "Transf.Pend.", 90, True, DataGridViewContentAlignment.MiddleCenter),
                    New ColumnConfig(6, "Em Trânsito", 90, True, DataGridViewContentAlignment.MiddleCenter),
                    New ColumnConfig(7, "Ped.Compra", 90, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(8, "Total", 80, True, DataGridViewContentAlignment.MiddleRight, "N2")
                )
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarColunasEstoqueOtimizado")
        End Try
    End Sub

    Private Sub ConfigurarColunasComprasOtimizado()
        Try
            If dgvCompras.DataSource IsNot Nothing AndAlso dgvCompras.Columns.Count >= 7 Then
                dgvCompras.ConfigurarColunas(
                    New ColumnConfig(0, "Código", 70),
                    New ColumnConfig(1, "Data", 80),
                    New ColumnConfig(2, "Cariacica", 85, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(3, "Serra", 85, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(4, "Linhares", 85, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(5, "Marechal", 85, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(6, "Atacado", 85, True, DataGridViewContentAlignment.MiddleRight, "N2")
                )
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarColunasComprasOtimizado")
        End Try
    End Sub

    Private Sub ConfigurarColunasVendasOtimizado()
        Try
            If dgvVendas.DataSource IsNot Nothing AndAlso dgvVendas.Columns.Count >= 7 Then
                dgvVendas.ConfigurarColunas(
                    New ColumnConfig(0, "Código", 70),
                    New ColumnConfig(1, "Data", 80),
                    New ColumnConfig(2, "Cariacica", 85, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(3, "Serra", 85, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(4, "Linhares", 85, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(5, "Marechal", 85, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(6, "Atacado", 85, True, DataGridViewContentAlignment.MiddleRight, "N2")
                )
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarColunasVendasOtimizado")
        End Try
    End Sub

    ' Eventos otimizados
    Private Sub TxtFiltro_TextChanged(sender As Object, e As EventArgs)
        Try
            If filtroTimer IsNot Nothing Then
                filtroTimer.Stop()
                filtroTimer.Start()
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.TxtFiltro_TextChanged")
        End Try
    End Sub

    Private Sub FiltroTimer_Tick(sender As Object, e As EventArgs)
        If filtroTimer IsNot Nothing Then
            filtroTimer.Stop()
        End If
        filtroAtual = txtFiltro.Text.Trim()
        AplicarFiltro()
    End Sub

    Private Sub DgvProdutos_SelectionChanged(sender As Object, e As EventArgs)
        Try
            If debounceTimer IsNot Nothing Then
                debounceTimer.Stop()
                debounceTimer.Start()
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DgvProdutos_SelectionChanged")
        End Try
    End Sub

    Private Sub DebounceTimer_Tick(sender As Object, e As EventArgs)
        Try
            If debounceTimer IsNot Nothing Then
                debounceTimer.Stop()
            End If

            If dgvProdutos.SelectedRows.Count > 0 Then
                Dim produtoSelecionadoRow As DataGridViewRow = dgvProdutos.SelectedRows(0)

                If produtoSelecionadoRow.Cells.Count > 0 Then
                    Dim codigoProduto As String = If(produtoSelecionadoRow.Cells(0).Value IsNot Nothing,
                                                produtoSelecionadoRow.Cells(0).Value.ToString(), "")

                    If Not String.IsNullOrEmpty(codigoProduto) AndAlso codigoProduto <> produtoSelecionado Then
                        produtoSelecionado = codigoProduto

                        ' Carregar dados de forma otimizada
                        CarregarDadosProdutoOtimizado(codigoProduto)
                    End If
                End If
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DebounceTimer_Tick")
        End Try
    End Sub

    Private Sub AplicarFiltro()
        Try
            If dadosProdutosOriginais Is Nothing Then Return

            ' Aplicar filtro de forma otimizada
            Dim dadosFiltrados As System.Data.DataTable = dadosProdutosOriginais.Clone()

            If String.IsNullOrEmpty(filtroAtual) Then
                For Each row As System.Data.DataRow In dadosProdutosOriginais.Rows
                    dadosFiltrados.ImportRow(row)
                Next
            Else
                ' Filtro otimizado - parar na primeira correspondência por linha
                For Each row As System.Data.DataRow In dadosProdutosOriginais.Rows
                    Dim incluirRow As Boolean = False

                    For Each column As System.Data.DataColumn In dadosProdutosOriginais.Columns
                        If column.DataType = GetType(String) Then
                            Dim valorCelula As String = If(row(column) IsNot Nothing, row(column).ToString(), "")
                            If valorCelula.IndexOf(filtroAtual, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                incluirRow = True
                                Exit For ' Parar na primeira correspondência
                            End If
                        End If
                    Next

                    If incluirRow Then
                        dadosFiltrados.ImportRow(row)
                    End If
                Next
            End If

            dgvProdutos.DataSource = dadosFiltrados
            AtualizarContadorProdutos(dadosFiltrados.Rows.Count)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.AplicarFiltro")
        End Try
    End Sub

    Private Sub AtualizarContadorProdutos(quantidade As Integer)
        Try
            grpProdutos.Text = $"📦 Lista de Produtos ({quantidade} registros)"
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.AtualizarContadorProdutos")
        End Try
    End Sub

    Private Sub LimparDadosSecundarios()
        Try
            dgvEstoque.DataSource = Nothing
            dgvCompras.DataSource = Nothing
            dgvVendas.DataSource = Nothing

            grpEstoque.Text = "📊 Estoque Atual"
            grpCompras.Text = "📈 Compras"
            grpVendas.Text = "📉 Vendas"

            If pbProduto.Image IsNot Nothing Then
                pbProduto.Image.Dispose()
                pbProduto.Image = Nothing
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.LimparDadosSecundarios")
        End Try
    End Sub

    Private Sub ConfigurarEventos()
        Try
            AddHandler dgvProdutos.SelectionChanged, AddressOf DgvProdutos_SelectionChanged
            AddHandler dgvProdutos.CellDoubleClick, AddressOf DgvProdutos_CellDoubleClick
            AddHandler dgvProdutos.DataBindingComplete, AddressOf DgvProdutos_DataBindingComplete
            AddHandler txtFiltro.TextChanged, AddressOf TxtFiltro_TextChanged
            AddHandler txtFiltro.Enter, AddressOf TxtFiltro_Enter
            AddHandler txtFiltro.Leave, AddressOf TxtFiltro_Leave
            AddHandler btnAtualizar.Click, AddressOf BtnAtualizar_Click
            AddHandler dgvProdutos.KeyDown, AddressOf DgvProdutos_KeyDown

            LogErros.RegistrarInfo("Eventos configurados", "UcReposicaoEstoque.ConfigurarEventos")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarEventos")
        End Try
    End Sub

    Private Sub DgvProdutos_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        Try
            If dgvProdutos.Columns.Count > 0 Then
                dgvProdutos.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells)
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DgvProdutos_DataBindingComplete")
        End Try
    End Sub

    Private Sub TxtFiltro_Enter(sender As Object, e As EventArgs)
        Try
            txtFiltro.SelectAll()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.TxtFiltro_Enter")
        End Try
    End Sub

    Private Sub TxtFiltro_Leave(sender As Object, e As EventArgs)
        ' Placeholder para funcionalidade futura
    End Sub

    Private Sub DgvProdutos_KeyDown(sender As Object, e As KeyEventArgs)
        Try
            Select Case e.KeyCode
                Case Keys.Enter
                    If dgvProdutos.SelectedRows.Count > 0 Then
                        DgvProdutos_CellDoubleClick(sender, New DataGridViewCellEventArgs(0, dgvProdutos.SelectedRows(0).Index))
                    End If
                    e.Handled = True

                Case Keys.F5
                    btnAtualizar.PerformClick()
                    e.Handled = True

                Case Keys.F, Keys.F3
                    If e.Control OrElse e.KeyCode = Keys.F3 Then
                        txtFiltro.Focus()
                        e.Handled = True
                    End If
            End Select

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DgvProdutos_KeyDown")
        End Try
    End Sub

    Private Sub DgvProdutos_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            If e.RowIndex >= 0 Then
                Dim row = dgvProdutos.Rows(e.RowIndex)
                Dim info As New System.Text.StringBuilder()

                For Each col As DataGridViewColumn In dgvProdutos.Columns
                    If col.Visible Then
                        Dim valor = If(row.Cells(col.Index).Value IsNot Nothing, row.Cells(col.Index).Value.ToString(), "")
                        info.AppendLine($"{col.HeaderText}: {valor}")
                    End If
                Next

                If dgvEstoque.Rows.Count > 0 Then
                    info.AppendLine()
                    info.AppendLine("=== ESTOQUE ===")
                    Dim totalEstoque As Decimal = 0

                    For Each estoqueRow As DataGridViewRow In dgvEstoque.Rows
                        If estoqueRow.Cells.Count > 8 AndAlso IsNumeric(estoqueRow.Cells(8).Value) Then
                            totalEstoque += Convert.ToDecimal(estoqueRow.Cells(8).Value)
                        End If
                    Next

                    info.AppendLine($"Total em Estoque: {totalEstoque:N2}")
                End If

                MessageBox.Show(info.ToString(), "Detalhes do Produto", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DgvProdutos_CellDoubleClick")
        End Try
    End Sub

    Private Sub AtualizarStatus(mensagem As String)
        Try
            ' Se tiver referência ao form pai, atualizar status
            Dim parentForm = Me.FindForm()
            If TypeOf parentForm Is MainForm Then
                CType(parentForm, MainForm).AtualizarStatus(mensagem)
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.AtualizarStatus")
        End Try
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing Then
                ' Limpar timers
                If debounceTimer IsNot Nothing Then
                    debounceTimer.Stop()
                    debounceTimer.Dispose()
                    debounceTimer = Nothing
                End If

                If filtroTimer IsNot Nothing Then
                    filtroTimer.Stop()
                    filtroTimer.Dispose()
                    filtroTimer = Nothing
                End If

                ' Limpar imagem
                If pbProduto IsNot Nothing AndAlso pbProduto.Image IsNot Nothing Then
                    pbProduto.Image.Dispose()
                End If

                ' Limpar dados
                If dadosProdutosOriginais IsNot Nothing Then
                    dadosProdutosOriginais.Dispose()
                    dadosProdutosOriginais = Nothing
                End If

                ' Limpar cache
                InvalidarCache()

                powerQueryManager = Nothing
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

End Class