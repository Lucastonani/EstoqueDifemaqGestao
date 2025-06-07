Imports System.Drawing
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel

Public Class UcReposicaoEstoque
    Private powerQueryManager As PowerQueryManager
    Private produtoSelecionado As String = String.Empty
    Private debounceTimer As Timer
    Private filtroTimer As Timer
    Private filtroAtual As String = String.Empty
    Private dadosProdutosOriginais As System.Data.DataTable
    Private isCarregandoImagem As Boolean = False

    Public Sub New()
        InitializeComponent()
        ConfigurarComponentes()
        InicializarDados()
    End Sub

    Private Sub ConfigurarComponentes()
        Try
            ' Aplicar estilização aos DataGridViews
            dgvProdutos.EstilizarDataGridView()
            dgvEstoque.EstilizarDataGridView()
            dgvCompras.EstilizarDataGridView()
            dgvVendas.EstilizarDataGridView()

            ' Configurar PictureBox
            ConfigurarPictureBox()

            ' Configurar timers
            ConfigurarTimers()

            ' Configurar eventos
            ConfigurarEventos()

            ' Configurar controles de filtro
            ConfigurarFiltros()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarComponentes")
            MessageBox.Show(String.Format("Erro ao configurar componentes: {0}", ex.Message), "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ConfigurarTimers()
        ' Timer para debounce da seleção de produtos
        debounceTimer = New Timer()
        debounceTimer.Interval = ConfiguracaoApp.DEBOUNCE_DELAY
        AddHandler debounceTimer.Tick, AddressOf DebounceTimer_Tick

        ' Timer para filtro de produtos
        filtroTimer = New Timer()
        filtroTimer.Interval = 500 ' 500ms para filtro
        AddHandler filtroTimer.Tick, AddressOf FiltroTimer_Tick
    End Sub

    Private Sub ConfigurarEventos()
        AddHandler dgvProdutos.SelectionChanged, AddressOf DgvProdutos_SelectionChanged
        AddHandler dgvProdutos.CellDoubleClick, AddressOf DgvProdutos_CellDoubleClick
        AddHandler txtFiltro.TextChanged, AddressOf TxtFiltro_TextChanged
        AddHandler btnAtualizar.Click, AddressOf BtnAtualizar_Click
    End Sub

    Private Sub ConfigurarFiltros()
        ' Configuração básica do filtro - PlaceholderText não disponível em versões antigas
        txtFiltro.Text = ""
        txtFiltro.ForeColor = Color.Black
    End Sub

    Private Sub InicializarDados()
        Try
            powerQueryManager = New PowerQueryManager(CType(Globals.ThisWorkbook, Microsoft.Office.Interop.Excel.Workbook))

            ' Mostrar cursor de espera
            Me.Cursor = Cursors.WaitCursor

            ' Carregar produtos inicialmente
            CarregarProdutos()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.InicializarDados")
            MessageBox.Show(String.Format("Erro ao inicializar dados: {0}", ex.Message), "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnAtualizar_Click(sender As Object, e As EventArgs)
        Try
            btnAtualizar.Enabled = False
            btnAtualizar.Text = "🔄 Atualizando..."
            Me.Cursor = Cursors.WaitCursor

            ' Atualizar Power Query em thread separada para evitar travamento
            Dim atualizarTask As Task = Task.Run(Sub() AtualizarDadosPowerQuery())
            atualizarTask.Wait()

            ' Recarregar produtos
            CarregarProdutos()

            MessageBox.Show("Dados atualizados com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.BtnAtualizar_Click")
            MessageBox.Show(String.Format("Erro ao atualizar dados: {0}", ex.Message), "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            btnAtualizar.Enabled = True
            btnAtualizar.Text = "🔄 Atualizar"
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub AtualizarDadosPowerQuery()
        Try
            powerQueryManager.AtualizarTodasConsultas()
            LogErros.RegistrarInfo("Power Query atualizado com sucesso", "UcReposicaoEstoque.AtualizarDadosPowerQuery")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.AtualizarDadosPowerQuery")
            Throw New Exception(String.Format("Erro ao atualizar Power Query: {0}", ex.Message))
        End Try
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

    Private Sub CarregarProdutos()
        Try
            Dim tabelaProdutos As ListObject = powerQueryManager.ObterTabela(ConfiguracaoApp.TABELA_PRODUTOS)
            If tabelaProdutos Is Nothing Then
                MessageBox.Show(String.Format("Tabela '{0}' não encontrada!", ConfiguracaoApp.TABELA_PRODUTOS), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Converter para DataTable e armazenar cópia original
            dadosProdutosOriginais = DataHelper.ConvertListObjectToDataTable(tabelaProdutos)

            ' Aplicar filtro se existir
            AplicarFiltro()

            ' Configurar colunas
            ConfigurarColunasProdutos()

            ' Selecionar primeiro produto se existir
            If dgvProdutos.Rows.Count > 0 Then
                dgvProdutos.Rows(0).Selected = True
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.CarregarProdutos")
            MessageBox.Show(String.Format("Erro ao carregar produtos: {0}", ex.Message), "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxtFiltro_TextChanged(sender As Object, e As EventArgs)
        filtroTimer.Stop()
        filtroTimer.Start()
    End Sub

    Private Sub FiltroTimer_Tick(sender As Object, e As EventArgs)
        filtroTimer.Stop()
        filtroAtual = txtFiltro.Text.Trim()
        AplicarFiltro()
    End Sub

    Private Sub AplicarFiltro()
        Try
            If dadosProdutosOriginais Is Nothing Then Return

            Dim dadosFiltrados As System.Data.DataTable = dadosProdutosOriginais.Clone()

            If String.IsNullOrEmpty(filtroAtual) Then
                ' Sem filtro, mostrar todos os dados
                For Each row As System.Data.DataRow In dadosProdutosOriginais.Rows
                    dadosFiltrados.ImportRow(row)
                Next
            Else
                ' Aplicar filtro em todas as colunas de texto
                For Each row As System.Data.DataRow In dadosProdutosOriginais.Rows
                    Dim incluirRow As Boolean = False

                    For Each column As System.Data.DataColumn In dadosProdutosOriginais.Columns
                        If column.DataType = GetType(String) Then
                            Dim valorCelula As String = If(row(column) IsNot Nothing, row(column).ToString(), "")
                            If valorCelula.IndexOf(filtroAtual, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                incluirRow = True
                                Exit For
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
            grpProdutos.Text = String.Format("📦 Lista de Produtos ({0} registros)", quantidade)
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.AtualizarContadorProdutos")
        End Try
    End Sub

    Private Sub ConfigurarColunasProdutos()
        Try
            If dgvProdutos.Columns.Count > 0 Then
                ' Configuração específica baseada nas colunas esperadas
                If dgvProdutos.Columns.Count >= 5 Then
                    dgvProdutos.ConfigurarColunas(
                        New ColumnConfig(0, "Código", 100),
                        New ColumnConfig(1, "Descrição", 300),
                        New ColumnConfig(2, "Categoria", 150),
                        New ColumnConfig(3, "Unidade", 80),
                        New ColumnConfig(4, "Preço", 100, True, DataGridViewContentAlignment.MiddleRight, "C2")
                    )
                Else
                    ' Configuração genérica
                    dgvProdutos.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                    dgvProdutos.Columns(0).MinimumWidth = 80

                    For i As Integer = 1 To dgvProdutos.Columns.Count - 1
                        dgvProdutos.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                    Next
                End If
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarColunasProdutos")
        End Try
    End Sub

    Private Sub DgvProdutos_SelectionChanged(sender As Object, e As EventArgs)
        Try
            ' Usar debounce para evitar múltiplas chamadas
            debounceTimer.Stop()
            debounceTimer.Start()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DgvProdutos_SelectionChanged")
        End Try
    End Sub

    Private Sub DebounceTimer_Tick(sender As Object, e As EventArgs)
        Try
            debounceTimer.Stop()

            If dgvProdutos.SelectedRows.Count > 0 Then
                Dim produtoSelecionadoRow As DataGridViewRow = dgvProdutos.SelectedRows(0)
                Dim codigoProduto As String = If(produtoSelecionadoRow.Cells(0).Value IsNot Nothing, produtoSelecionadoRow.Cells(0).Value.ToString(), "")

                If Not String.IsNullOrEmpty(codigoProduto) AndAlso codigoProduto <> produtoSelecionado Then
                    produtoSelecionado = codigoProduto
                    CarregarDadosProdutoAsync(codigoProduto)
                End If
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DebounceTimer_Tick")
        End Try
    End Sub

    Private Sub DgvProdutos_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            If e.RowIndex >= 0 Then
                Dim produto As String = If(dgvProdutos.Rows(e.RowIndex).Cells(0).Value IsNot Nothing, dgvProdutos.Rows(e.RowIndex).Cells(0).Value.ToString(), "")
                Dim descricao As String = ""

                If dgvProdutos.Columns.Count > 1 Then
                    descricao = If(dgvProdutos.Rows(e.RowIndex).Cells(1).Value IsNot Nothing, dgvProdutos.Rows(e.RowIndex).Cells(1).Value.ToString(), "")
                End If

                ' Criar formulário de detalhes mais elaborado
                Dim detalhes As New System.Text.StringBuilder()
                detalhes.AppendLine(String.Format("Produto: {0}", produto))
                detalhes.AppendLine(String.Format("Descrição: {0}", descricao))

                ' Adicionar informações de estoque se disponível
                If dgvEstoque.Rows.Count > 0 Then
                    Dim totalEstoque As Decimal = 0
                    For Each row As DataGridViewRow In dgvEstoque.Rows
                        If row.Cells.Count > 2 AndAlso IsNumeric(row.Cells(2).Value) Then
                            totalEstoque += Convert.ToDecimal(row.Cells(2).Value)
                        End If
                    Next
                    detalhes.AppendLine(String.Format("Estoque Total: {0:N2}", totalEstoque))
                End If

                MessageBox.Show(detalhes.ToString(), "Detalhes do Produto", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DgvProdutos_CellDoubleClick")
        End Try
    End Sub

    Private Sub CarregarDadosProdutoAsync(codigoProduto As String)
        Try
            ' Limpar dados anteriores
            LimparDadosSecundarios()

            ' Mostrar indicadores de carregamento
            MostrarIndicadoresCarregamento(True)

            ' Carregar dados de forma assíncrona
            Task.Run(Sub()
                         Try
                             ' Carregar estoque
                             Dim estoqueData = CarregarDadosFiltrados(ConfiguracaoApp.TABELA_ESTOQUE, codigoProduto)
                             Me.Invoke(Sub()
                                           dgvEstoque.DataSource = estoqueData
                                           ConfigurarColunasEstoque()
                                           grpEstoque.Text = String.Format("📊 Estoque Atual ({0} registros)", estoqueData.Rows.Count)
                                       End Sub)

                             ' Carregar compras
                             Dim comprasData = CarregarDadosFiltrados(ConfiguracaoApp.TABELA_COMPRAS, codigoProduto)
                             Me.Invoke(Sub()
                                           dgvCompras.DataSource = comprasData
                                           ConfigurarColunasCompras()
                                           grpCompras.Text = String.Format("📈 Compras ({0} registros)", comprasData.Rows.Count)
                                       End Sub)

                             ' Carregar vendas
                             Dim vendasData = CarregarDadosFiltrados(ConfiguracaoApp.TABELA_VENDAS, codigoProduto)
                             Me.Invoke(Sub()
                                           dgvVendas.DataSource = vendasData
                                           ConfigurarColunasVendas()
                                           grpVendas.Text = String.Format("📉 Vendas ({0} registros)", vendasData.Rows.Count)
                                       End Sub)

                             ' Carregar imagem por último
                             Me.Invoke(Sub() CarregarImagemProdutoAsync(codigoProduto))

                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, String.Format("UcReposicaoEstoque.CarregarDadosProdutoAsync({0})", codigoProduto))
                             Me.Invoke(Sub()
                                           MostrarIndicadoresCarregamento(False)
                                           MessageBox.Show(String.Format("Erro ao carregar dados do produto: {0}", ex.Message), "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                       End Sub)
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.CarregarDadosProdutoAsync")
            MostrarIndicadoresCarregamento(False)
        End Try
    End Sub

    Private Sub ConfigurarColunasEstoque()
        Try
            If dgvEstoque.DataSource IsNot Nothing AndAlso dgvEstoque.Columns.Count >= 5 Then
                dgvEstoque.ConfigurarColunas(
                    New ColumnConfig(0, "Produto", 120),
                    New ColumnConfig(1, "Local", 150),
                    New ColumnConfig(2, "Quantidade", 100, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(3, "Mín.", 80, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(4, "Máx.", 80, True, DataGridViewContentAlignment.MiddleRight, "N2")
                )
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarColunasEstoque")
        End Try
    End Sub

    Private Sub ConfigurarColunasCompras()
        Try
            If dgvCompras.DataSource IsNot Nothing AndAlso dgvCompras.Columns.Count >= 6 Then
                dgvCompras.ConfigurarColunas(
                    New ColumnConfig(0, "Produto", 120),
                    New ColumnConfig(1, "Data", 100),
                    New ColumnConfig(2, "Fornecedor", 200),
                    New ColumnConfig(3, "Qtd.", 80, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(4, "Valor Unit.", 100, True, DataGridViewContentAlignment.MiddleRight, "C2"),
                    New ColumnConfig(5, "Total", 100, True, DataGridViewContentAlignment.MiddleRight, "C2")
                )
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarColunasCompras")
        End Try
    End Sub

    Private Sub ConfigurarColunasVendas()
        Try
            If dgvVendas.DataSource IsNot Nothing AndAlso dgvVendas.Columns.Count >= 6 Then
                dgvVendas.ConfigurarColunas(
                    New ColumnConfig(0, "Produto", 120),
                    New ColumnConfig(1, "Data", 100),
                    New ColumnConfig(2, "Cliente", 200),
                    New ColumnConfig(3, "Qtd.", 80, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(4, "Valor Unit.", 100, True, DataGridViewContentAlignment.MiddleRight, "C2"),
                    New ColumnConfig(5, "Total", 100, True, DataGridViewContentAlignment.MiddleRight, "C2")
                )
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarColunasVendas")
        End Try
    End Sub

    Private Sub LimparDadosSecundarios()
        Try
            dgvEstoque.DataSource = Nothing
            dgvCompras.DataSource = Nothing
            dgvVendas.DataSource = Nothing

            ' Restaurar textos dos grupos
            grpEstoque.Text = "📊 Estoque Atual"
            grpCompras.Text = "📈 Compras"
            grpVendas.Text = "📉 Vendas"

            ' Limpar imagem
            If pbProduto.Image IsNot Nothing Then
                pbProduto.Image.Dispose()
                pbProduto.Image = Nothing
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.LimparDadosSecundarios")
        End Try
    End Sub

    Private Sub MostrarIndicadoresCarregamento(mostrar As Boolean)
        Try
            If mostrar Then
                grpEstoque.Text = "📊 Estoque Atual - Carregando..."
                grpCompras.Text = "📈 Compras - Carregando..."
                grpVendas.Text = "📉 Vendas - Carregando..."
                grpImagem.Text = "🖼️ Imagem do Produto - Carregando..."
            Else
                grpImagem.Text = "🖼️ Imagem do Produto"
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.MostrarIndicadoresCarregamento")
        End Try
    End Sub

    Private Function CarregarDadosFiltrados(nomeTabela As String, codigoProduto As String) As System.Data.DataTable
        Try
            Dim tabela As ListObject = powerQueryManager.ObterTabela(nomeTabela)
            If tabela Is Nothing Then Return New System.Data.DataTable()

            Dim dataTable As System.Data.DataTable = DataHelper.ConvertListObjectToDataTable(tabela)

            ' Filtrar por código do produto (assumindo que a primeira coluna é o código)
            If dataTable.Columns.Count > 0 Then
                Return DataHelper.FiltrarDataTable(dataTable, dataTable.Columns(0).ColumnName, codigoProduto)
            End If

            Return dataTable

        Catch ex As Exception
            LogErros.RegistrarErro(ex, String.Format("UcReposicaoEstoque.CarregarDadosFiltrados({0}, {1})", nomeTabela, codigoProduto))
            Return New System.Data.DataTable()
        End Try
    End Function

    Private Sub CarregarImagemProdutoAsync(codigoProduto As String)
        If isCarregandoImagem Then Return

        Try
            isCarregandoImagem = True

            Task.Run(Sub()
                         Try
                             ' Limpar imagem anterior no thread principal
                             Me.Invoke(Sub()
                                           If pbProduto.Image IsNot Nothing Then
                                               pbProduto.Image.Dispose()
                                               pbProduto.Image = Nothing
                                           End If
                                       End Sub)

                             ' Procurar imagem com diferentes extensões
                             Dim imagemEncontrada As Boolean = False
                             Dim imagemCarregada As Image = Nothing

                             For Each extensao As String In ConfiguracaoApp.EXTENSOES_IMAGEM
                                 Dim caminhoImagem As String = Path.Combine(ConfiguracaoApp.CAMINHO_IMAGENS, String.Format("{0}{1}", codigoProduto, extensao))

                                 If File.Exists(caminhoImagem) Then
                                     Try
                                         ' Verificar tamanho do arquivo
                                         Dim fileInfo As New FileInfo(caminhoImagem)
                                         If fileInfo.Length > ConfiguracaoApp.TAMANHO_MAXIMO_IMAGEM Then
                                             LogErros.RegistrarInfo(String.Format("Imagem muito grande: {0} ({1} bytes)", caminhoImagem, fileInfo.Length), "UcReposicaoEstoque.CarregarImagemProdutoAsync")
                                             Continue For
                                         End If

                                         ' Carregar imagem usando FileStream para evitar bloqueio do arquivo
                                         Using fs As New FileStream(caminhoImagem, FileMode.Open, FileAccess.Read)
                                             imagemCarregada = Image.FromStream(fs)
                                         End Using

                                         imagemEncontrada = True
                                         Exit For

                                     Catch imgEx As Exception
                                         LogErros.RegistrarErro(imgEx, String.Format("UcReposicaoEstoque.CarregarImagemProdutoAsync - Erro ao carregar {0}", caminhoImagem))
                                         Continue For
                                     End Try
                                 End If
                             Next

                             ' Atualizar UI no thread principal
                             Me.Invoke(Sub()
                                           If imagemEncontrada AndAlso imagemCarregada IsNot Nothing Then
                                               pbProduto.Image = imagemCarregada
                                           Else
                                               ' Verificar se o diretório existe
                                               If Not Directory.Exists(ConfiguracaoApp.CAMINHO_IMAGENS) Then
                                                   Try
                                                       Directory.CreateDirectory(ConfiguracaoApp.CAMINHO_IMAGENS)
                                                       LogErros.RegistrarInfo(String.Format("Diretório criado: {0}", ConfiguracaoApp.CAMINHO_IMAGENS), "UcReposicaoEstoque.CarregarImagemProdutoAsync")
                                                   Catch dirEx As Exception
                                                       LogErros.RegistrarErro(dirEx, "UcReposicaoEstoque.CarregarImagemProdutoAsync - Erro ao criar diretório")
                                                   End Try
                                               End If
                                           End If

                                           MostrarIndicadoresCarregamento(False)
                                       End Sub)

                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, String.Format("UcReposicaoEstoque.CarregarImagemProdutoAsync({0})", codigoProduto))
                             Me.Invoke(Sub()
                                           pbProduto.Image = Nothing
                                           MostrarIndicadoresCarregamento(False)
                                       End Sub)
                         Finally
                             isCarregandoImagem = False
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, String.Format("UcReposicaoEstoque.CarregarImagemProdutoAsync({0}) - Outer", codigoProduto))
            isCarregandoImagem = False
            MostrarIndicadoresCarregamento(False)
        End Try
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing Then
                ' Parar timers
                If debounceTimer IsNot Nothing Then
                    debounceTimer.Stop()
                    debounceTimer.Dispose()
                End If

                If filtroTimer IsNot Nothing Then
                    filtroTimer.Stop()
                    filtroTimer.Dispose()
                End If

                ' Limpar imagem
                If pbProduto IsNot Nothing AndAlso pbProduto.Image IsNot Nothing Then
                    pbProduto.Image.Dispose()
                End If

                ' Limpar referências
                If dadosProdutosOriginais IsNot Nothing Then
                    dadosProdutosOriginais.Dispose()
                End If

                powerQueryManager = Nothing
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

End Class