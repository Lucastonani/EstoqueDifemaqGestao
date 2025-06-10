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

        ' Adicione esta linha para debug:
        VerificarEstruturaDados()
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
        ' Timer para debounce da seleção de produtos - reduzir delay
        debounceTimer = New Timer()
        debounceTimer.Interval = 150 ' Reduzido de 300ms para 150ms
        AddHandler debounceTimer.Tick, AddressOf DebounceTimer_Tick

        ' Timer para filtro de produtos
        filtroTimer = New Timer()
        filtroTimer.Interval = 300 ' Reduzido de 500ms
        AddHandler filtroTimer.Tick, AddressOf FiltroTimer_Tick
    End Sub

    Private Sub ConfigurarFiltros()
        ' Configuração básica do filtro
        txtFiltro.Text = ""
        txtFiltro.ForeColor = Color.Black
    End Sub

    Private Sub InicializarDados()
        Try
            ' Obter a instância correta do Workbook usando o método seguro
            Dim workbookObj As Microsoft.Office.Interop.Excel.Workbook = Nothing

            Try
                ' Tentar obter através do método público primeiro
                If TypeOf Globals.ThisWorkbook Is ThisWorkbook Then
                    Dim thisWb As ThisWorkbook = CType(Globals.ThisWorkbook, ThisWorkbook)
                    workbookObj = thisWb.ObterWorkbook()
                End If

                ' Se ainda não conseguiu, tentar através da aplicação
                If workbookObj Is Nothing Then
                    Dim excelApp As Microsoft.Office.Interop.Excel.Application = CType(Globals.ThisWorkbook.Application, Microsoft.Office.Interop.Excel.Application)
                    workbookObj = excelApp.ActiveWorkbook
                End If

                ' Último recurso: tentar casting direto
                If workbookObj Is Nothing Then
                    workbookObj = CType(Globals.ThisWorkbook.InnerObject, Microsoft.Office.Interop.Excel.Workbook)
                End If

            Catch castEx As Exception
                LogErros.RegistrarErro(castEx, "UcReposicaoEstoque.InicializarDados - Erro no casting do Workbook")
                MessageBox.Show("Erro ao acessar o workbook do Excel. Verifique se o Excel está funcionando corretamente.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            If workbookObj Is Nothing Then
                MessageBox.Show("Não foi possível acessar o workbook do Excel.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            powerQueryManager = New PowerQueryManager(workbookObj)

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
            If powerQueryManager IsNot Nothing Then
                powerQueryManager.AtualizarTodasConsultas()
                LogErros.RegistrarInfo("Power Query atualizado com sucesso", "UcReposicaoEstoque.AtualizarDadosPowerQuery")
            End If

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
            If powerQueryManager Is Nothing Then
                MessageBox.Show("PowerQueryManager não está inicializado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Mostrar cursor de espera
            Me.Cursor = Cursors.WaitCursor

            Dim tabelaProdutos As ListObject = powerQueryManager.ObterTabela(ConfiguracaoApp.TABELA_PRODUTOS)
            If tabelaProdutos Is Nothing Then
                MessageBox.Show($"Tabela '{ConfiguracaoApp.TABELA_PRODUTOS}' não encontrada!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.Cursor = Cursors.Default
                Return
            End If

            ' Suspender atualizações visuais para melhor performance
            dgvProdutos.SuspendLayout()

            Try
                ' Converter para DataTable e armazenar cópia original
                dadosProdutosOriginais = DataHelper.ConvertListObjectToDataTable(tabelaProdutos)

                ' Aplicar filtro se existir
                AplicarFiltro()

                ' Configurar colunas baseado na estrutura real
                ConfigurarColunasProdutosOtimizado()

                ' Selecionar primeiro produto se existir
                If dgvProdutos.Rows.Count > 0 Then
                    dgvProdutos.Rows(0).Selected = True
                End If

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

    Private Sub CarregarDadosProdutoOtimizado(codigoProduto As String)
        Try
            ' Limpar dados anteriores
            LimparDadosSecundarios()

            ' Suspender layouts para performance
            dgvEstoque.SuspendLayout()
            dgvCompras.SuspendLayout()
            dgvVendas.SuspendLayout()

            Try
                ' Carregar estoque
                Dim estoqueData = CarregarDadosFiltrados(ConfiguracaoApp.TABELA_ESTOQUE, codigoProduto)
                dgvEstoque.DataSource = estoqueData
                ConfigurarColunasEstoqueOtimizado()
                grpEstoque.Text = $"📊 Estoque Atual ({estoqueData.Rows.Count} registros)"

                ' Carregar compras
                Dim comprasData = CarregarDadosFiltrados(ConfiguracaoApp.TABELA_COMPRAS, codigoProduto)
                dgvCompras.DataSource = comprasData
                ConfigurarColunasComprasOtimizado()
                grpCompras.Text = $"📈 Compras ({comprasData.Rows.Count} registros)"

                ' Carregar vendas
                Dim vendasData = CarregarDadosFiltrados(ConfiguracaoApp.TABELA_VENDAS, codigoProduto)
                dgvVendas.DataSource = vendasData
                ConfigurarColunasVendasOtimizado()
                grpVendas.Text = $"📉 Vendas ({vendasData.Rows.Count} registros)"

                ' Carregar imagem de forma segura
                CarregarImagemProdutoSeguro(codigoProduto)

            Finally
                dgvEstoque.ResumeLayout(True)
                dgvCompras.ResumeLayout(True)
                dgvVendas.ResumeLayout(True)
            End Try

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.CarregarDadosProdutoOtimizado")
        End Try
    End Sub

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


    ' 5. Configurar colunas de compras otimizado
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

    ' 6. Configurar colunas de vendas otimizado
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

    ' 7. Carregar imagem de forma segura para evitar exceptions
    Private Sub CarregarImagemProdutoSeguro(codigoProduto As String)
        Try
            ' Limpar imagem anterior de forma segura
            If pbProduto.Image IsNot Nothing Then
                Dim oldImage = pbProduto.Image
                pbProduto.Image = Nothing
                oldImage.Dispose()
            End If

            grpImagem.Text = "🖼️ Imagem do Produto - Carregando..."

            ' Procurar imagem
            Dim imagemEncontrada As Boolean = False

            For Each extensao As String In ConfiguracaoApp.EXTENSOES_IMAGEM
                Dim caminhoImagem As String = Path.Combine(ConfiguracaoApp.CAMINHO_IMAGENS, $"{codigoProduto}{extensao}")

                If File.Exists(caminhoImagem) Then
                    Try
                        ' Verificar tamanho do arquivo
                        Dim fileInfo As New FileInfo(caminhoImagem)
                        If fileInfo.Length > ConfiguracaoApp.TAMANHO_MAXIMO_IMAGEM Then
                            Continue For
                        End If

                        ' Carregar imagem de forma segura
                        Using fs As New FileStream(caminhoImagem, FileMode.Open, FileAccess.Read, FileShare.Read)
                            Using ms As New MemoryStream()
                                fs.CopyTo(ms)
                                ms.Position = 0
                                pbProduto.Image = Image.FromStream(ms)
                            End Using
                        End Using

                        imagemEncontrada = True
                        grpImagem.Text = "🖼️ Imagem do Produto"
                        Exit For

                    Catch imgEx As Exception
                        LogErros.RegistrarErro(imgEx, $"Erro ao carregar imagem: {caminhoImagem}")
                    End Try
                End If
            Next

            If Not imagemEncontrada Then
                grpImagem.Text = "🖼️ Imagem do Produto - Não disponível"
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "CarregarImagemProdutoSeguro")
            grpImagem.Text = "🖼️ Imagem do Produto - Erro"
        End Try
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

                        ' Usar cursor de espera
                        Me.Cursor = Cursors.WaitCursor

                        Try
                            ' Chamar método otimizado
                            CarregarDadosProdutoOtimizado(codigoProduto)
                        Finally
                            Me.Cursor = Cursors.Default
                        End Try
                    End If
                End If
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DebounceTimer_Tick")
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub DgvProdutos_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            If e.RowIndex >= 0 Then
                Dim row = dgvProdutos.Rows(e.RowIndex)

                ' Coletar informações do produto
                Dim info As New System.Text.StringBuilder()

                ' Adicionar todas as colunas visíveis
                For Each col As DataGridViewColumn In dgvProdutos.Columns
                    If col.Visible Then
                        Dim valor = If(row.Cells(col.Index).Value IsNot Nothing, row.Cells(col.Index).Value.ToString(), "")
                        info.AppendLine($"{col.HeaderText}: {valor}")
                    End If
                Next

                ' Adicionar informações de estoque se disponível
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

    Private Sub CarregarDadosProdutoAsync(codigoProduto As String)
        Try
            LogDebug($"Iniciando carregamento para produto: {codigoProduto}")

            ' Limpar dados anteriores
            LimparDadosSecundarios()

            ' Mostrar indicadores de carregamento
            MostrarIndicadoresCarregamento(True)

            ' Verificar se PowerQueryManager está disponível
            If powerQueryManager Is Nothing Then
                LogDebug("ERRO: PowerQueryManager é Nothing!")
                MostrarIndicadoresCarregamento(False)
                Return
            End If

            ' Carregar dados SINCRONAMENTE primeiro para debug
            Try
                LogDebug("Carregando estoque...")
                Dim estoqueData = CarregarDadosFiltrados(ConfiguracaoApp.TABELA_ESTOQUE, codigoProduto)
                dgvEstoque.DataSource = estoqueData
                ConfigurarColunasEstoque()
                grpEstoque.Text = String.Format("📊 Estoque Atual ({0} registros)", estoqueData.Rows.Count)
                LogDebug($"Estoque carregado: {estoqueData.Rows.Count} registros")

                LogDebug("Carregando compras...")
                Dim comprasData = CarregarDadosFiltrados(ConfiguracaoApp.TABELA_COMPRAS, codigoProduto)
                dgvCompras.DataSource = comprasData
                ConfigurarColunasCompras()
                grpCompras.Text = String.Format("📈 Compras ({0} registros)", comprasData.Rows.Count)
                LogDebug($"Compras carregadas: {comprasData.Rows.Count} registros")

                LogDebug("Carregando vendas...")
                Dim vendasData = CarregarDadosFiltrados(ConfiguracaoApp.TABELA_VENDAS, codigoProduto)
                dgvVendas.DataSource = vendasData
                ConfigurarColunasVendas()
                grpVendas.Text = String.Format("📉 Vendas ({0} registros)", vendasData.Rows.Count)
                LogDebug($"Vendas carregadas: {vendasData.Rows.Count} registros")

                ' Carregar imagem
                CarregarImagemProduto(codigoProduto)

                ' Limpar indicadores
                MostrarIndicadoresCarregamento(False)

            Catch ex As Exception
                LogDebug($"ERRO ao carregar dados: {ex.Message}")
                LogErros.RegistrarErro(ex, "CarregarDadosProdutoAsync.Sync")
                MostrarIndicadoresCarregamento(False)
            End Try

        Catch ex As Exception
            LogDebug($"ERRO geral: {ex.Message}")
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.CarregarDadosProdutoAsync")
            MostrarIndicadoresCarregamento(False)
        End Try
    End Sub

    Private Sub CarregarImagemProduto(codigoProduto As String)
        Try
            LogDebug($"Carregando imagem para: {codigoProduto}")

            ' Limpar imagem anterior
            If pbProduto.Image IsNot Nothing Then
                pbProduto.Image.Dispose()
                pbProduto.Image = Nothing
            End If

            ' Procurar imagem
            For Each extensao As String In ConfiguracaoApp.EXTENSOES_IMAGEM
                Dim caminhoImagem As String = Path.Combine(ConfiguracaoApp.CAMINHO_IMAGENS, $"{codigoProduto}{extensao}")

                If File.Exists(caminhoImagem) Then
                    LogDebug($"Imagem encontrada: {caminhoImagem}")
                    Using fs As New FileStream(caminhoImagem, FileMode.Open, FileAccess.Read)
                        pbProduto.Image = Image.FromStream(fs)
                    End Using
                    grpImagem.Text = "🖼️ Imagem do Produto"
                    Return
                End If
            Next

            LogDebug("Nenhuma imagem encontrada")
            grpImagem.Text = "🖼️ Imagem do Produto - Não disponível"

        Catch ex As Exception
            LogDebug($"ERRO ao carregar imagem: {ex.Message}")
            LogErros.RegistrarErro(ex, "CarregarImagemProduto")
            grpImagem.Text = "🖼️ Imagem do Produto - Erro"
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
            LogDebug($"CarregarDadosFiltrados: {nomeTabela} para produto {codigoProduto}")

            If powerQueryManager Is Nothing Then
                LogDebug("ERRO: powerQueryManager é Nothing")
                Return New System.Data.DataTable()
            End If

            Dim tabela As ListObject = powerQueryManager.ObterTabela(nomeTabela)
            If tabela Is Nothing Then
                LogDebug($"AVISO: Tabela {nomeTabela} não encontrada")
                Return New System.Data.DataTable()
            End If

            LogDebug($"Tabela {nomeTabela} encontrada, convertendo...")
            Dim dataTable As System.Data.DataTable = DataHelper.ConvertListObjectToDataTable(tabela)
            LogDebug($"Total de registros na tabela: {dataTable.Rows.Count}")

            ' Debug: mostrar nomes das colunas
            Dim colunas As String = String.Join(", ", dataTable.Columns.Cast(Of DataColumn).Select(Function(c) c.ColumnName))
            LogDebug($"Colunas: {colunas}")

            ' Filtrar por código do produto
            If dataTable.Columns.Count > 0 Then
                Dim nomeColuna As String = dataTable.Columns(0).ColumnName
                LogDebug($"Filtrando pela coluna: {nomeColuna}")

                Dim resultado = DataHelper.FiltrarDataTable(dataTable, nomeColuna, codigoProduto)
                LogDebug($"Registros após filtro: {resultado.Rows.Count}")

                Return resultado
            End If

            Return dataTable

        Catch ex As Exception
            LogDebug($"ERRO em CarregarDadosFiltrados: {ex.Message}")
            LogErros.RegistrarErro(ex, $"UcReposicaoEstoque.CarregarDadosFiltrados({nomeTabela}, {codigoProduto})")
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

                ' Limpar referências
                If dadosProdutosOriginais IsNot Nothing Then
                    dadosProdutosOriginais.Dispose()
                    dadosProdutosOriginais = Nothing
                End If

                powerQueryManager = Nothing
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' Método para debug - adicione temporariamente
    Private Sub LogDebug(mensagem As String)
        Console.WriteLine($"[DEBUG] {DateTime.Now:HH:mm:ss.fff} - {mensagem}")
        LogErros.RegistrarInfo(mensagem, "UcReposicaoEstoque.Debug")
    End Sub

    ' Adicione este método no UcReposicaoEstoque.vb para verificar estrutura das tabelas:

    Private Sub VerificarEstruturaDados()
        Try
            Console.WriteLine("=== VERIFICANDO ESTRUTURA DAS TABELAS ===")

            Dim tabelas As String() = {
                ConfiguracaoApp.TABELA_PRODUTOS,
                ConfiguracaoApp.TABELA_ESTOQUE,
                ConfiguracaoApp.TABELA_COMPRAS,
                ConfiguracaoApp.TABELA_VENDAS
            }

            For Each nomeTabela In tabelas
                Console.WriteLine($"{vbCrLf}Tabela: {nomeTabela}")

                Dim tabela = powerQueryManager.ObterTabela(nomeTabela)
                If tabela IsNot Nothing Then
                    Dim dt = DataHelper.ConvertListObjectToDataTable(tabela)

                    Console.WriteLine($"  Total de registros: {dt.Rows.Count}")
                    Console.WriteLine($"  Colunas ({dt.Columns.Count}):")

                    For i As Integer = 0 To dt.Columns.Count - 1
                        Console.WriteLine($"    [{i}] {dt.Columns(i).ColumnName} ({dt.Columns(i).DataType.Name})")
                    Next

                    ' Mostrar primeira linha como exemplo
                    If dt.Rows.Count > 0 Then
                        Console.WriteLine("  Primeira linha (exemplo):")
                        For i As Integer = 0 To Math.Min(4, dt.Columns.Count - 1)
                            Dim valor = If(dt.Rows(0)(i) IsNot Nothing, dt.Rows(0)(i).ToString(), "(vazio)")
                            Console.WriteLine($"    {dt.Columns(i).ColumnName}: {valor}")
                        Next
                    End If
                Else
                    Console.WriteLine($"  ERRO: Tabela não encontrada!")
                End If
            Next

            Console.WriteLine($"{vbCrLf}=== FIM DA VERIFICAÇÃO ===")

        Catch ex As Exception
            Console.WriteLine($"ERRO na verificação: {ex.Message}")
            LogErros.RegistrarErro(ex, "VerificarEstruturaDados")
        End Try
    End Sub

    Private Sub ConfigurarEventos()
        Try
            ' Eventos do DataGridView de Produtos
            AddHandler dgvProdutos.SelectionChanged, AddressOf DgvProdutos_SelectionChanged
            AddHandler dgvProdutos.CellDoubleClick, AddressOf DgvProdutos_CellDoubleClick
            AddHandler dgvProdutos.DataBindingComplete, AddressOf DgvProdutos_DataBindingComplete

            ' Eventos do filtro
            AddHandler txtFiltro.TextChanged, AddressOf TxtFiltro_TextChanged
            AddHandler txtFiltro.Enter, AddressOf TxtFiltro_Enter
            AddHandler txtFiltro.Leave, AddressOf TxtFiltro_Leave

            ' Evento do botão atualizar
            AddHandler btnAtualizar.Click, AddressOf BtnAtualizar_Click

            ' Eventos adicionais para melhor UX
            AddHandler dgvProdutos.KeyDown, AddressOf DgvProdutos_KeyDown

            LogErros.RegistrarInfo("Eventos configurados", "UcReposicaoEstoque.ConfigurarEventos")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarEventos")
        End Try
    End Sub

    ' Adicione também estes manipuladores de eventos que podem estar faltando:

    Private Sub DgvProdutos_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        Try
            ' Ajustar larguras das colunas após binding
            If dgvProdutos.Columns.Count > 0 Then
                dgvProdutos.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells)
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DgvProdutos_DataBindingComplete")
        End Try
    End Sub

    Private Sub TxtFiltro_Enter(sender As Object, e As EventArgs)
        Try
            ' Selecionar todo o texto ao entrar no campo
            txtFiltro.SelectAll()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.TxtFiltro_Enter")
        End Try
    End Sub

    Private Sub TxtFiltro_Leave(sender As Object, e As EventArgs)
        Try
            ' Pode adicionar lógica adicional se necessário
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.TxtFiltro_Leave")
        End Try
    End Sub

    Private Sub DgvProdutos_KeyDown(sender As Object, e As KeyEventArgs)
        Try
            ' Permitir navegação com teclado
            Select Case e.KeyCode
                Case Keys.Enter
                    ' Enter funciona como double-click
                    If dgvProdutos.SelectedRows.Count > 0 Then
                        DgvProdutos_CellDoubleClick(sender, New DataGridViewCellEventArgs(0, dgvProdutos.SelectedRows(0).Index))
                    End If
                    e.Handled = True

                Case Keys.F5
                    ' F5 para atualizar
                    btnAtualizar.PerformClick()
                    e.Handled = True

                Case Keys.F, Keys.F3
                    If e.Control OrElse e.KeyCode = Keys.F3 Then
                        ' Ctrl+F ou F3 para focar no filtro
                        txtFiltro.Focus()
                        e.Handled = True
                    End If
            End Select

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DgvProdutos_KeyDown")
        End Try
    End Sub

End Class