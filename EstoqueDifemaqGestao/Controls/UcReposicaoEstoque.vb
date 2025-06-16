Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
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
    Private Shared tabelasEstaticas As New Dictionary(Of String, System.Data.DataTable)
    Private Shared ultimaAtualizacaoEstatica As DateTime = DateTime.MinValue
    Private colunasConfiguradas As Boolean = False
    Private ultimaLimpezaCache As DateTime = DateTime.MinValue

    ' ✅ CACHE DE IMAGENS - Declarações necessárias
    Private Shared cacheImagens As New Dictionary(Of String, Image)
    Private Shared cacheStatusImagens As New Dictionary(Of String, String)
    Private Shared ultimaLimpezaImagensCache As DateTime = DateTime.MinValue
    Private Const CACHE_IMAGENS_TIMEOUT_MINUTES As Integer = 30

    ' APIs do Windows para otimização de redesenho
    <DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As Boolean, lParam As Integer) As Integer
    End Function
    Private Const WM_SETREDRAW As Integer = 11

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

    ' ✅ MÉTODO PÚBLICO: Permitir inicialização externa
    Public Sub InicializarDadosSeNecessario()
        Try
            If Not dadosCarregados Then
                InicializarDados()
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.InicializarDadosSeNecessario")
        End Try
    End Sub

    ' Lazy loading - só carrega dados quando necessário
    Private Sub InicializarDados()
        Try
            If dadosCarregados Then
                RestaurarBotaoAtualizar()
                Return
            End If

            ' Obter workbook de forma otimizada
            Dim workbookObj = ObterWorkbookOtimizado()
            If workbookObj Is Nothing Then
                MessageBox.Show("Não foi possível acessar o workbook do Excel.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                RestaurarBotaoAtualizar()
                Return
            End If

            powerQueryManager = New PowerQueryManager(workbookObj)
            Me.Cursor = Cursors.WaitCursor
            AtualizarStatus("Inicializando dados pela primeira vez...")

            ' Carregar produtos em background
            Task.Run(Sub()
                         Try
                             Me.Invoke(Sub()
                                           CarregarProdutos()
                                           dadosCarregados = True
                                           AtualizarStatus("Dados inicializados com sucesso!")
                                       End Sub)
                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, "UcReposicaoEstoque.InicializarDados.Background")
                             Me.Invoke(Sub()
                                           AtualizarStatus("Erro ao inicializar dados")
                                           MessageBox.Show($"Erro ao carregar dados: {ex.Message}", "Erro",
                                                         MessageBoxButtons.OK, MessageBoxIcon.Error)
                                       End Sub)
                         Finally
                             Me.Invoke(Sub()
                                           RestaurarBotaoAtualizar()
                                       End Sub)
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.InicializarDados")
            MessageBox.Show($"Erro ao inicializar dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            RestaurarBotaoAtualizar()
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
            ' Verificar se já está atualizando
            If Not btnAtualizar.Enabled Then
                Return
            End If

            btnAtualizar.Enabled = False
            btnAtualizar.Text = "🔄 Atualizando..."
            Me.Cursor = Cursors.WaitCursor
            AtualizarStatus("Atualizando dados...")

            ' Invalidar cache completo
            InvalidarCacheCompleto()

            ' Verificar se dados estão inicializados
            If Not dadosCarregados Then
                RestaurarBotaoAtualizar()
                InicializarDados()
                Return
            End If

            Task.Run(Sub()
                         Try
                             LogErros.RegistrarInfo("Iniciando atualização Power Query", "BtnAtualizar_Click")
                             AtualizarDadosPowerQuery()

                             Me.Invoke(Sub()
                                           Try
                                               LogErros.RegistrarInfo("Recarregando produtos após Power Query", "BtnAtualizar_Click")
                                               CarregarProdutos()
                                               AtualizarStatus("Dados atualizados com sucesso!")

                                               ' Mostrar mensagem de sucesso apenas se não houve erro
                                               MessageBox.Show("Dados atualizados com sucesso!", "Sucesso",
                                                         MessageBoxButtons.OK, MessageBoxIcon.Information)
                                           Catch loadEx As Exception
                                               LogErros.RegistrarErro(loadEx, "BtnAtualizar_Click.CarregarProdutos")
                                               AtualizarStatus("Erro ao carregar produtos após atualização")
                                               MessageBox.Show($"Erro ao carregar produtos: {loadEx.Message}", "Erro",
                                                         MessageBoxButtons.OK, MessageBoxIcon.Error)
                                           End Try
                                       End Sub)

                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, "UcReposicaoEstoque.BtnAtualizar_Click.Background")
                             Me.Invoke(Sub()
                                           AtualizarStatus("Erro na atualização dos dados")
                                           MessageBox.Show($"Erro ao atualizar dados: {ex.Message}", "Erro",
                                                     MessageBoxButtons.OK, MessageBoxIcon.Error)
                                       End Sub)
                         Finally
                             ' SEMPRE restaurar o botão, independente de sucesso ou erro
                             Me.Invoke(Sub()
                                           RestaurarBotaoAtualizar()
                                       End Sub)
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.BtnAtualizar_Click")
            AtualizarStatus("Erro na atualização")
            MessageBox.Show($"Erro ao atualizar dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            RestaurarBotaoAtualizar()
        End Try
    End Sub

    ' Método separado para garantir que o botão seja sempre restaurado
    Private Sub RestaurarBotaoAtualizar()
        Try
            If btnAtualizar IsNot Nothing AndAlso Not btnAtualizar.IsDisposed Then
                btnAtualizar.Enabled = True
                btnAtualizar.Text = "🔄 Atualizar"
            End If

            Me.Cursor = Cursors.Default

            If String.IsNullOrEmpty(filtroAtual) Then
                AtualizarStatus("Pronto")
            Else
                AtualizarStatus($"Pronto - Filtro ativo: '{filtroAtual}'")
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.RestaurarBotaoAtualizar")
        End Try
    End Sub

    ' Método público para forçar restauração do botão (fallback)
    Public Sub ForcarRestauracaoBotao()
        Try
            RestaurarBotaoAtualizar()
            LogErros.RegistrarInfo("Restauração forçada do botão atualizar", "ForcarRestauracaoBotao")
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ForcarRestauracaoBotao")
        End Try
    End Sub

    ' Verificação periódica do estado do botão (executada a cada 5 segundos)
    Private Sub VerificarEstadoBotao()
        Try
            ' Se o botão estiver desabilitado por mais de 2 minutos, restaurar
            Static ultimaVerificacao As DateTime = DateTime.MinValue
            Static botaoDesabilitadoDesde As DateTime = DateTime.MinValue

            If DateTime.Now.Subtract(ultimaVerificacao).TotalSeconds < 5 Then
                Return ' Verificar apenas a cada 5 segundos
            End If

            ultimaVerificacao = DateTime.Now

            If btnAtualizar IsNot Nothing AndAlso Not btnAtualizar.IsDisposed Then
                If Not btnAtualizar.Enabled Then
                    If botaoDesabilitadoDesde = DateTime.MinValue Then
                        botaoDesabilitadoDesde = DateTime.Now
                    ElseIf DateTime.Now.Subtract(botaoDesabilitadoDesde).TotalMinutes > 2 Then
                        ' Botão desabilitado por mais de 2 minutos - forçar restauração
                        LogErros.RegistrarInfo("Botão desabilitado por mais de 2 minutos - forçando restauração", "VerificarEstadoBotao")
                        RestaurarBotaoAtualizar()
                        botaoDesabilitadoDesde = DateTime.MinValue
                    End If
                Else
                    botaoDesabilitadoDesde = DateTime.MinValue
                End If
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.VerificarEstadoBotao")
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

    ' CARREGAMENTO OTIMIZADO: Uma consulta, filtro rápido
    Private Function CarregarDadosOtimizado(codigoProduto As String) As Dictionary(Of String, System.Data.DataTable)
        Dim resultado As New Dictionary(Of String, System.Data.DataTable)

        Try
            ' USAR cache estático das tabelas completas (recarrega apenas a cada 5 min)
            If Not TabelasEstaticasValidas() Then
                CarregarTabelasEstaticas()
            End If

            ' FILTRAR usando DataView (100x mais rápido que loop manual)
            resultado("estoque") = FiltrarRapidoComDataView(tabelasEstaticas("estoque"), codigoProduto)
            resultado("compras") = FiltrarRapidoComDataView(tabelasEstaticas("compras"), codigoProduto)
            resultado("vendas") = FiltrarRapidoComDataView(tabelasEstaticas("vendas"), codigoProduto)

            ' ARMAZENAR no cache individual
            ArmazenarNoCacheIndividual(codigoProduto, resultado)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "CarregarDadosOtimizado")
            ' Retornar DataTables vazios em caso de erro
            resultado("estoque") = CriarDataTableVazio()
            resultado("compras") = CriarDataTableVazio()
            resultado("vendas") = CriarDataTableVazio()
        End Try

        Return resultado
    End Function

    ' CACHE ESTÁTICO: Carrega tabelas completas apenas uma vez
    Private Function TabelasEstaticasValidas() As Boolean
        Return DateTime.Now.Subtract(ultimaAtualizacaoEstatica).TotalMinutes < 5 AndAlso
           tabelasEstaticas.ContainsKey("estoque") AndAlso
           tabelasEstaticas.ContainsKey("compras") AndAlso
           tabelasEstaticas.ContainsKey("vendas")
    End Function

    Private Sub CarregarTabelasEstaticas()
        Try
            LogErros.RegistrarInfo("Carregando tabelas estáticas...", "CarregarTabelasEstaticas")

            ' Obter referência ao Excel de forma correta
            Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
            Dim workbookInterface = ObterWorkbookOtimizado()

            If workbookInterface IsNot Nothing Then
                excelApp = workbookInterface.Application
            End If

            ' Desabilitar cálculo automático do Excel
            Dim calculationState As Microsoft.Office.Interop.Excel.XlCalculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic
            If excelApp IsNot Nothing Then
                calculationState = excelApp.Calculation
                excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual
            End If

            Try
                ' Carregar tabelas completas UMA VEZ
                Dim tabelaEstoque = powerQueryManager.ObterTabela(ConfiguracaoApp.TABELA_ESTOQUE)
                Dim tabelaCompras = powerQueryManager.ObterTabela(ConfiguracaoApp.TABELA_COMPRAS)
                Dim tabelaVendas = powerQueryManager.ObterTabela(ConfiguracaoApp.TABELA_VENDAS)

                ' Converter para DataTable e armazenar
                tabelasEstaticas("estoque") = If(tabelaEstoque IsNot Nothing,
                    DataHelper.ConvertListObjectToDataTable(tabelaEstoque), CriarDataTableVazio())
                tabelasEstaticas("compras") = If(tabelaCompras IsNot Nothing,
                    DataHelper.ConvertListObjectToDataTable(tabelaCompras), CriarDataTableVazio())
                tabelasEstaticas("vendas") = If(tabelaVendas IsNot Nothing,
                    DataHelper.ConvertListObjectToDataTable(tabelaVendas), CriarDataTableVazio())

                ultimaAtualizacaoEstatica = DateTime.Now

                LogErros.RegistrarInfo($"Tabelas carregadas - Estoque: {tabelasEstaticas("estoque").Rows.Count}, " &
                                 $"Compras: {tabelasEstaticas("compras").Rows.Count}, " &
                                 $"Vendas: {tabelasEstaticas("vendas").Rows.Count}", "CarregarTabelasEstaticas")
            Finally
                If excelApp IsNot Nothing Then
                    excelApp.Calculation = calculationState
                End If
            End Try

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "CarregarTabelasEstaticas")
            ultimaAtualizacaoEstatica = DateTime.MinValue ' Invalidar cache
        End Try
    End Sub

    ' Helper para criar DataTable vazio
    Private Function CriarDataTableVazio() As System.Data.DataTable
        Return New System.Data.DataTable()
    End Function

    ' FILTRO ULTRA-RÁPIDO: DataView em vez de loop manual
    Private Function FiltrarRapidoComDataView(tabelaCompleta As System.Data.DataTable, codigoProduto As String) As System.Data.DataTable
        Try
            If tabelaCompleta Is Nothing OrElse tabelaCompleta.Rows.Count = 0 Then
                Return CriarDataTableVazio()
            End If

            ' Criar DataView com filtro (muito mais rápido que loop)
            Dim dataView As New DataView(tabelaCompleta)
            Dim nomeColuna = tabelaCompleta.Columns(0).ColumnName

            ' Escapar caracteres especiais para RowFilter
            Dim codigoEscapado = codigoProduto.Replace("'", "''").Replace("[", "\[").Replace("]", "\]")
            dataView.RowFilter = $"[{nomeColuna}] = '{codigoEscapado}'"

            ' Converter para DataTable
            Return dataView.ToTable()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"FiltrarRapidoComDataView - Produto: {codigoProduto}")
            Return CriarDataTableVazio()
        End Try
    End Function

    ' APLICAÇÃO ULTRA-RÁPIDA: Sem estilização repetitiva
    Private Sub AplicarDadosUltraRapido(dados As Dictionary(Of String, System.Data.DataTable), codigoProduto As String)
        Try
            ' Aplicar DataSource diretamente
            dgvEstoque.DataSource = dados("estoque")
            dgvCompras.DataSource = dados("compras")
            dgvVendas.DataSource = dados("vendas")

            ' Configurar colunas apenas UMA VEZ
            If Not colunasConfiguradas Then
                ConfigurarTodasAsColunasUmaVez()
                colunasConfiguradas = True
            End If

            ' Atualizar contadores nos GroupBox
            grpEstoque.Text = $"📊 Estoque Atual ({dados("estoque").Rows.Count} registros)"
            grpCompras.Text = $"📈 Compras ({dados("compras").Rows.Count} registros)"
            grpVendas.Text = $"📉 Vendas ({dados("vendas").Rows.Count} registros)"

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AplicarDadosUltraRapido")
        End Try
    End Sub

    ' CONFIGURAÇÃO UMA VEZ: Evita reconfigurar colunas
    Private Sub ConfigurarTodasAsColunasUmaVez()
        Try
            ConfigurarColunasEstoqueOtimizado()
            ConfigurarColunasComprasOtimizado()
            ConfigurarColunasVendasOtimizado()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ConfigurarTodasAsColunasUmaVez")
        End Try
    End Sub

    ' CACHE INDIVIDUAL: Para produtos já consultados
    Private Function VerificarCacheIndividual(codigoProduto As String) As Boolean
        Return CacheEstaValido() AndAlso
           cacheEstoque.ContainsKey($"estoque_{codigoProduto}") AndAlso
           cacheCompras.ContainsKey($"compras_{codigoProduto}") AndAlso
           cacheVendas.ContainsKey($"vendas_{codigoProduto}")
    End Function

    Private Sub AplicarDadosDoCache(codigoProduto As String)
        Try
            PararRedesenhoCompleto()

            Try
                dgvEstoque.DataSource = cacheEstoque($"estoque_{codigoProduto}")
                dgvCompras.DataSource = cacheCompras($"compras_{codigoProduto}")
                dgvVendas.DataSource = cacheVendas($"vendas_{codigoProduto}")

                ' Atualizar contadores
                grpEstoque.Text = $"📊 Estoque Atual ({dgvEstoque.Rows.Count} registros)"
                grpCompras.Text = $"📈 Compras ({dgvCompras.Rows.Count} registros)"
                grpVendas.Text = $"📉 Vendas ({dgvVendas.Rows.Count} registros)"
            Finally
                ReabilitarRedesenhoCompleto()
            End Try

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AplicarDadosDoCache")
        End Try
    End Sub

    Private Sub ArmazenarNoCacheIndividual(codigoProduto As String, dados As Dictionary(Of String, System.Data.DataTable))
        Try
            cacheEstoque($"estoque_{codigoProduto}") = dados("estoque")
            cacheCompras($"compras_{codigoProduto}") = dados("compras")
            cacheVendas($"vendas_{codigoProduto}") = dados("vendas")
            cacheValido = DateTime.Now
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ArmazenarNoCacheIndividual")
        End Try
    End Sub

    ' IMAGEM ASSÍNCRONA: Não trava a UI
    Private Sub CarregarImagemAsync(codigoProduto As String)
        Task.Run(Sub()
                     Try
                         CarregarImagemProdutoAsync(codigoProduto)
                     Catch ex As Exception
                         LogErros.RegistrarErro(ex, "CarregarImagemAsync")
                     End Try
                 End Sub)
    End Sub

    ' HANDLER para DataBindingComplete (se necessário)
    Private Sub DataGrid_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        ' Configurações pós-binding se necessário
    End Sub

    Private Sub CarregarDadosProdutoUltraRapido(codigoProduto As String)
        Try
            Dim sw As New Stopwatch()
            sw.Start()

            LogErros.RegistrarInfo($"Iniciando carregamento para produto: {codigoProduto}", "CarregarDadosProdutoUltraRapido")

            ' 1. VERIFICAR cache individual primeiro
            If VerificarCacheIndividual(codigoProduto) Then
                AplicarDadosDoCache(codigoProduto)
                sw.Stop()
                LogErros.RegistrarInfo($"✅ Cache hit - Total: {sw.ElapsedMilliseconds}ms", "CarregarDadosProdutoUltraRapido")
                Return
            End If

            ' 2. PARAR COMPLETAMENTE o redesenho (crítico para performance!)
            PararRedesenhoCompleto()

            Try
                ' 3. CARREGAR dados de forma otimizada
                sw.Restart()
                Dim dadosFiltrados = CarregarDadosOtimizado(codigoProduto)
                LogErros.RegistrarInfo($"⚡ Dados obtidos: {sw.ElapsedMilliseconds}ms", "CarregarDados")

                ' 4. APLICAR aos grids rapidamente
                sw.Restart()
                AplicarDadosUltraRapido(dadosFiltrados, codigoProduto)
                LogErros.RegistrarInfo($"📊 Grids atualizados: {sw.ElapsedMilliseconds}ms", "AplicarDados")

            Finally
                ' 5. SEMPRE reabilitar redesenho
                ReabilitarRedesenhoCompleto()
            End Try

            ' 6. IMAGEM em background (não trava UI)
            CarregarImagemAsync(codigoProduto)

            sw.Stop()
            LogErros.RegistrarInfo($"🎯 Total concluído: {sw.ElapsedMilliseconds}ms", "CarregarDadosProdutoUltraRapido")

        Catch ex As Exception
            ReabilitarRedesenhoCompleto()
            LogErros.RegistrarErro(ex, "CarregarDadosProdutoUltraRapido")
            MessageBox.Show($"Erro ao carregar dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ' OTIMIZAÇÃO CRÍTICA: Parar redesenho COMPLETAMENTE
    Private Sub PararRedesenhoCompleto()
        Try
            ' API Windows - PARA completamente o redesenho
            SendMessage(dgvEstoque.Handle, WM_SETREDRAW, False, 0)
            SendMessage(dgvCompras.Handle, WM_SETREDRAW, False, 0)
            SendMessage(dgvVendas.Handle, WM_SETREDRAW, False, 0)

            ' Suspender layouts
            dgvEstoque.SuspendLayout()
            dgvCompras.SuspendLayout()
            dgvVendas.SuspendLayout()

            ' Desabilitar auto-sizing (MUITO lento!)
            dgvEstoque.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvCompras.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvVendas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            ' Desabilitar eventos temporariamente
            RemoveHandler dgvEstoque.DataBindingComplete, AddressOf DataGrid_DataBindingComplete
            RemoveHandler dgvCompras.DataBindingComplete, AddressOf DataGrid_DataBindingComplete
            RemoveHandler dgvVendas.DataBindingComplete, AddressOf DataGrid_DataBindingComplete

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PararRedesenhoCompleto")
        End Try
    End Sub

    Private Sub ReabilitarRedesenhoCompleto()
        Try
            ' Reabilitar eventos
            AddHandler dgvEstoque.DataBindingComplete, AddressOf DataGrid_DataBindingComplete
            AddHandler dgvCompras.DataBindingComplete, AddressOf DataGrid_DataBindingComplete
            AddHandler dgvVendas.DataBindingComplete, AddressOf DataGrid_DataBindingComplete

            ' Resumir layouts
            dgvEstoque.ResumeLayout(False) ' False = não força refresh ainda
            dgvCompras.ResumeLayout(False)
            dgvVendas.ResumeLayout(False)

            ' Reabilitar auto-sizing APENAS se necessário
            If Not colunasConfiguradas Then
                dgvEstoque.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
                dgvCompras.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
                dgvVendas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
            End If

            ' Reabilitar redesenho
            SendMessage(dgvEstoque.Handle, WM_SETREDRAW, True, 0)
            SendMessage(dgvCompras.Handle, WM_SETREDRAW, True, 0)
            SendMessage(dgvVendas.Handle, WM_SETREDRAW, True, 0)

            ' Forçar atualização visual
            dgvEstoque.Invalidate()
            dgvCompras.Invalidate()
            dgvVendas.Invalidate()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ReabilitarRedesenhoCompleto")
        End Try
    End Sub

    ' ✅ CARREGAMENTO DE IMAGEM OTIMIZADO COM CACHE COMPLETO (VERSÃO FINAL)
    Private Sub CarregarImagemProdutoAsync(codigoProduto As String)
        If isCarregandoImagem Then Return

        Try
            LogErros.RegistrarInfo($"🔍 Iniciando carregamento de imagem para: {codigoProduto}", "CarregarImagem")

            ' ✅ VERIFICAR CACHE DE IMAGEM PRIMEIRO
            If cacheImagens.ContainsKey(codigoProduto) Then
                LogErros.RegistrarInfo($"📦 Imagem encontrada no cache para: {codigoProduto}", "CarregarImagem")
                ' Imagem já está no cache - aplicar imediatamente
                Try
                    If Me.InvokeRequired Then
                        Me.Invoke(Sub() AplicarImagemDoCache(codigoProduto))
                    Else
                        AplicarImagemDoCache(codigoProduto)
                    End If
                    Return
                Catch cacheEx As Exception
                    LogErros.RegistrarErro(cacheEx, "CarregarImagemProdutoAsync - Cache")
                    ' Continue para carregamento normal se cache falhar
                End Try
            Else
                LogErros.RegistrarInfo($"❌ Imagem NÃO encontrada no cache para: {codigoProduto}", "CarregarImagem")
            End If

            ' ✅ VERIFICAR STATUS CACHE - se já tentou carregar antes
            If cacheStatusImagens.ContainsKey(codigoProduto) Then
                Dim status = cacheStatusImagens(codigoProduto)
                LogErros.RegistrarInfo($"📊 Status cache para {codigoProduto}: {status}", "CarregarImagem")

                If status = "NAO_ENCONTRADA" Then
                    ' Já tentou e não encontrou - não tentar novamente
                    AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Não disponível", Nothing)
                    LogErros.RegistrarInfo($"⚠️ Status cache: imagem não disponível para {codigoProduto}", "CarregarImagem")
                    Return
                ElseIf status = "ERRO" Then
                    ' Já tentou e deu erro - não tentar novamente
                    AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Erro", Nothing)
                    LogErros.RegistrarInfo($"❌ Status cache: erro anterior para {codigoProduto}", "CarregarImagem")
                    Return
                End If
            End If

            isCarregandoImagem = True

            ' ✅ Atualizar UI no thread principal
            AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Carregando...", Nothing)

            Task.Run(Sub()
                         Try
                             Dim imagemEncontrada As Boolean = False
                             Dim imagemCarregada As Image = Nothing
                             Dim caminhoEncontrado As String = ""
                             Dim errosDetalhes As New List(Of String)

                             LogErros.RegistrarInfo($"🔄 Procurando arquivos de imagem para: {codigoProduto}", "CarregarImagem")

                             ' Procurar imagem com diferentes extensões
                             For Each extensao As String In ConfiguracaoApp.EXTENSOES_IMAGEM
                                 Dim caminhoImagem = Path.Combine(ConfiguracaoApp.CAMINHO_IMAGENS, $"{codigoProduto}{extensao}")

                                 If File.Exists(caminhoImagem) Then
                                     LogErros.RegistrarInfo($"📁 Arquivo encontrado: {caminhoImagem}", "CarregarImagem")
                                     Try
                                         Dim fileInfo As New FileInfo(caminhoImagem)

                                         ' Validações de tamanho
                                         If fileInfo.Length = 0 Then
                                             errosDetalhes.Add($"{extensao}: arquivo vazio")
                                             Continue For
                                         End If

                                         If fileInfo.Length < 100 Then
                                             errosDetalhes.Add($"{extensao}: arquivo muito pequeno ({fileInfo.Length} bytes)")
                                             Continue For
                                         End If

                                         If fileInfo.Length > ConfiguracaoApp.TAMANHO_MAXIMO_IMAGEM Then
                                             errosDetalhes.Add($"{extensao}: arquivo muito grande ({fileInfo.Length / 1024:F0}KB)")
                                             Continue For
                                         End If

                                         ' Tentar carregar imagem
                                         imagemCarregada = TentarCarregarImagem(caminhoImagem)

                                         If imagemCarregada IsNot Nothing Then
                                             caminhoEncontrado = caminhoImagem
                                             imagemEncontrada = True
                                             LogErros.RegistrarInfo($"✅ Imagem carregada com sucesso: {caminhoImagem} ({fileInfo.Length / 1024:F0}KB)", "CarregarImagem")
                                             Exit For
                                         Else
                                             errosDetalhes.Add($"{extensao}: formato inválido")
                                         End If

                                     Catch imgEx As Exception
                                         errosDetalhes.Add($"{extensao}: {imgEx.Message}")
                                         LogErros.RegistrarErro(imgEx, $"Erro ao carregar {caminhoImagem}")
                                         Continue For
                                     End Try
                                 Else
                                     errosDetalhes.Add($"{extensao}: arquivo não existe")
                                 End If
                             Next

                             ' ✅ ARMAZENAR NO CACHE
                             If imagemEncontrada Then
                                 LogErros.RegistrarInfo($"💾 Armazenando no cache: {codigoProduto}", "CarregarImagem")
                                 cacheImagens(codigoProduto) = imagemCarregada
                                 cacheStatusImagens(codigoProduto) = "ENCONTRADA"
                                 AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto", imagemCarregada)
                                 LogErros.RegistrarInfo($"📸 Imagem aplicada com sucesso: {caminhoEncontrado}", "CarregarImagem")
                             Else
                                 LogErros.RegistrarInfo($"💾 Armazenando status 'não encontrada' no cache: {codigoProduto}", "CarregarImagem")
                                 cacheImagens(codigoProduto) = Nothing
                                 cacheStatusImagens(codigoProduto) = "NAO_ENCONTRADA"
                                 AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Não disponível", Nothing)
                                 LogErros.RegistrarInfo($"❌ Nenhuma imagem válida para {codigoProduto}: {String.Join("; ", errosDetalhes)}", "CarregarImagem")
                             End If

                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, $"UcReposicaoEstoque.CarregarImagemProdutoAsync({codigoProduto})")
                             cacheStatusImagens(codigoProduto) = "ERRO"
                             AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Erro", Nothing)
                         Finally
                             isCarregandoImagem = False
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"UcReposicaoEstoque.CarregarImagemProdutoAsync({codigoProduto}) - Outer")
            isCarregandoImagem = False
            cacheStatusImagens(codigoProduto) = "ERRO"
            AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Erro", Nothing)
        End Try
    End Sub

    ' ✅ MÉTODO AUXILIAR: Aplicar imagem do cache (CORRIGIDO - VERSÃO FINAL)
    Private Sub AplicarImagemDoCache(codigoProduto As String)
        Try
            If pbProduto.IsDisposed Then Return

            Dim imagemCache = cacheImagens(codigoProduto)
            If imagemCache IsNot Nothing Then
                ' ✅ CORREÇÃO FINAL: Sempre aplicar a imagem do cache
                ' Limpar imagem atual primeiro (sem dispose - está no cache)
                If pbProduto.Image IsNot Nothing Then
                    pbProduto.Image = Nothing
                End If

                ' Aplicar imagem do cache diretamente
                pbProduto.Image = imagemCache
                grpImagem.Text = "🖼️ Imagem do Produto"
                LogErros.RegistrarInfo($"📸 Imagem aplicada do cache: {codigoProduto}", "CarregarImagem")
            Else
                ' Limpar imagem se não há imagem no cache
                If pbProduto.Image IsNot Nothing Then
                    pbProduto.Image = Nothing
                End If
                grpImagem.Text = "🖼️ Imagem do Produto - Não disponível"
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AplicarImagemDoCache")
            grpImagem.Text = "🖼️ Imagem do Produto - Erro"
        End Try
    End Sub

    ' ✅ MÉTODO AUXILIAR: Aplicar status da imagem (thread-safe)
    Private Sub AplicarImagemStatus(codigoProduto As String, statusTexto As String, imagem As Image)
        Try
            If Me.InvokeRequired Then
                Me.Invoke(Sub() AplicarImagemStatusSeguro(codigoProduto, statusTexto, imagem))
            Else
                AplicarImagemStatusSeguro(codigoProduto, statusTexto, imagem)
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AplicarImagemStatus")
        End Try
    End Sub

    ' ✅ MÉTODO AUXILIAR: Aplicar status da imagem (CORRIGIDO - VERSÃO FINAL)
    Private Sub AplicarImagemStatusSeguro(codigoProduto As String, statusTexto As String, imagem As Image)
        Try
            If Me.IsDisposed OrElse pbProduto.IsDisposed Then Return

            ' ✅ CORREÇÃO FINAL: Sempre aplicar a nova imagem
            ' Limpar imagem atual (sem dispose - pode estar no cache)
            If pbProduto.Image IsNot Nothing Then
                pbProduto.Image = Nothing
            End If

            ' Aplicar nova imagem
            If imagem IsNot Nothing Then
                pbProduto.Image = imagem
            End If

            If Not grpImagem.IsDisposed Then
                grpImagem.Text = statusTexto
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AplicarImagemStatusSeguro")
            Try
                If Not grpImagem.IsDisposed Then
                    grpImagem.Text = "🖼️ Imagem do Produto - Erro"
                End If
            Catch
                ' Ignorar se não conseguir atualizar UI
            End Try
        End Try
    End Sub

    ' ✅ MÉTODO ROBUSTO: Tentativa de carregamento de imagem
    Private Function TentarCarregarImagem(caminhoArquivo As String) As Image
        Try
            ' MÉTODO 1: Carregamento direto (mais simples)
            Try
                Using fs As New FileStream(caminhoArquivo, FileMode.Open, FileAccess.Read, FileShare.Read)
                    ' Verificar se consegue ler pelo menos alguns bytes
                    Dim buffer(10) As Byte
                    If fs.Read(buffer, 0, 10) < 10 Then
                        LogErros.RegistrarInfo($"❌ Arquivo muito pequeno para ser imagem: {caminhoArquivo}", "TentarCarregarImagem")
                        Return Nothing
                    End If

                    ' Voltar ao início e tentar carregar
                    fs.Seek(0, SeekOrigin.Begin)
                    Return Image.FromStream(fs)
                End Using
            Catch ex As ArgumentException
                LogErros.RegistrarInfo($"❌ Método 1 falhou (formato inválido): {caminhoArquivo} - {ex.Message}", "TentarCarregarImagem")
            End Try

            ' MÉTODO 2: Via array de bytes (mais robusto)
            Try
                Dim bytes() As Byte = File.ReadAllBytes(caminhoArquivo)

                ' Validar header básico da imagem
                If Not ValidarHeaderImagem(bytes) Then
                    LogErros.RegistrarInfo($"❌ Header de imagem inválido: {caminhoArquivo}", "TentarCarregarImagem")
                    Return Nothing
                End If

                Using ms As New MemoryStream(bytes)
                    Return Image.FromStream(ms)
                End Using
            Catch ex As ArgumentException
                LogErros.RegistrarInfo($"❌ Método 2 falhou (formato inválido): {caminhoArquivo} - {ex.Message}", "TentarCarregarImagem")
            End Try

            LogErros.RegistrarInfo($"❌ Todos os métodos falharam: {caminhoArquivo}", "TentarCarregarImagem")
            Return Nothing

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"TentarCarregarImagem({caminhoArquivo})")
            Return Nothing
        End Try
    End Function

    ' ✅ VALIDADOR DE HEADER: Verificar formato básico da imagem
    Private Function ValidarHeaderImagem(bytes() As Byte) As Boolean
        Try
            If bytes Is Nothing OrElse bytes.Length < 10 Then Return False

            ' Verificar headers conhecidos
            ' JPEG: FF D8 FF
            If bytes.Length >= 3 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HD8 AndAlso bytes(2) = &HFF Then
                Return True
            End If

            ' PNG: 89 50 4E 47 0D 0A 1A 0A
            If bytes.Length >= 8 AndAlso bytes(0) = &H89 AndAlso bytes(1) = &H50 AndAlso bytes(2) = &H4E AndAlso bytes(3) = &H47 Then
                Return True
            End If

            ' BMP: 42 4D
            If bytes.Length >= 2 AndAlso bytes(0) = &H42 AndAlso bytes(1) = &H4D Then
                Return True
            End If

            ' GIF: 47 49 46 38
            If bytes.Length >= 4 AndAlso bytes(0) = &H47 AndAlso bytes(1) = &H49 AndAlso bytes(2) = &H46 AndAlso bytes(3) = &H38 Then
                Return True
            End If

            LogErros.RegistrarInfo($"❌ Header desconhecido: {bytes(0):X2} {bytes(1):X2} {bytes(2):X2} {bytes(3):X2}", "ValidarHeaderImagem")
            Return False

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ValidarHeaderImagem")
            Return False
        End Try
    End Function

    ' ✅ LIMPEZA AUTOMÁTICA DE CACHE DE IMAGENS (CORRIGIDA)
    Private Sub LimpezaAutomaticaCacheImagens()
        Try
            ' Executar limpeza a cada 30 minutos
            If DateTime.Now.Subtract(ultimaLimpezaImagensCache).TotalMinutes > CACHE_IMAGENS_TIMEOUT_MINUTES Then
                Dim itensRemovidos As Integer = 0

                ' ✅ CORREÇÃO: Não remover imagem atualmente em uso
                Dim imagemAtualEmUso As Image = Nothing
                If pbProduto IsNot Nothing AndAlso pbProduto.Image IsNot Nothing Then
                    imagemAtualEmUso = pbProduto.Image
                End If

                ' Limpar imagens que não foram usadas recentemente, EXCETO a atual
                Dim chavesParaRemover As New List(Of String)

                For Each kvp In cacheImagens.ToList()
                    ' ✅ NÃO remover a imagem que está sendo exibida atualmente
                    If kvp.Value IsNot Nothing AndAlso Not ReferenceEquals(kvp.Value, imagemAtualEmUso) Then
                        Try
                            kvp.Value.Dispose()
                            itensRemovidos += 1
                            chavesParaRemover.Add(kvp.Key)
                        Catch
                            ' Ignorar erros de dispose
                        End Try
                    ElseIf kvp.Value Is Nothing Then
                        ' Remover entradas vazias
                        chavesParaRemover.Add(kvp.Key)
                    End If
                Next

                ' Remover do cache
                For Each chave In chavesParaRemover
                    cacheImagens.Remove(chave)
                    cacheStatusImagens.Remove(chave)
                Next

                ultimaLimpezaImagensCache = DateTime.Now
                LogErros.RegistrarInfo($"🧹 Cache de imagens limpo: {itensRemovidos} itens removidos, imagem atual preservada", "LimpezaCacheImagens")

                ' Forçar coleta de lixo apenas se removeu muitos itens
                If itensRemovidos > 5 Then
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                    GC.Collect()
                End If
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "LimpezaAutomaticaCacheImagens")
        End Try
    End Sub

    ' Configuração de colunas otimizada
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

            ' Verificar estado do botão periodicamente
            VerificarEstadoBotao()

            ' Limpeza automática de cache de imagens
            LimpezaAutomaticaCacheImagens()

            If dgvProdutos.SelectedRows.Count > 0 Then
                Dim produtoSelecionadoRow As DataGridViewRow = dgvProdutos.SelectedRows(0)

                If produtoSelecionadoRow.Cells.Count > 0 Then
                    Dim codigoProduto As String = If(produtoSelecionadoRow.Cells(0).Value IsNot Nothing,
                                            produtoSelecionadoRow.Cells(0).Value.ToString(), "")

                    If Not String.IsNullOrEmpty(codigoProduto) AndAlso codigoProduto <> produtoSelecionado Then
                        produtoSelecionado = codigoProduto
                        CarregarDadosProdutoUltraRapido(codigoProduto)
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

    ' Métodos para invalidar cache quando atualizar dados
    Private Sub InvalidarCacheCompleto()
        ' Invalidar cache estático
        tabelasEstaticas.Clear()
        ultimaAtualizacaoEstatica = DateTime.MinValue

        ' Invalidar cache individual
        cacheEstoque.Clear()
        cacheCompras.Clear()
        cacheVendas.Clear()
        cacheValido = DateTime.MinValue

        ' Resetar flag de colunas
        colunasConfiguradas = False
    End Sub

    ' ✅ MÉTODOS PÚBLICOS DE DIAGNÓSTICO E UTILITÁRIOS

    ' Método de diagnóstico para medir performance
    Public Sub DiagnosticarPerformance()
        Dim sw As New Stopwatch()

        ' Teste 1: Cache estático
        sw.Start()
        If Not TabelasEstaticasValidas() Then
            CarregarTabelasEstaticas()
        End If
        sw.Stop()
        LogErros.RegistrarInfo($"🔧 Carregamento tabelas estáticas: {sw.ElapsedMilliseconds}ms", "Diagnóstico")

        ' Teste 2: Filtro rápido
        If tabelasEstaticas.ContainsKey("estoque") Then
            sw.Restart()
            Dim resultado = FiltrarRapidoComDataView(tabelasEstaticas("estoque"), produtoSelecionado)
            sw.Stop()
            LogErros.RegistrarInfo($"⚡ Filtro DataView: {sw.ElapsedMilliseconds}ms para {resultado.Rows.Count} registros", "Diagnóstico")
        End If

        ' Teste 3: Aplicação aos grids
        sw.Restart()
        Dim dadosTest As New Dictionary(Of String, System.Data.DataTable)
        dadosTest("estoque") = CriarDataTableVazio()
        dadosTest("compras") = CriarDataTableVazio()
        dadosTest("vendas") = CriarDataTableVazio()

        PararRedesenhoCompleto()
        AplicarDadosUltraRapido(dadosTest, "TEST")
        ReabilitarRedesenhoCompleto()
        sw.Stop()
        LogErros.RegistrarInfo($"📊 Aplicação aos grids: {sw.ElapsedMilliseconds}ms", "Diagnóstico")
    End Sub

    ' Limpeza automática de cache
    Private Sub LimpezaAutomaticaCache()
        ' A cada 30 minutos, limpar cache para liberar memória
        If DateTime.Now.Minute = 0 OrElse DateTime.Now.Minute = 30 Then
            If DateTime.Now.Subtract(ultimaLimpezaCache).TotalMinutes > 25 Then
                InvalidarCacheCompleto()
                GC.Collect()
                ultimaLimpezaCache = DateTime.Now
                LogErros.RegistrarInfo("🧹 Cache limpo automaticamente", "LimpezaCache")
            End If
        End If
    End Sub

    ' ✅ MÉTODOS PÚBLICOS PARA INTEGRAÇÃO

    ' Forçar atualização externa
    Public Sub ForcarAtualizacao()
        Try
            btnAtualizar.PerformClick()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ForcarAtualizacao")
        End Try
    End Sub

    ' Obter informações do estado atual
    Public Function ObterEstadoAtual() As Dictionary(Of String, Object)
        Try
            Dim estado As New Dictionary(Of String, Object)

            estado("DadosCarregados") = dadosCarregados
            estado("ProdutoSelecionado") = produtoSelecionado
            estado("FiltroAtual") = filtroAtual
            estado("TotalProdutos") = If(dgvProdutos.Rows IsNot Nothing, dgvProdutos.Rows.Count, 0)
            estado("CacheValido") = CacheEstaValido()
            estado("TabelasEstaticasValidas") = TabelasEstaticasValidas()
            estado("CacheImagensItens") = cacheImagens.Count
            estado("BotaoHabilitado") = If(btnAtualizar IsNot Nothing, btnAtualizar.Enabled, False)

            Return estado

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ObterEstadoAtual")
            Return New Dictionary(Of String, Object)
        End Try
    End Function

    ' Aplicar filtro externo
    Public Sub AplicarFiltroExterno(filtro As String)
        Try
            txtFiltro.Text = If(filtro, "")
            filtroAtual = txtFiltro.Text.Trim()
            AplicarFiltro()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.AplicarFiltroExterno")
        End Try
    End Sub

    ' Selecionar produto específico
    Public Sub SelecionarProduto(codigoProduto As String)
        Try
            If String.IsNullOrEmpty(codigoProduto) Then Return

            For Each row As DataGridViewRow In dgvProdutos.Rows
                If row.Cells.Count > 0 AndAlso
                   row.Cells(0).Value IsNot Nothing AndAlso
                   row.Cells(0).Value.ToString().Equals(codigoProduto, StringComparison.OrdinalIgnoreCase) Then

                    row.Selected = True
                    dgvProdutos.FirstDisplayedScrollingRowIndex = row.Index
                    Exit For
                End If
            Next

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.SelecionarProduto")
        End Try
    End Sub

    ' ✅ MÉTODO DE LIMPEZA E DISPOSE OTIMIZADO (VERSÃO FINAL)
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

                ' ✅ CORREÇÃO FINAL: Apenas limpar referência, sem dispose
                ' As imagens ficam no cache compartilhado
                If pbProduto IsNot Nothing AndAlso pbProduto.Image IsNot Nothing Then
                    pbProduto.Image = Nothing
                End If

                ' Limpar dados
                If dadosProdutosOriginais IsNot Nothing Then
                    dadosProdutosOriginais.Dispose()
                    dadosProdutosOriginais = Nothing
                End If

                ' Limpar cache de dados (não de imagens)
                InvalidarCache()

                powerQueryManager = Nothing

                LogErros.RegistrarInfo("UcReposicaoEstoque disposed (cache de imagens preservado)", "Dispose")
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.Dispose")
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' ✅ MÉTODO ESTÁTICO PARA LIMPEZA COMPLETA DE CACHE
    Public Shared Sub LimparCacheGlobal()
        Try
            ' Limpar tabelas estáticas
            tabelasEstaticas.Clear()
            ultimaAtualizacaoEstatica = DateTime.MinValue

            ' Limpar cache de imagens
            For Each kvp In cacheImagens.ToList()
                Try
                    If kvp.Value IsNot Nothing Then
                        kvp.Value.Dispose()
                    End If
                Catch
                    ' Ignorar erros de dispose
                End Try
            Next

            cacheImagens.Clear()
            cacheStatusImagens.Clear()

            ' Forçar coleta de lixo
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

            LogErros.RegistrarInfo("🧹 Cache global limpo completamente", "LimparCacheGlobal")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "LimparCacheGlobal")
        End Try
    End Sub

    ' ✅ PROPRIEDADES PÚBLICAS PARA STATUS
    Public ReadOnly Property EstaCarregado As Boolean
        Get
            Return dadosCarregados
        End Get
    End Property

    Public ReadOnly Property ProdutoAtual As String
        Get
            Return produtoSelecionado
        End Get
    End Property

    Public ReadOnly Property TotalProdutosFiltrados As Integer
        Get
            Try
                Return If(dgvProdutos.Rows IsNot Nothing, dgvProdutos.Rows.Count, 0)
            Catch
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property FiltroAplicado As String
        Get
            Return filtroAtual
        End Get
    End Property

    ' ✅ MÉTODO PÚBLICO PARA ACESSO AO DEBUG (PARA MAINFORM)
    Public Sub AcessarDebugImagens()
        DebugCacheImagens()
    End Sub

    ' ✅ MÉTODO DE DEBUG PARA INVESTIGAR CACHE DE IMAGENS
    Public Sub DebugCacheImagens()
        Try
            LogErros.RegistrarInfo("=== DEBUG CACHE DE IMAGENS ===", "DebugCache")
            LogErros.RegistrarInfo($"Total de itens no cache: {cacheImagens.Count}", "DebugCache")
            LogErros.RegistrarInfo($"Total de status no cache: {cacheStatusImagens.Count}", "DebugCache")
            LogErros.RegistrarInfo($"Produto atualmente selecionado: {produtoSelecionado}", "DebugCache")
            LogErros.RegistrarInfo($"Imagem atual no PictureBox: {If(pbProduto.Image IsNot Nothing, "SIM", "NÃO")}", "DebugCache")

            For Each kvp In cacheImagens
                Dim status = If(cacheStatusImagens.ContainsKey(kvp.Key), cacheStatusImagens(kvp.Key), "SEM_STATUS")
                LogErros.RegistrarInfo($"Cache: {kvp.Key} -> Imagem: {If(kvp.Value IsNot Nothing, "SIM", "NÃO")} | Status: {status}", "DebugCache")
            Next

            LogErros.RegistrarInfo("=== FIM DEBUG CACHE ===", "DebugCache")

            ' Também mostrar no MessageBox para debug visual
            Dim msg = $"Cache de Imagens:{vbCrLf}" &
                     $"Total: {cacheImagens.Count} itens{vbCrLf}" &
                     $"Produto atual: {produtoSelecionado}{vbCrLf}" &
                     $"PictureBox tem imagem: {If(pbProduto.Image IsNot Nothing, "SIM", "NÃO")}{vbCrLf}{vbCrLf}"

            For Each kvp In cacheImagens.Take(5) ' Mostrar apenas os 5 primeiros
                Dim status = If(cacheStatusImagens.ContainsKey(kvp.Key), cacheStatusImagens(kvp.Key), "SEM_STATUS")
                msg += $"{kvp.Key}: {If(kvp.Value IsNot Nothing, "✅", "❌")} ({status}){vbCrLf}"
            Next

            MessageBox.Show(msg, "Debug Cache Imagens", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "DebugCacheImagens")
        End Try
    End Sub

    ' ✅ MÉTODO PARA FORÇAR RECARREGAMENTO DE IMAGEM (PARA DEBUG)
    Public Sub ForcarRecarregamentoImagem(codigoProduto As String)
        Try
            ' Remover do cache
            If cacheImagens.ContainsKey(codigoProduto) Then
                If cacheImagens(codigoProduto) IsNot Nothing Then
                    cacheImagens(codigoProduto).Dispose()
                End If
                cacheImagens.Remove(codigoProduto)
            End If

            If cacheStatusImagens.ContainsKey(codigoProduto) Then
                cacheStatusImagens.Remove(codigoProduto)
            End If

            ' Recarregar
            CarregarImagemAsync(codigoProduto)
            LogErros.RegistrarInfo($"Forçado recarregamento de imagem para: {codigoProduto}", "ForcarRecarregamento")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ForcarRecarregamentoImagem")
        End Try
    End Sub

End Class