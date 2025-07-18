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

    ' Cache de imagens
    Private Shared cacheImagens As New Dictionary(Of String, Image)
    Private Shared cacheStatusImagens As New Dictionary(Of String, String)
    Private Shared ultimaLimpezaImagensCache As DateTime = DateTime.MinValue
    Private Const CACHE_IMAGENS_TIMEOUT_MINUTES As Integer = 30

    ' Data da última atualização do Power Query
    Private Shared ultimaAtualizacaoPowerQuery As DateTime = DateTime.MinValue
    Private Const CHAVE_ULTIMA_ATUALIZACAO As String = "UltimaAtualizacaoPowerQuery"

    ' Sistema de sessão de pedidos
    Private pedidoAtual As New Dictionary(Of String, PedidoItem)
    Private numeroPedidoAtual As String = String.Empty
    Private pedidoEmAndamento As Boolean = False
    Private produtosProcessados As New HashSet(Of String)

    ' Classe para armazenar dados do pedido
    Private Class PedidoItem
        Public Property CodigoProduto As String
        Public Property Descricao As String
        Public Property QuantidadePedir As Decimal
        Public Property LojaDestino As String
        Public Property Cliente As String
        Public Property Telefone As String
        Public Property EstoqueAtual As Decimal
        Public Property EstoqueMinimo As Decimal
        Public Property EstoqueMaximo As Decimal
    End Class

    ' APIs do Windows para otimização de redesenho
    <DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As Boolean, lParam As Integer) As Integer
    End Function
    Private Const WM_SETREDRAW As Integer = 11

    Public Sub New()
        InitializeComponent()
        ConfigurarComponentes()
        ' Dados são carregados automaticamente da sessão anterior em ConfigurarComponentes
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

            ' Configurar novos componentes
            ConfigurarNovosComponentes()

            ' Carregar dados iniciais da sessão anterior
            CarregarDadosIniciais()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarComponentes")
            MessageBox.Show($"Erro ao configurar componentes: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ConfigurarNovosComponentes()
        Try
            ' Configurar campo de aplicação do produto
            txtAplicacaoProduto.ReadOnly = True
            txtAplicacaoProduto.BackColor = SystemColors.Control

            ' Configurar campos de inclusão manual
            txtDataManual.Text = DateTime.Now.ToString("dd/MM/yyyy")
            txtCodigoManual.MaxLength = 5

            ' Configurar campos de estoque mínimo/máximo
            txtEstoqueMinimo.ReadOnly = True
            txtEstoqueMaximo.ReadOnly = True
            txtEstoqueMinimo.BackColor = SystemColors.Control
            txtEstoqueMaximo.BackColor = SystemColors.Control

            ' Configurar campo quantidade a pedir
            txtQtdPedir.BackColor = Color.LightYellow
            AddHandler txtQtdPedir.KeyPress, AddressOf TxtQtdPedir_KeyPress
            AddHandler txtQtdPedir.Leave, AddressOf TxtQtdPedir_Leave

            ' Configurar máscara de telefone
            txtTelefone.Mask = "(00)00000-0000"

            ' Configurar botões de sessão
            AtualizarEstadoBotoesSessao()

            ' Configurar gráficos
            ConfigurarGraficos()

            ' Configurar novo botão de atualização do Power Query
            ConfigurarBotaoPowerQuery()

            ' Carregar e exibir data da última atualização
            CarregarUltimaAtualizacao()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ConfigurarNovosComponentes")
        End Try
    End Sub

    Private Sub ConfigurarGraficos()
        Try
            ' Como os Charts foram criados no Designer como DataVisualization.Charting.Chart,
            ' vamos apenas configurar as propriedades básicas aqui
            ' A configuração de Series será feita quando os dados forem carregados

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ConfigurarGraficos")
        End Try
    End Sub

    Private Sub ConfigurarDataGridViewsBasico()
        Try
            ' Configuração básica e rápida - sem estilização pesada
            For Each dgv As DataGridView In {dgvProdutos, dgvEstoque}
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
                    .VirtualMode = False
                    .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                End With
            Next

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarDataGridViewsBasico")
        End Try
    End Sub

    Private Sub ConfigurarPictureBox()
        Try
            pbProduto.SizeMode = PictureBoxSizeMode.Zoom
            pbProduto.BackColor = Color.White
            pbProduto.BorderStyle = BorderStyle.FixedSingle
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarPictureBox")
        End Try
    End Sub

    Private Sub ConfigurarTimers()
        Try
            ' Timer de debounce para seleção de produtos
            debounceTimer = New System.Windows.Forms.Timer()
            debounceTimer.Interval = 300
            AddHandler debounceTimer.Tick, AddressOf DebounceTimer_Tick

            ' Timer para filtro
            filtroTimer = New System.Windows.Forms.Timer()
            filtroTimer.Interval = 500
            AddHandler filtroTimer.Tick, AddressOf FiltroTimer_Tick

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarTimers")
        End Try
    End Sub

    Private Sub ConfigurarFiltros()
        Try
            txtFiltro.Font = New System.Drawing.Font("Segoe UI", 9.0!)
            lblFiltro.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarFiltros")
        End Try
    End Sub

    Private Sub ConfigurarEventos()
        Try
            ' Eventos existentes
            AddHandler dgvProdutos.SelectionChanged, AddressOf DgvProdutos_SelectionChanged
            AddHandler dgvProdutos.CellDoubleClick, AddressOf DgvProdutos_CellDoubleClick
            AddHandler dgvProdutos.DataBindingComplete, AddressOf DgvProdutos_DataBindingComplete
            AddHandler txtFiltro.TextChanged, AddressOf TxtFiltro_TextChanged
            AddHandler txtFiltro.Enter, AddressOf TxtFiltro_Enter
            AddHandler txtFiltro.Leave, AddressOf TxtFiltro_Leave
            AddHandler btnAtualizar.Click, AddressOf BtnAtualizar_Click
            AddHandler dgvProdutos.KeyDown, AddressOf DgvProdutos_KeyDown

            ' Eventos dos novos componentes
            AddHandler btnIncluirProdutoManual.Click, AddressOf BtnIncluirProdutoManual_Click
            AddHandler btnProximoProduto.Click, AddressOf BtnProximoProduto_Click
            AddHandler btnNovoPedido.Click, AddressOf BtnNovoPedido_Click
            AddHandler btnFinalizarPedido.Click, AddressOf BtnFinalizarPedido_Click
            AddHandler btnDescartarPedido.Click, AddressOf BtnDescartarPedido_Click

            ' Validação de campos numéricos
            AddHandler txtCodigoManual.KeyPress, AddressOf CampoNumerico_KeyPress
            AddHandler txtQtdInicialManual.KeyPress, AddressOf CampoNumerico_KeyPress

            LogErros.RegistrarInfo("Eventos configurados", "UcReposicaoEstoque.ConfigurarEventos")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ConfigurarEventos")
        End Try
    End Sub

    ' Evento para campos numéricos
    Private Sub CampoNumerico_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    ' Eventos do campo quantidade a pedir
    Private Sub TxtQtdPedir_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "," AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub TxtQtdPedir_Leave(sender As Object, e As EventArgs)
        Try
            Dim valor As Decimal
            If Decimal.TryParse(txtQtdPedir.Text, valor) Then
                If valor < 0 Then
                    txtQtdPedir.Text = "0"
                    MessageBox.Show("A quantidade não pode ser negativa.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Else
                txtQtdPedir.Text = "0"
            End If

            ' Atualizar estado do pedido
            AtualizarPedidoItem()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "TxtQtdPedir_Leave")
        End Try
    End Sub

    ' Inclusão manual de produto
    Private Sub BtnIncluirProdutoManual_Click(sender As Object, e As EventArgs)
        Try
            ' Validar campos obrigatórios
            If String.IsNullOrWhiteSpace(txtCodigoManual.Text) Then
                MessageBox.Show("O código é obrigatório.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCodigoManual.Focus()
                Return
            End If

            If String.IsNullOrWhiteSpace(txtDescricaoManual.Text) Then
                MessageBox.Show("A descrição é obrigatória.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtDescricaoManual.Focus()
                Return
            End If

            If String.IsNullOrWhiteSpace(txtQtdInicialManual.Text) OrElse Not IsNumeric(txtQtdInicialManual.Text) Then
                MessageBox.Show("A quantidade inicial é obrigatória e deve ser numérica.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtQtdInicialManual.Focus()
                Return
            End If

            ' Verificar se código já existe
            If powerQueryManager IsNot Nothing Then
                If powerQueryManager.VerificarCodigoExistente(txtCodigoManual.Text) Then
                    MessageBox.Show("Este código já existe.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtCodigoManual.Focus()
                    Return
                End If
            End If

            ' Criar novo produto manual
            Dim novoProduto As New Dictionary(Of String, Object)
            novoProduto("Codigo") = txtCodigoManual.Text
            novoProduto("Descricao") = txtDescricaoManual.Text
            novoProduto("Fabricante") = txtFabricanteManual.Text
            novoProduto("QuantidadeInicial") = CDec(txtQtdInicialManual.Text)
            novoProduto("Data") = DateTime.Now

            ' Salvar na tabela tblProdutosManual
            If powerQueryManager IsNot Nothing Then
                powerQueryManager.InserirProdutoManual(novoProduto)

                ' Atualizar grid
                btnAtualizar.PerformClick()

                ' Selecionar o produto recém-criado
                SelecionarProdutoPorCodigo(txtCodigoManual.Text)

                ' Limpar campos
                LimparCamposInclusaoManual()

                MessageBox.Show("Produto incluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "BtnIncluirProdutoManual_Click")
            MessageBox.Show($"Erro ao incluir produto: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LimparCamposInclusaoManual()
        txtCodigoManual.Clear()
        txtDescricaoManual.Clear()
        txtFabricanteManual.Clear()
        txtQtdInicialManual.Clear()
        txtDataManual.Text = DateTime.Now.ToString("dd/MM/yyyy")
    End Sub

    Private Sub SelecionarProdutoPorCodigo(codigo As String)
        Try
            For Each row As DataGridViewRow In dgvProdutos.Rows
                If row.Cells(0).Value?.ToString() = codigo Then
                    dgvProdutos.ClearSelection()
                    row.Selected = True
                    dgvProdutos.FirstDisplayedScrollingRowIndex = row.Index
                    Exit For
                End If
            Next
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "SelecionarProdutoPorCodigo")
        End Try
    End Sub

    ' Sistema de sessão de pedidos
    Private Sub BtnNovoPedido_Click(sender As Object, e As EventArgs)
        Try
            If pedidoEmAndamento AndAlso pedidoAtual.Count > 0 Then
                Dim resultado = MessageBox.Show("Existe um pedido em andamento. Deseja descartar e criar um novo?",
                                              "Confirmar Novo Pedido", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If resultado = DialogResult.No Then Return
            End If

            ' Limpar pedido atual
            pedidoAtual.Clear()
            produtosProcessados.Clear()

            ' Gerar novo número de pedido
            numeroPedidoAtual = $"PED-{DateTime.Now:yyyyMMdd-HHmmss}"

            ' Limpar campos
            LimparCamposPedido()

            ' Resetar seleção
            If dgvProdutos.Rows.Count > 0 Then
                dgvProdutos.ClearSelection()
                dgvProdutos.Rows(0).Selected = True
            End If

            pedidoEmAndamento = True
            AtualizarEstadoBotoesSessao()

            MessageBox.Show($"Novo pedido iniciado: {numeroPedidoAtual}", "Novo Pedido",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "BtnNovoPedido_Click")
            MessageBox.Show($"Erro ao criar novo pedido: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnFinalizarPedido_Click(sender As Object, e As EventArgs)
        Try
            If pedidoAtual.Count = 0 Then
                MessageBox.Show("Não há itens no pedido atual.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Verificar se há pelo menos um item com quantidade
            Dim temQuantidade = pedidoAtual.Any(Function(kvp) kvp.Value.QuantidadePedir > 0)
            If Not temQuantidade Then
                MessageBox.Show("É necessário definir a quantidade a pedir para pelo menos um produto.",
                              "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Mostrar resumo do pedido
            Dim resumo As New System.Text.StringBuilder()
            resumo.AppendLine($"RESUMO DO PEDIDO: {numeroPedidoAtual}")
            resumo.AppendLine(New String("-"c, 50))

            Dim totalItens = 0
            For Each item In pedidoAtual.Values.Where(Function(i) i.QuantidadePedir > 0)
                totalItens += 1
                resumo.AppendLine($"Código: {item.CodigoProduto}")
                resumo.AppendLine($"Descrição: {item.Descricao}")
                resumo.AppendLine($"Quantidade: {item.QuantidadePedir:N2}")
                resumo.AppendLine($"Loja: {item.LojaDestino}")
                If Not String.IsNullOrWhiteSpace(item.Cliente) Then
                    resumo.AppendLine($"Cliente: {item.Cliente}")
                End If
                If Not String.IsNullOrWhiteSpace(item.Telefone) Then
                    resumo.AppendLine($"Telefone: {item.Telefone}")
                End If
                resumo.AppendLine()
            Next

            resumo.AppendLine($"Total de itens: {totalItens}")

            ' Confirmar finalização
            Dim resultado = MessageBox.Show(resumo.ToString() & vbCrLf & "Deseja finalizar este pedido?",
                                          "Confirmar Finalização", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                ' Escolher local para salvar
                Using sfd As New SaveFileDialog()
                    sfd.Filter = "Arquivo Excel|*.xlsx"
                    sfd.FileName = $"{numeroPedidoAtual}.xlsx"
                    sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

                    If sfd.ShowDialog() = DialogResult.OK Then
                        ' Salvar pedido em Excel
                        SalvarPedidoExcel(sfd.FileName)

                        ' Limpar pedido
                        pedidoAtual.Clear()
                        produtosProcessados.Clear()
                        pedidoEmAndamento = False
                        numeroPedidoAtual = String.Empty

                        LimparCamposPedido()
                        AtualizarEstadoBotoesSessao()

                        MessageBox.Show("Pedido finalizado e salvo com sucesso!", "Sucesso",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "BtnFinalizarPedido_Click")
            MessageBox.Show($"Erro ao finalizar pedido: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SalvarPedidoExcel(caminho As String)
        Try
            Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
            Dim workbook As Workbook = excelApp.Workbooks.Add()
            Dim worksheet As Worksheet = CType(workbook.Sheets(1), Worksheet)

            ' Título
            worksheet.Cells(1, 1) = $"PEDIDO DE REPOSIÇÃO - {numeroPedidoAtual}"
            worksheet.Range("A1:H1").Merge()
            worksheet.Range("A1").Font.Bold = True
            worksheet.Range("A1").Font.Size = 14
            worksheet.Range("A1").HorizontalAlignment = XlHAlign.xlHAlignCenter

            ' Data
            worksheet.Cells(2, 1) = "Data:"
            worksheet.Cells(2, 2) = DateTime.Now.ToString("dd/MM/yyyy HH:mm")

            ' Cabeçalhos
            Dim linha = 4
            worksheet.Cells(linha, 1) = "Código"
            worksheet.Cells(linha, 2) = "Descrição"
            worksheet.Cells(linha, 3) = "Estoque Atual"
            worksheet.Cells(linha, 4) = "Est. Mínimo"
            worksheet.Cells(linha, 5) = "Est. Máximo"
            worksheet.Cells(linha, 6) = "Qtd Pedir"
            worksheet.Cells(linha, 7) = "Loja Destino"
            worksheet.Cells(linha, 8) = "Cliente"
            worksheet.Cells(linha, 9) = "Telefone"

            ' Formatar cabeçalhos
            Dim rangeHeaders = worksheet.Range("A4:I4")
            rangeHeaders.Font.Bold = True
            rangeHeaders.Interior.Color = RGB(200, 200, 200)

            ' Dados
            linha = 5
            For Each item In pedidoAtual.Values.Where(Function(i) i.QuantidadePedir > 0)
                worksheet.Cells(linha, 1) = item.CodigoProduto
                worksheet.Cells(linha, 2) = item.Descricao
                worksheet.Cells(linha, 3) = item.EstoqueAtual
                worksheet.Cells(linha, 4) = item.EstoqueMinimo
                worksheet.Cells(linha, 5) = item.EstoqueMaximo
                worksheet.Cells(linha, 6) = item.QuantidadePedir
                worksheet.Cells(linha, 7) = item.LojaDestino
                worksheet.Cells(linha, 8) = item.Cliente
                worksheet.Cells(linha, 9) = item.Telefone
                linha += 1
            Next

            ' Ajustar largura das colunas
            worksheet.Columns("A:I").AutoFit()

            ' Salvar e fechar
            workbook.SaveAs(caminho)
            workbook.Close()
            excelApp.Quit()

            ' Liberar objetos COM
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "SalvarPedidoExcel")
            Throw
        End Try
    End Sub

    Private Sub BtnDescartarPedido_Click(sender As Object, e As EventArgs)
        Try
            If Not pedidoEmAndamento OrElse pedidoAtual.Count = 0 Then
                MessageBox.Show("Não há pedido em andamento para descartar.", "Aviso",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim resultado = MessageBox.Show("Tem certeza que deseja descartar o pedido atual?",
                                          "Confirmar Descarte", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                pedidoAtual.Clear()
                produtosProcessados.Clear()
                pedidoEmAndamento = False
                numeroPedidoAtual = String.Empty

                LimparCamposPedido()
                AtualizarEstadoBotoesSessao()

                MessageBox.Show("Pedido descartado.", "Pedido Descartado",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "BtnDescartarPedido_Click")
            MessageBox.Show($"Erro ao descartar pedido: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnProximoProduto_Click(sender As Object, e As EventArgs)
        Try
            ' Salvar informações do produto atual
            AtualizarPedidoItem()

            ' Marcar produto como processado
            If Not String.IsNullOrEmpty(produtoSelecionado) Then
                produtosProcessados.Add(produtoSelecionado)
            End If

            ' Encontrar próximo produto não processado
            Dim proximoIndice = -1
            Dim indiceAtual = If(dgvProdutos.SelectedRows.Count > 0, dgvProdutos.SelectedRows(0).Index, -1)

            For i = indiceAtual + 1 To dgvProdutos.Rows.Count - 1
                Dim codigo = dgvProdutos.Rows(i).Cells(0).Value?.ToString()
                If Not String.IsNullOrEmpty(codigo) AndAlso Not produtosProcessados.Contains(codigo) Then
                    proximoIndice = i
                    Exit For
                End If
            Next

            If proximoIndice >= 0 Then
                dgvProdutos.ClearSelection()
                dgvProdutos.Rows(proximoIndice).Selected = True
                dgvProdutos.FirstDisplayedScrollingRowIndex = proximoIndice
            Else
                MessageBox.Show("Não há mais produtos para processar.", "Aviso",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "BtnProximoProduto_Click")
            MessageBox.Show($"Erro ao avançar para próximo produto: {ex.Message}", "Erro",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub AtualizarPedidoItem()
        Try
            If String.IsNullOrEmpty(produtoSelecionado) Then Return

            Dim quantidade As Decimal
            If Not Decimal.TryParse(txtQtdPedir.Text, quantidade) Then
                quantidade = 0
            End If

            If quantidade > 0 OrElse pedidoAtual.ContainsKey(produtoSelecionado) Then
                If Not pedidoAtual.ContainsKey(produtoSelecionado) Then
                    pedidoAtual(produtoSelecionado) = New PedidoItem()
                End If

                With pedidoAtual(produtoSelecionado)
                    .CodigoProduto = produtoSelecionado
                    .Descricao = If(dgvProdutos.SelectedRows.Count > 0,
                                   dgvProdutos.SelectedRows(0).Cells(1).Value?.ToString(), "")
                    .QuantidadePedir = quantidade
                    .LojaDestino = comboLoja.Text
                    .Cliente = txtCliente.Text
                    .Telefone = txtTelefone.Text
                    .EstoqueMinimo = CDec(Val(txtEstoqueMinimo.Text))
                    .EstoqueMaximo = CDec(Val(txtEstoqueMaximo.Text))

                    ' Obter estoque atual do dgvEstoque
                    If dgvEstoque.DataSource IsNot Nothing Then
                        Dim dt = TryCast(dgvEstoque.DataSource, System.Data.DataTable)
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            .EstoqueAtual = CDec(Val(dt.Rows(0)("Disponível")))
                        End If
                    End If
                End With

                pedidoEmAndamento = True
                AtualizarEstadoBotoesSessao()
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AtualizarPedidoItem")
        End Try
    End Sub

    Private Sub LimparCamposPedido()
        txtQtdPedir.Text = "0"
        txtCliente.Clear()
        txtTelefone.Clear()
    End Sub

    Private Sub AtualizarEstadoBotoesSessao()
        Try
            ' Habilitar/desabilitar botões baseado no estado
            btnFinalizarPedido.Enabled = pedidoEmAndamento AndAlso pedidoAtual.Any(Function(kvp) kvp.Value.QuantidadePedir > 0)
            btnDescartarPedido.Enabled = pedidoEmAndamento AndAlso pedidoAtual.Count > 0

            ' Atualizar visual dos botões
            If btnFinalizarPedido.Enabled Then
                btnFinalizarPedido.BackColor = Color.FromArgb(220, 53, 69)
            Else
                btnFinalizarPedido.BackColor = Color.LightGray
            End If

            If btnDescartarPedido.Enabled Then
                btnDescartarPedido.BackColor = Color.FromArgb(108, 117, 125)
            Else
                btnDescartarPedido.BackColor = Color.LightGray
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AtualizarEstadoBotoesSessao")
        End Try
    End Sub

    Private Sub CarregarLojas()
        Try
            comboLoja.Items.Clear()

            If powerQueryManager IsNot Nothing Then
                Dim lojas = powerQueryManager.ObterLojasDistintas()
                For Each loja In lojas
                    comboLoja.Items.Add(loja)
                Next

                ' Selecionar Cariacica como padrão se existir
                Dim indexCariacica = comboLoja.Items.IndexOf("Cariacica")
                If indexCariacica >= 0 Then
                    comboLoja.SelectedIndex = indexCariacica
                ElseIf comboLoja.Items.Count > 0 Then
                    comboLoja.SelectedIndex = 0
                End If
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "CarregarLojas")
        End Try
    End Sub

    Private Sub AtualizarCamposProdutoSelecionado()
        Try
            ' Atualizar aplicação do produto (placeholder por enquanto)
            txtAplicacaoProduto.Text = ""  ' TODO: Preencher quando a coluna Aplicacao for implementada

            ' Atualizar estoque mínimo/máximo (placeholder por enquanto)
            txtEstoqueMinimo.Text = "0"  ' TODO: Preencher quando as colunas forem implementadas
            txtEstoqueMaximo.Text = "0"

            ' Calcular quantidade a pedir
            Dim estoqueMaximo = CDec(Val(txtEstoqueMaximo.Text))
            Dim estoqueAtual As Decimal = 0

            If dgvEstoque.DataSource IsNot Nothing Then
                Dim dt = TryCast(dgvEstoque.DataSource, System.Data.DataTable)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    estoqueAtual = CDec(Val(dt.Rows(0)("Disponível")))
                End If
            End If

            Dim qtdSugerida = Math.Max(0, estoqueMaximo - estoqueAtual)

            ' Verificar se já existe valor no pedido
            If pedidoAtual.ContainsKey(produtoSelecionado) Then
                txtQtdPedir.Text = pedidoAtual(produtoSelecionado).QuantidadePedir.ToString("N2")
                txtCliente.Text = pedidoAtual(produtoSelecionado).Cliente
                txtTelefone.Text = pedidoAtual(produtoSelecionado).Telefone
                If Not String.IsNullOrEmpty(pedidoAtual(produtoSelecionado).LojaDestino) Then
                    comboLoja.Text = pedidoAtual(produtoSelecionado).LojaDestino
                End If
            Else
                txtQtdPedir.Text = qtdSugerida.ToString("N2")
                txtCliente.Clear()
                txtTelefone.Clear()
            End If

            ' Atualizar gráficos
            AtualizarGraficos()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AtualizarCamposProdutoSelecionado")
        End Try
    End Sub

    Private Sub AtualizarGraficos()
        Try
            If String.IsNullOrEmpty(produtoSelecionado) Then
                LogErros.RegistrarInfo("⚠️ Produto não selecionado, pulando atualização dos gráficos", "AtualizarGraficos")
                Return
            End If

            LogErros.RegistrarInfo($"📈 Atualizando gráficos para produto: {produtoSelecionado}", "AtualizarGraficos")

            ' ✅ CORREÇÃO: Inicializar séries se não existirem
            If chartComprasMensais.Series.Count = 0 Then
                chartComprasMensais.Series.Add("Compras")
                chartComprasMensais.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Column
                LogErros.RegistrarInfo("✅ Série de compras criada", "AtualizarGraficos")
            Else
                chartComprasMensais.Series(0).Points.Clear()
            End If
            
            If chartVendasMensais.Series.Count = 0 Then
                chartVendasMensais.Series.Add("Vendas")
                chartVendasMensais.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Column
                LogErros.RegistrarInfo("✅ Série de vendas criada", "AtualizarGraficos")
            Else
                chartVendasMensais.Series(0).Points.Clear()
            End If

            If powerQueryManager IsNot Nothing Then
                ' Obter dados dos últimos 24 meses
                Dim dataInicio = DateTime.Now.AddMonths(-23).Date
                Dim dataFim = DateTime.Now.Date
                
                LogErros.RegistrarInfo($"📅 Período: {dataInicio:dd/MM/yyyy} até {dataFim:dd/MM/yyyy}", "AtualizarGraficos")

                ' ✅ CORREÇÃO: Dados de compras com verificação de debug
                Dim dadosCompras = powerQueryManager.ObterHistoricoComprasPorMes(produtoSelecionado, dataInicio, dataFim)
                LogErros.RegistrarInfo($"📊 Compras encontradas: {dadosCompras.Count} meses", "AtualizarGraficos")
                
                For Each item In dadosCompras
                    Try
                        Dim ponto = chartComprasMensais.Series(0).Points.AddXY(item.Key.ToString("MMM/yy"), item.Value)
                        ' Destacar picos
                        If dadosCompras.Values.Count > 1 AndAlso item.Value > dadosCompras.Values.Average() * 1.5 Then
                            chartComprasMensais.Series(0).Points(ponto).MarkerSize = 12
                            chartComprasMensais.Series(0).Points(ponto).MarkerColor = Color.DarkBlue
                        End If
                    Catch pointEx As Exception
                        LogErros.RegistrarErro(pointEx, $"AtualizarGraficos.AdicionarPontoCompras - {item.Key}: {item.Value}")
                    End Try
                Next

                ' ✅ CORREÇÃO: Dados de vendas com verificação de debug
                Dim dadosVendas = powerQueryManager.ObterHistoricoVendasPorMes(produtoSelecionado, dataInicio, dataFim)
                LogErros.RegistrarInfo($"📊 Vendas encontradas: {dadosVendas.Count} meses", "AtualizarGraficos")
                
                For Each item In dadosVendas
                    Try
                        Dim ponto = chartVendasMensais.Series(0).Points.AddXY(item.Key.ToString("MMM/yy"), item.Value)
                        ' Destacar picos
                        If dadosVendas.Values.Count > 1 AndAlso item.Value > dadosVendas.Values.Average() * 1.5 Then
                            chartVendasMensais.Series(0).Points(ponto).MarkerSize = 12
                            chartVendasMensais.Series(0).Points(ponto).MarkerColor = Color.DarkRed
                        End If
                    Catch pointEx As Exception
                        LogErros.RegistrarErro(pointEx, $"AtualizarGraficos.AdicionarPontoVendas - {item.Key}: {item.Value}")
                    End Try
                Next
                
                LogErros.RegistrarInfo($"✅ Gráficos atualizados - Compras: {chartComprasMensais.Series(0).Points.Count} pontos, Vendas: {chartVendasMensais.Series(0).Points.Count} pontos", "AtualizarGraficos")
            Else
                LogErros.RegistrarInfo("⚠️ PowerQueryManager é null, não foi possível obter dados dos gráficos", "AtualizarGraficos")
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AtualizarGraficos")
        End Try
    End Sub

    Private Sub AtualizarGraficosComDados(dados As Dictionary(Of String, System.Data.DataTable), codigoProduto As String)
        Try
            If String.IsNullOrEmpty(codigoProduto) Then
                LogErros.RegistrarInfo("⚠️ Produto não informado, pulando atualização dos gráficos", "AtualizarGraficosComDados")
                Return
            End If

            LogErros.RegistrarInfo($"📈 Atualizando gráficos com dados carregados para produto: {codigoProduto}", "AtualizarGraficosComDados")

            ' Inicializar séries se não existirem
            If chartComprasMensais.Series.Count = 0 Then
                chartComprasMensais.Series.Add("Compras")
                chartComprasMensais.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Column
                LogErros.RegistrarInfo("✅ Série de compras criada", "AtualizarGraficosComDados")
            Else
                chartComprasMensais.Series(0).Points.Clear()
            End If
            
            If chartVendasMensais.Series.Count = 0 Then
                chartVendasMensais.Series.Add("Vendas")
                chartVendasMensais.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Column
                LogErros.RegistrarInfo("✅ Série de vendas criada", "AtualizarGraficosComDados")
            Else
                chartVendasMensais.Series(0).Points.Clear()
            End If

            ' Processar dados de compras diretamente do DataTable
            If dados.ContainsKey("compras") Then
                Dim dadosCompras = dados("compras")
                LogErros.RegistrarInfo($"📊 Processando {dadosCompras.Rows.Count} registros de compras", "AtualizarGraficosComDados")
                
                If dadosCompras.Rows.Count > 0 Then
                    PreencherGraficoDeDataTable(chartComprasMensais, dadosCompras, "Compras")
                Else
                    ' Sem dados de compras - exibir gráfico vazio com mensagem
                    ExibirGraficoVazio(chartComprasMensais, "Sem dados de compras")
                End If
            Else
                ExibirGraficoVazio(chartComprasMensais, "Sem dados de compras")
            End If

            ' Processar dados de vendas diretamente do DataTable
            If dados.ContainsKey("vendas") Then
                Dim dadosVendas = dados("vendas")
                LogErros.RegistrarInfo($"📊 Processando {dadosVendas.Rows.Count} registros de vendas", "AtualizarGraficosComDados")
                
                If dadosVendas.Rows.Count > 0 Then
                    PreencherGraficoDeDataTable(chartVendasMensais, dadosVendas, "Vendas")
                Else
                    ' Sem dados de vendas - exibir gráfico vazio com mensagem
                    ExibirGraficoVazio(chartVendasMensais, "Sem dados de vendas")
                End If
            Else
                ExibirGraficoVazio(chartVendasMensais, "Sem dados de vendas")
            End If

            ' Atualizar contadores dos títulos dos grupos
            Dim pontosCompras = If(chartComprasMensais.Series.Count > 0, chartComprasMensais.Series(0).Points.Count, 0)
            grpCompras.Text = $"📈 Compras ({pontosCompras} meses)"
            
            Dim pontosVendas = If(chartVendasMensais.Series.Count > 0, chartVendasMensais.Series(0).Points.Count, 0)
            grpVendas.Text = $"📉 Vendas ({pontosVendas} meses)"

            LogErros.RegistrarInfo("✅ Gráficos atualizados com dados carregados", "AtualizarGraficosComDados")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AtualizarGraficosComDados")
        End Try
    End Sub

    Private Sub PreencherGraficoDeDataTable(chart As DataVisualization.Charting.Chart, dataTable As System.Data.DataTable, tipoOperacao As String)
        Try
            If dataTable Is Nothing OrElse dataTable.Rows.Count = 0 Then
                LogErros.RegistrarInfo($"⚠️ Nenhum dado de {tipoOperacao.ToLower()} para processar", "PreencherGraficoDeDataTable")
                ExibirGraficoVazio(chart, $"Sem dados de {tipoOperacao.ToLower()}")
                Return
            End If

            ' Debug: Listar todas as colunas disponíveis
            Dim colunas = String.Join(", ", dataTable.Columns.Cast(Of DataColumn)().Select(Function(c) c.ColumnName))
            LogErros.RegistrarInfo($"📋 Colunas disponíveis para {tipoOperacao}: {colunas}", "PreencherGraficoDeDataTable")

            ' Agrupar dados por mês/ano (assumindo que há colunas de data e quantidade/valor)
            Dim dadosPorMes As New Dictionary(Of DateTime, Decimal)

            ' Tentar encontrar colunas de data e valor
            Dim colunaData As DataColumn = Nothing
            Dim colunaValor As DataColumn = Nothing

            ' Buscar coluna de data (várias possibilidades)
            For Each col As DataColumn In dataTable.Columns
                Dim nomeCol = col.ColumnName.ToLower()
                If nomeCol.Contains("data") OrElse nomeCol.Contains("date") OrElse nomeCol.Contains("mes") OrElse nomeCol.Contains("ano") Then
                    colunaData = col
                    LogErros.RegistrarInfo($"🗓️ Coluna de data encontrada: {col.ColumnName}", "PreencherGraficoDeDataTable")
                    Exit For
                End If
            Next

            ' Buscar coluna de valor/quantidade (mais opções)
            For Each col As DataColumn In dataTable.Columns
                Dim nomeCol = col.ColumnName.ToLower()
                If nomeCol.Contains("quantidade") OrElse nomeCol.Contains("valor") OrElse nomeCol.Contains("qtd") OrElse 
                   nomeCol.Contains("preco") OrElse nomeCol.Contains("total") OrElse nomeCol.Contains("amount") OrElse
                   nomeCol.Contains("volume") OrElse nomeCol.Contains("sum") Then
                    colunaValor = col
                    LogErros.RegistrarInfo($"💰 Coluna de valor encontrada: {col.ColumnName}", "PreencherGraficoDeDataTable")
                    Exit For
                End If
            Next

            ' Se não encontrou colunas específicas, tentar usar as primeiras disponíveis
            If colunaData Is Nothing AndAlso dataTable.Columns.Count > 0 Then
                colunaData = dataTable.Columns(0)
                LogErros.RegistrarInfo($"🔄 Usando primeira coluna como data: {colunaData.ColumnName}", "PreencherGraficoDeDataTable")
            End If

            If colunaValor Is Nothing AndAlso dataTable.Columns.Count > 1 Then
                colunaValor = dataTable.Columns(1)
                LogErros.RegistrarInfo($"🔄 Usando segunda coluna como valor: {colunaValor.ColumnName}", "PreencherGraficoDeDataTable")
            End If

            If colunaData IsNot Nothing AndAlso colunaValor IsNot Nothing Then
                ' Processar dados com colunas identificadas
                For Each row As DataRow In dataTable.Rows
                    Try
                        Dim dataOperacao As DateTime
                        Dim valor As Decimal = 0

                        If DateTime.TryParse(row(colunaData).ToString(), dataOperacao) AndAlso
                           Decimal.TryParse(row(colunaValor).ToString(), valor) Then
                            
                            Dim mesAno = New DateTime(dataOperacao.Year, dataOperacao.Month, 1)
                            
                            If dadosPorMes.ContainsKey(mesAno) Then
                                dadosPorMes(mesAno) += valor
                            Else
                                dadosPorMes(mesAno) = valor
                            End If
                        End If
                    Catch rowEx As Exception
                        LogErros.RegistrarErro(rowEx, $"PreencherGraficoDeDataTable.ProcessarLinha_{tipoOperacao}")
                    End Try
                Next

                ' Adicionar pontos ao gráfico
                If dadosPorMes.Count > 0 Then
                    For Each item In dadosPorMes.OrderBy(Function(x) x.Key)
                        Try
                            chart.Series(0).Points.AddXY(item.Key.ToString("MMM/yy"), item.Value)
                        Catch pointEx As Exception
                            LogErros.RegistrarErro(pointEx, $"PreencherGraficoDeDataTable.AdicionarPonto_{tipoOperacao}")
                        End Try
                    Next

                    LogErros.RegistrarInfo($"✅ Gráfico de {tipoOperacao.ToLower()} preenchido com {dadosPorMes.Count} pontos", "PreencherGraficoDeDataTable")
                Else
                    LogErros.RegistrarInfo($"⚠️ Nenhum ponto válido gerado para {tipoOperacao.ToLower()}", "PreencherGraficoDeDataTable")
                    ExibirGraficoVazio(chart, $"Dados inválidos para {tipoOperacao.ToLower()}")
                End If
            Else
                LogErros.RegistrarInfo($"⚠️ Colunas de data/valor não encontradas para {tipoOperacao.ToLower()}", "PreencherGraficoDeDataTable")
                
                ' Fallback: Tentar criar um gráfico simples com dados brutos
                If dataTable.Rows.Count <= 24 Then ' Máximo 24 pontos
                    TentarGraficoSimples(chart, dataTable, tipoOperacao)
                Else
                    ExibirGraficoVazio(chart, $"Estrutura de dados não reconhecida para {tipoOperacao.ToLower()}")
                End If
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"PreencherGraficoDeDataTable_{tipoOperacao}")
        End Try
    End Sub

    Private Sub ExibirGraficoVazio(chart As DataVisualization.Charting.Chart, mensagem As String)
        Try
            ' Limpar pontos existentes
            If chart.Series.Count > 0 Then
                chart.Series(0).Points.Clear()
            End If

            ' Adicionar um ponto com valor zero para mostrar que o gráfico está vazio
            If chart.Series.Count > 0 Then
                chart.Series(0).Points.AddXY("Sem dados", 0)
                chart.Series(0).Points(0).Color = Color.LightGray
                chart.Series(0).Points(0).Label = mensagem
            End If

            LogErros.RegistrarInfo($"📊 Gráfico configurado como vazio: {mensagem}", "ExibirGraficoVazio")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ExibirGraficoVazio")
        End Try
    End Sub

    Private Sub TentarGraficoSimples(chart As DataVisualization.Charting.Chart, dataTable As System.Data.DataTable, tipoOperacao As String)
        Try
            LogErros.RegistrarInfo($"🔄 Tentando gráfico simples para {tipoOperacao} com {dataTable.Rows.Count} linhas", "TentarGraficoSimples")

            Dim pontoAdicionado As Boolean = False

            ' Tentar usar primeira e segunda colunas como X e Y
            If dataTable.Columns.Count >= 2 Then
                Dim contador As Integer = 0
                For Each row As DataRow In dataTable.Rows
                    If contador >= 24 Then Exit For ' Limitar a 24 pontos
                    
                    Try
                        Dim x = If(row(0) IsNot Nothing, row(0).ToString(), $"Item {contador + 1}")
                        Dim y As Decimal = 0
                        
                        ' Tentar converter segunda coluna para número
                        If Not Decimal.TryParse(row(1).ToString(), y) Then
                            y = contador + 1 ' Valor padrão crescente
                        End If

                        chart.Series(0).Points.AddXY(x, y)
                        pontoAdicionado = True
                        contador += 1

                    Catch rowEx As Exception
                        LogErros.RegistrarErro(rowEx, $"TentarGraficoSimples.ProcessarLinha_{tipoOperacao}")
                    End Try
                Next
            End If

            If pontoAdicionado Then
                LogErros.RegistrarInfo($"✅ Gráfico simples criado para {tipoOperacao} com {chart.Series(0).Points.Count} pontos", "TentarGraficoSimples")
            Else
                LogErros.RegistrarInfo($"⚠️ Não foi possível criar gráfico simples para {tipoOperacao}", "TentarGraficoSimples")
                ExibirGraficoVazio(chart, $"Falha ao processar dados de {tipoOperacao.ToLower()}")
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"TentarGraficoSimples_{tipoOperacao}")
            ExibirGraficoVazio(chart, $"Erro ao processar {tipoOperacao.ToLower()}")
        End Try
    End Sub

    ' Eventos existentes
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
                        AtualizarCamposProdutoSelecionado()
                    End If
                End If
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DebounceTimer_Tick")
        End Try
    End Sub

    Private Sub BtnAtualizar_Click(sender As Object, e As EventArgs)
        Try
            ' Aplicar filtro nos dados já carregados
            AplicarFiltro()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.BtnAtualizar_Click")
            MessageBox.Show($"Erro ao aplicar filtro: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RestaurarBotaoAtualizar()
        Try
            btnAtualizar.Enabled = True
            btnAtualizar.Text = "Filtrar"
            btnAtualizar.BackColor = Color.FromArgb(0, 123, 255)
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "RestaurarBotaoAtualizar")
        End Try
    End Sub

    Private Sub ConfigurarBotaoPowerQuery()
        Try
            ' Configurar evento do novo botão
            AddHandler btnAtualizarPowerQuery.Click, AddressOf BtnAtualizarPowerQuery_Click
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ConfigurarBotaoPowerQuery")
        End Try
    End Sub

    Private Sub BtnAtualizarPowerQuery_Click(sender As Object, e As EventArgs)
        Try
            ' Desabilitar botão imediatamente
            btnAtualizarPowerQuery.Enabled = False
            btnAtualizarPowerQuery.Text = "Atualizando..."
            btnAtualizarPowerQuery.BackColor = Color.Orange

            ' Mostrar status
            AtualizarStatus("Atualizando dados do Power Query...")

            ' Executar em background para não travar UI
            Task.Run(Sub()
                         Try
                             ' Atualizar Power Query
                             powerQueryManager = PowerQueryManager.GetInstance()
                             If powerQueryManager IsNot Nothing Then
                                 powerQueryManager.AtualizarDados()
                                 
                                 ' Debug das tabelas disponíveis
                                 Dim tabelasDisponiveis = powerQueryManager.ListarTabelas()
                                 LogErros.RegistrarInfo($"📁 Tabelas disponíveis após atualização: {String.Join(", ", tabelasDisponiveis)}", "BtnAtualizarPowerQuery_Click")
                             End If

                             ' Atualizar data da última atualização
                             SalvarUltimaAtualizacao()

                             ' Voltar para thread principal
                             Me.Invoke(Sub()
                                           Try
                                               ' Carregar dados
                                               CarregarProdutos()

                                               ' Carregar lojas para o combo
                                               CarregarLojas()

                                               ' Limpar caches
                                               InvalidarCacheCompleto()

                                               ' Atualizar label da data
                                               AtualizarLabelUltimaAtualizacao()

                                               ' Mostrar mensagem de sucesso
                                               MessageBox.Show("Dados do Power Query atualizados com sucesso!", "Sucesso",
                                                         MessageBoxButtons.OK, MessageBoxIcon.Information)
                                           Catch loadEx As Exception
                                               LogErros.RegistrarErro(loadEx, "BtnAtualizarPowerQuery_Click.CarregarProdutos")
                                               AtualizarStatus("Erro ao carregar produtos após atualização")
                                               MessageBox.Show($"Erro ao carregar produtos: {loadEx.Message}", "Erro",
                                                         MessageBoxButtons.OK, MessageBoxIcon.Error)
                                           End Try
                                       End Sub)

                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, "UcReposicaoEstoque.BtnAtualizarPowerQuery_Click.Background")
                             Me.Invoke(Sub()
                                           AtualizarStatus("Erro na atualização dos dados")
                                           MessageBox.Show($"Erro ao atualizar dados: {ex.Message}", "Erro",
                                                     MessageBoxButtons.OK, MessageBoxIcon.Error)
                                       End Sub)
                         Finally
                             ' SEMPRE restaurar o botão
                             Me.Invoke(Sub()
                                           RestaurarBotaoPowerQuery()
                                       End Sub)
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.BtnAtualizarPowerQuery_Click")
            AtualizarStatus("Erro na atualização")
            MessageBox.Show($"Erro ao atualizar dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            RestaurarBotaoPowerQuery()
        End Try
    End Sub

    Private Sub RestaurarBotaoPowerQuery()
        Try
            btnAtualizarPowerQuery.Enabled = True
            btnAtualizarPowerQuery.Text = "📊 Atualizar"
            btnAtualizarPowerQuery.BackColor = Color.FromArgb(40, 167, 69)
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "RestaurarBotaoPowerQuery")
        End Try
    End Sub

    Private Sub CarregarUltimaAtualizacao()
        Try
            ' Tentar carregar da configuração (pode ser Registry, arquivo config, etc.)
            ' Por simplicidade, vamos usar um arquivo temporário no diretório de logs
            Dim arquivoConfig As String = Path.Combine("C:\Logs\GestaoEstoque", "ultima_atualizacao.txt")
            
            If File.Exists(arquivoConfig) Then
                Dim textoData As String = File.ReadAllText(arquivoConfig).Trim()
                If DateTime.TryParse(textoData, ultimaAtualizacaoPowerQuery) Then
                    AtualizarLabelUltimaAtualizacao()
                Else
                    ultimaAtualizacaoPowerQuery = DateTime.MinValue
                    lblUltimaAtualizacao.Text = "Última atualização: Nunca"
                End If
            Else
                ultimaAtualizacaoPowerQuery = DateTime.MinValue
                lblUltimaAtualizacao.Text = "Última atualização: Nunca"
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "CarregarUltimaAtualizacao")
            lblUltimaAtualizacao.Text = "Última atualização: Nunca"
        End Try
    End Sub

    Private Sub SalvarUltimaAtualizacao()
        Try
            ultimaAtualizacaoPowerQuery = DateTime.Now
            
            ' Salvar em arquivo temporário
            Dim arquivoConfig As String = Path.Combine("C:\Logs\GestaoEstoque", "ultima_atualizacao.txt")
            Dim diretorio As String = Path.GetDirectoryName(arquivoConfig)
            
            If Not Directory.Exists(diretorio) Then
                Directory.CreateDirectory(diretorio)
            End If
            
            File.WriteAllText(arquivoConfig, ultimaAtualizacaoPowerQuery.ToString("yyyy-MM-dd HH:mm:ss"))
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "SalvarUltimaAtualizacao")
        End Try
    End Sub

    Private Sub AtualizarLabelUltimaAtualizacao()
        Try
            If ultimaAtualizacaoPowerQuery = DateTime.MinValue Then
                lblUltimaAtualizacao.Text = "Última atualização: Nunca"
            Else
                lblUltimaAtualizacao.Text = $"✅ Atualizado: {ultimaAtualizacaoPowerQuery:dd/MM/yyyy HH:mm}"
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AtualizarLabelUltimaAtualizacao")
        End Try
    End Sub

    Private Sub CarregarDadosIniciais()
        Try
            ' Mostrar status de carregamento inicial
            AtualizarStatus("Verificando dados da sessão anterior...")
            
            ' Executar em background para não travar a UI
            Task.Run(Sub()
                         Try
                             ' Inicializar PowerQueryManager se necessário
                             If powerQueryManager Is Nothing Then
                                 powerQueryManager = PowerQueryManager.GetInstance()
                             End If

                             If powerQueryManager IsNot Nothing Then
                                 ' Verificar se existem tabelas Power Query disponíveis
                                 Dim tabelasDisponiveis = powerQueryManager.ListarTabelas()
                                 LogErros.RegistrarInfo($"Tabelas encontradas na inicialização: {String.Join(", ", tabelasDisponiveis)}", "CarregarDadosIniciais")
                                 
                                 ' Verificar especificamente se tblProdutos existe
                                 Dim tabelaProdutos = powerQueryManager.ObterTabela("tblProdutos")
                                 
                                 If tabelaProdutos IsNot Nothing Then
                                     ' Há dados disponíveis - carregar
                                     Me.Invoke(Sub()
                                                   Try
                                                       ' Carregar produtos sem atualizar Power Query
                                                       CarregarProdutosSemAtualizacao()
                                                       
                                                       ' Carregar lojas para o combo
                                                       CarregarLojas()
                                                       
                                                       ' Mostrar status de sucesso
                                                       AtualizarStatus($"Dados carregados ({DateTime.Now:HH:mm}) - Use 📊 Atualizar para dados mais recentes")
                                                       
                                                       ' Marcar como dados carregados
                                                       dadosCarregados = True
                                                       
                                                       ' Atualizar indicador de que dados são da sessão anterior
                                                       If ultimaAtualizacaoPowerQuery <> DateTime.MinValue Then
                                                           lblUltimaAtualizacao.Text = $"📄 Dados da sessão: {ultimaAtualizacaoPowerQuery:dd/MM/yyyy HH:mm}"
                                                       End If
                                                       
                                                   Catch loadEx As Exception
                                                       LogErros.RegistrarErro(loadEx, "CarregarDadosIniciais.CarregarDados")
                                                       AtualizarStatus("Erro ao carregar dados - Clique em 📊 Atualizar")
                                                   End Try
                                               End Sub)
                                 Else
                                     ' Não há dados disponíveis - primeira execução
                                     Me.Invoke(Sub()
                                                   AtualizarStatus("Primeira execução - Clique em 📊 Atualizar para carregar dados")
                                                   LogErros.RegistrarInfo("Nenhuma tabela Power Query encontrada - primeira execução", "CarregarDadosIniciais")
                                               End Sub)
                                 End If
                             Else
                                 Me.Invoke(Sub()
                                               AtualizarStatus("PowerQuery não disponível - Clique em 📊 Atualizar")
                                           End Sub)
                             End If

                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, "CarregarDadosIniciais.Background")
                             Me.Invoke(Sub()
                                           AtualizarStatus("Erro na verificação inicial - Clique em 📊 Atualizar")
                                       End Sub)
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "CarregarDadosIniciais")
            AtualizarStatus("Erro na verificação inicial - Clique em 📊 Atualizar")
        End Try
    End Sub

    Private Sub CarregarProdutosSemAtualizacao()
        Try
            ' Tentar usar dados já existentes do PowerQuery sem atualizar
            dadosProdutosOriginais = powerQueryManager.ObterProdutos()

            If dadosProdutosOriginais IsNot Nothing AndAlso dadosProdutosOriginais.Rows.Count > 0 Then
                ' Aplicar dados ao grid
                dgvProdutos.DataSource = dadosProdutosOriginais
                
                ' Atualizar contador de produtos
                AtualizarContadorProdutos(dadosProdutosOriginais.Rows.Count)
                
                ' Marcar colunas como configuradas
                If Not colunasConfiguradas Then
                    colunasConfiguradas = True
                End If

                ' Aplicar filtro se existir
                If Not String.IsNullOrEmpty(filtroAtual) Then
                    AplicarFiltro()
                End If

                LogErros.RegistrarInfo($"Produtos carregados da sessão anterior: {dadosProdutosOriginais.Rows.Count} itens", "CarregarProdutosSemAtualizacao")
            Else
                ' Se não há dados, mostrar mensagem para atualizar
                AtualizarStatus("Nenhum dado encontrado - Clique em 📊 Atualizar para carregar")
                LogErros.RegistrarInfo("Nenhum produto encontrado na sessão anterior", "CarregarProdutosSemAtualizacao")
                
                ' Mostrar mensagem mais específica no label
                lblUltimaAtualizacao.Text = "⚠️ Primeira execução - Execute 📊 Atualizar"
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "CarregarProdutosSemAtualizacao")
            AtualizarStatus("Erro ao carregar produtos - Clique em 📊 Atualizar")
            lblUltimaAtualizacao.Text = "❌ Erro no carregamento"
        End Try
    End Sub

    Private Sub CarregarProdutos()
        Try
            If powerQueryManager Is Nothing Then
                powerQueryManager = PowerQueryManager.GetInstance()
            End If

            If powerQueryManager IsNot Nothing Then
                dadosProdutosOriginais = powerQueryManager.ObterProdutos()

                If dadosProdutosOriginais IsNot Nothing AndAlso dadosProdutosOriginais.Rows.Count > 0 Then
                    dgvProdutos.DataSource = dadosProdutosOriginais
                    AtualizarContadorProdutos(dadosProdutosOriginais.Rows.Count)
                    dadosCarregados = True
                    carregamentoInicial = False
                    AtualizarStatus($"Dados carregados: {dadosProdutosOriginais.Rows.Count} produtos")
                Else
                    AtualizarStatus("Nenhum produto encontrado")
                End If
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.CarregarProdutos")
            Throw
        End Try
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

                ' SEMPRE carregar imagem, mesmo com cache de dados
                CarregarImagemAsync(codigoProduto)
                Return
            End If

            ' 2. PARAR COMPLETAMENTE o redesenho
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

    Private Function CarregarDadosOtimizado(codigoProduto As String) As Dictionary(Of String, System.Data.DataTable)
        Try
            Dim dados As New Dictionary(Of String, System.Data.DataTable)

            If powerQueryManager IsNot Nothing Then
                dados("estoque") = powerQueryManager.ObterEstoqueProduto(codigoProduto)
                dados("compras") = powerQueryManager.ObterComprasProduto(codigoProduto)
                dados("vendas") = powerQueryManager.ObterVendasProduto(codigoProduto)
            Else
                dados("estoque") = New System.Data.DataTable()
                dados("compras") = New System.Data.DataTable()
                dados("vendas") = New System.Data.DataTable()
            End If

            ' Armazenar no cache
            ArmazenarNoCacheIndividual(codigoProduto, dados)

            Return dados

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "CarregarDadosOtimizado")
            Return New Dictionary(Of String, System.Data.DataTable)
        End Try
    End Function

    Private Sub AplicarDadosUltraRapido(dados As Dictionary(Of String, System.Data.DataTable), codigoProduto As String)
        Try
            ' ✅ CORREÇÃO: Aplicar dados com verificação de debug
            Dim dadosEstoque = dados("estoque")
            LogErros.RegistrarInfo($"📊 Aplicando dados de estoque: {dadosEstoque.Rows.Count} registros para produto {codigoProduto}", "AplicarDadosUltraRapido")
            
            dgvEstoque.DataSource = dadosEstoque

            ' Configurar colunas apenas UMA VEZ
            If Not colunasConfiguradas Then
                ConfigurarTodasAsColunasUmaVez()
                colunasConfiguradas = True
            End If

            ' ✅ CORREÇÃO: Atualizar contadores com verificação de séries
            grpEstoque.Text = $"📊 Estoque Atual ({dadosEstoque.Rows.Count} registros)"
            
            Try
                Dim pontosCompras = If(chartComprasMensais.Series.Count > 0, chartComprasMensais.Series(0).Points.Count, 0)
                grpCompras.Text = $"📈 Compras ({pontosCompras} meses)"
            Catch
                grpCompras.Text = "📈 Histórico de Compras (24 meses)"
            End Try
            
            Try
                Dim pontosVendas = If(chartVendasMensais.Series.Count > 0, chartVendasMensais.Series(0).Points.Count, 0)
                grpVendas.Text = $"📉 Vendas ({pontosVendas} meses)"
            Catch
                grpVendas.Text = "📉 Histórico de Vendas (24 meses)"
            End Try
            
            ' ✅ CORREÇÃO: Atualizar gráficos com dados de compras e vendas
            AtualizarGraficosComDados(dados, codigoProduto)
            
            LogErros.RegistrarInfo($"✅ Dados aplicados com sucesso no dgvEstoque", "AplicarDadosUltraRapido")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AplicarDadosUltraRapido")
        End Try
    End Sub

    ' Métodos auxiliares para cache e otimização
    Private Sub PararRedesenhoCompleto()
        Try
            SendMessage(dgvEstoque.Handle, WM_SETREDRAW, False, 0)
            dgvEstoque.SuspendLayout()
            dgvEstoque.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PararRedesenhoCompleto")
        End Try
    End Sub

    Private Sub ReabilitarRedesenhoCompleto()
        Try
            dgvEstoque.ResumeLayout(False)

            If Not colunasConfiguradas Then
                dgvEstoque.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
            End If

            SendMessage(dgvEstoque.Handle, WM_SETREDRAW, True, 0)
            dgvEstoque.Invalidate()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ReabilitarRedesenhoCompleto")
        End Try
    End Sub

    Private Function VerificarCacheIndividual(codigoProduto As String) As Boolean
        Return CacheEstaValido() AndAlso
           cacheEstoque.ContainsKey($"estoque_{codigoProduto}") AndAlso
           cacheCompras.ContainsKey($"compras_{codigoProduto}") AndAlso
           cacheVendas.ContainsKey($"vendas_{codigoProduto}")
    End Function

    Private Function CacheEstaValido() As Boolean
        Return DateTime.Now.Subtract(cacheValido).TotalMinutes < CACHE_TIMEOUT_MINUTES
    End Function

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

    Private Sub AplicarDadosDoCache(codigoProduto As String)
        Try
            PararRedesenhoCompleto()

            Try
                dgvEstoque.DataSource = cacheEstoque($"estoque_{codigoProduto}")

                ' Atualizar contadores
                grpEstoque.Text = $"📊 Estoque Atual ({dgvEstoque.Rows.Count} registros)"

                LogErros.RegistrarInfo($"✅ Dados aplicados do cache para produto: {codigoProduto}", "AplicarDadosDoCache")

            Finally
                ReabilitarRedesenhoCompleto()
            End Try

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AplicarDadosDoCache")
        End Try
    End Sub

    Private Sub InvalidarCacheCompleto()
        Try
            cacheEstoque.Clear()
            cacheCompras.Clear()
            cacheVendas.Clear()
            cacheValido = DateTime.MinValue

            ' Limpar cache de imagens também
            For Each img In cacheImagens.Values
                img?.Dispose()
            Next
            cacheImagens.Clear()
            cacheStatusImagens.Clear()

            LogErros.RegistrarInfo("Cache completo invalidado", "InvalidarCacheCompleto")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "InvalidarCacheCompleto")
        End Try
    End Sub

    Private Sub CarregarImagemAsync(codigoProduto As String)
        Task.Run(Sub()
                     Try
                         CarregarImagemProdutoAsync(codigoProduto)
                     Catch ex As Exception
                         LogErros.RegistrarErro(ex, "CarregarImagemAsync")
                     End Try
                 End Sub)
    End Sub

    Private Sub CarregarImagemProdutoAsync(codigoProduto As String)
        If isCarregandoImagem Then Return

        Try
            LogErros.RegistrarInfo($"🔍 Iniciando carregamento de imagem para: {codigoProduto}", "CarregarImagem")

            ' ✅ CORREÇÃO: Forçar recarregamento sempre para corrigir problema de cache
            ' Verificar cache de imagem primeiro (mas não aplicar se produto mudou)
            If cacheImagens.ContainsKey(codigoProduto) AndAlso produtoSelecionado = codigoProduto Then
                LogErros.RegistrarInfo($"📦 Imagem encontrada no cache para: {codigoProduto}", "CarregarImagem")
                If Me.InvokeRequired Then
                    Me.Invoke(Sub() AplicarImagemDoCache(codigoProduto))
                Else
                    AplicarImagemDoCache(codigoProduto)
                End If
                Return
            End If

            ' Verificar status cache
            If cacheStatusImagens.ContainsKey(codigoProduto) Then
                Dim status = cacheStatusImagens(codigoProduto)
                If status = "NAO_ENCONTRADA" OrElse status = "ERRO" Then
                    AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Não disponível", Nothing)
                    Return
                End If
            End If

            isCarregandoImagem = True
            AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Carregando...", Nothing)

            ' Buscar imagem
            Dim caminhoImagem = BuscarImagemProduto(codigoProduto)

            If Not String.IsNullOrEmpty(caminhoImagem) AndAlso File.Exists(caminhoImagem) Then
                Try
                    ' ✅ CORREÇÃO: Criar cópia da imagem para evitar problemas de concorrência
                    Using fs As New FileStream(caminhoImagem, FileMode.Open, FileAccess.Read)
                        Using tempImage = Image.FromStream(fs)
                            ' Criar cópia para evitar lock do arquivo
                            Dim imagem As New Bitmap(tempImage)
                            
                            ' Adicionar ao cache
                            cacheImagens(codigoProduto) = imagem
                            cacheStatusImagens(codigoProduto) = "OK"

                            ' Aplicar imagem
                            If Me.InvokeRequired Then
                                Me.Invoke(Sub() AplicarImagem(imagem, codigoProduto))
                            Else
                                AplicarImagem(imagem, codigoProduto)
                            End If
                        End Using
                    End Using
                Catch imageLoadEx As Exception
                    LogErros.RegistrarErro(imageLoadEx, $"CarregarImagemProdutoAsync.CarregarArquivo - {caminhoImagem}")
                    cacheStatusImagens(codigoProduto) = "ERRO_ARQUIVO"
                    AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Arquivo corrompido", Nothing)
                End Try
            Else
                cacheStatusImagens(codigoProduto) = "NAO_ENCONTRADA"
                AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Não encontrada", Nothing)
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "CarregarImagemProdutoAsync")
            cacheStatusImagens(codigoProduto) = "ERRO"
            AplicarImagemStatus(codigoProduto, "🖼️ Imagem do Produto - Erro", Nothing)
        Finally
            isCarregandoImagem = False
        End Try
    End Sub

    Private Function BuscarImagemProduto(codigo As String) As String
        Try
            Dim caminhoBase = ConfiguracaoApp.CAMINHO_IMAGENS

            For Each ext In ConfiguracaoApp.EXTENSOES_IMAGEM
                Dim arquivo = Path.Combine(caminhoBase, $"{codigo}{ext}")
                If File.Exists(arquivo) Then
                    Return arquivo
                End If
            Next

            Return String.Empty

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "BuscarImagemProduto")
            Return String.Empty
        End Try
    End Function

    Private Sub AplicarImagem(imagem As Image, codigoProduto As String)
        Try
            ' ✅ CORREÇÃO: Verificar se a imagem é válida antes de aplicar
            If imagem IsNot Nothing Then
                Try
                    ' Testar se a imagem é válida
                    Dim testSize = imagem.Size
                    
                    ' Limpar imagem anterior apenas se nova imagem é válida
                    If pbProduto.Image IsNot Nothing Then
                        Dim oldImage = pbProduto.Image
                        pbProduto.Image = Nothing
                        oldImage.Dispose()
                    End If
                    
                    pbProduto.Image = imagem
                    grpImagem.Text = $"🖼️ Imagem do Produto - {codigoProduto}"
                    
                    LogErros.RegistrarInfo($"✅ Imagem aplicada com sucesso para produto: {codigoProduto}", "AplicarImagem")
                    
                Catch imageEx As Exception
                    LogErros.RegistrarErro(imageEx, $"AplicarImagem.ValidarImagem - {codigoProduto}")
                    grpImagem.Text = $"🖼️ Imagem do Produto - {codigoProduto} (Erro)"
                End Try
            Else
                grpImagem.Text = $"🖼️ Imagem do Produto - {codigoProduto} (Não disponível)"
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AplicarImagem")
            grpImagem.Text = $"🖼️ Imagem do Produto - Erro"
        End Try
    End Sub

    Private Sub AplicarImagemDoCache(codigoProduto As String)
        Try
            If cacheImagens.ContainsKey(codigoProduto) Then
                AplicarImagem(cacheImagens(codigoProduto), codigoProduto)
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AplicarImagemDoCache")
        End Try
    End Sub

    Private Sub AplicarImagemStatus(codigoProduto As String, status As String, imagem As Image)
        Try
            If Me.InvokeRequired Then
                Me.Invoke(Sub() AplicarImagemStatus(codigoProduto, status, imagem))
                Return
            End If

            If pbProduto.Image IsNot Nothing Then
                pbProduto.Image.Dispose()
            End If

            pbProduto.Image = imagem
            grpImagem.Text = status

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "AplicarImagemStatus")
        End Try
    End Sub

    Private Sub ConfigurarTodasAsColunasUmaVez()
        Try
            ConfigurarColunasEstoqueOtimizado()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ConfigurarTodasAsColunasUmaVez")
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
                    New ColumnConfig(5, "Transf.Pend.", 90, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(6, "Empenho", 80, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(7, "Venda", 80, True, DataGridViewContentAlignment.MiddleRight, "N2"),
                    New ColumnConfig(8, "Compra", 80, True, DataGridViewContentAlignment.MiddleRight, "N2")
                )
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ConfigurarColunasEstoqueOtimizado")
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
                ' Filtro otimizado
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
            grpProdutos.Text = $"📦 Lista de Produtos ({quantidade} registros)"
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.AtualizarContadorProdutos")
        End Try
    End Sub

    Private Sub LimparDadosSecundarios()
        Try
            dgvEstoque.DataSource = Nothing
            chartComprasMensais.Series(0).Points.Clear()
            chartVendasMensais.Series(0).Points.Clear()

            grpEstoque.Text = "📊 Estoque Atual"
            grpCompras.Text = "📈 Histórico de Compras (24 meses)"
            grpVendas.Text = "📉 Histórico de Vendas (24 meses)"

            If pbProduto.Image IsNot Nothing Then
                pbProduto.Image.Dispose()
                pbProduto.Image = Nothing
            End If

            ' Limpar campos novos
            txtAplicacaoProduto.Clear()
            txtEstoqueMinimo.Text = "0"
            txtEstoqueMaximo.Text = "0"
            txtQtdPedir.Text = "0"
            txtCliente.Clear()
            txtTelefone.Clear()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.LimparDadosSecundarios")
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
                    info.AppendLine($"{col.HeaderText}: {row.Cells(col.Index).Value}")
                Next

                MessageBox.Show(info.ToString(), "Detalhes do Produto", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DgvProdutos_CellDoubleClick")
        End Try
    End Sub

    Private Sub AtualizarStatus(mensagem As String)
        Try
            ' Se houver uma barra de status, atualizar aqui
            LogErros.RegistrarInfo(mensagem, "Status")
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.AtualizarStatus")
        End Try
    End Sub

    Private Sub VerificarEstadoBotao()
        Try
            If btnAtualizar IsNot Nothing AndAlso Not btnAtualizar.Enabled Then
                ' Se o botão está desabilitado há muito tempo, reabilitar
                If DateTime.Now.Subtract(ultimaAtualizacaoEstatica).TotalMinutes > 2 Then
                    RestaurarBotaoAtualizar()
                End If
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "VerificarEstadoBotao")
        End Try
    End Sub

    Private Sub LimpezaAutomaticaCacheImagens()
        Try
            If DateTime.Now.Subtract(ultimaLimpezaImagensCache).TotalMinutes > CACHE_IMAGENS_TIMEOUT_MINUTES Then
                Dim imagemAtualEmUso = pbProduto.Image
                Dim itensRemovidos = 0
                Dim chavesParaRemover As New List(Of String)

                For Each kvp In cacheImagens.ToList()
                    If kvp.Value IsNot Nothing AndAlso Not ReferenceEquals(kvp.Value, imagemAtualEmUso) Then
                        Try
                            kvp.Value.Dispose()
                            itensRemovidos += 1
                            chavesParaRemover.Add(kvp.Key)
                        Catch
                            ' Ignorar erros de dispose
                        End Try
                    ElseIf kvp.Value Is Nothing Then
                        chavesParaRemover.Add(kvp.Key)
                    End If
                Next

                For Each chave In chavesParaRemover
                    cacheImagens.Remove(chave)
                    cacheStatusImagens.Remove(chave)
                Next

                ultimaLimpezaImagensCache = DateTime.Now
                LogErros.RegistrarInfo($"🧹 Cache de imagens limpo: {itensRemovidos} itens removidos", "LimpezaCacheImagens")

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

    ' Métodos públicos para integração
    Public Sub ForcarAtualizacao()
        Try
            btnAtualizar.PerformClick()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ForcarAtualizacao")
        End Try
    End Sub

    Public Function ObterEstadoAtual() As Dictionary(Of String, Object)
        Try
            Dim estado As New Dictionary(Of String, Object)

            estado("DadosCarregados") = dadosCarregados
            estado("ProdutoSelecionado") = produtoSelecionado
            estado("FiltroAtual") = filtroAtual
            estado("TotalProdutos") = If(dgvProdutos.Rows IsNot Nothing, dgvProdutos.Rows.Count, 0)
            estado("CacheValido") = CacheEstaValido()
            estado("CacheImagensItens") = cacheImagens.Count
            estado("BotaoHabilitado") = If(btnAtualizar IsNot Nothing, btnAtualizar.Enabled, False)
            estado("PedidoEmAndamento") = pedidoEmAndamento
            estado("NumeroPedido") = numeroPedidoAtual
            estado("ItensPedido") = pedidoAtual.Count
            estado("ProdutosProcessados") = produtosProcessados.Count

            Return estado

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ObterEstadoAtual")
            Return New Dictionary(Of String, Object)
        End Try
    End Function

    Public Sub DefinirProdutoSelecionado(codigo As String)
        Try
            If Not String.IsNullOrEmpty(codigo) Then
                SelecionarProdutoPorCodigo(codigo)
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.DefinirProdutoSelecionado")
        End Try
    End Sub

    ' Propriedade para obter o produto atualmente selecionado
    Public ReadOnly Property ProdutoAtual As String
        Get
            Return produtoSelecionado
        End Get
    End Property

    ' Método para maximizar no container
    Public Sub MaximizarNoContainer()
        Try
            If Me.Parent IsNot Nothing Then
                Me.Dock = DockStyle.Fill
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.MaximizarNoContainer")
        End Try
    End Sub

    ' Método para forçar maximização
    Public Sub ForcarMaximizacao()
        Try
            MaximizarNoContainer()
            Me.BringToFront()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ForcarMaximizacao")
        End Try
    End Sub

    ' Método para acessar debug de imagens
    Public Sub AcessarDebugImagens()
        Try
            Dim debugInfo As New System.Text.StringBuilder()
            debugInfo.AppendLine("=== DEBUG DE IMAGENS ===")
            debugInfo.AppendLine($"Cache de imagens: {cacheImagens.Count} itens")
            debugInfo.AppendLine($"Cache de status: {cacheStatusImagens.Count} itens")
            debugInfo.AppendLine($"Produto selecionado: {produtoSelecionado}")
            debugInfo.AppendLine($"Carregando imagem: {isCarregandoImagem}")
            debugInfo.AppendLine()

            debugInfo.AppendLine("Itens no cache:")
            For Each kvp In cacheImagens
                debugInfo.AppendLine($"  - {kvp.Key}: {If(kvp.Value IsNot Nothing, "OK", "NULL")}")
            Next

            debugInfo.AppendLine()
            debugInfo.AppendLine("Status no cache:")
            For Each kvp In cacheStatusImagens
                debugInfo.AppendLine($"  - {kvp.Key}: {kvp.Value}")
            Next

            MessageBox.Show(debugInfo.ToString(), "Debug de Imagens", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.AcessarDebugImagens")
        End Try
    End Sub

    ' Método para forçar recarregamento de imagem
    Public Sub ForcarRecarregamentoImagem(codigoProduto As String)
        Try
            ' Remover do cache
            If cacheImagens.ContainsKey(codigoProduto) Then
                Dim img = cacheImagens(codigoProduto)
                If img IsNot Nothing AndAlso Not ReferenceEquals(img, pbProduto.Image) Then
                    img.Dispose()
                End If
                cacheImagens.Remove(codigoProduto)
            End If

            If cacheStatusImagens.ContainsKey(codigoProduto) Then
                cacheStatusImagens.Remove(codigoProduto)
            End If

            ' Recarregar
            CarregarImagemAsync(codigoProduto)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.ForcarRecarregamentoImagem")
        End Try
    End Sub

    ' Método para limpar cache global
    Public Sub LimparCacheGlobal()
        Try
            ' Preservar imagem atual
            Dim imagemAtual = pbProduto.Image

            ' Limpar cache de imagens
            For Each kvp In cacheImagens
                If kvp.Value IsNot Nothing AndAlso Not ReferenceEquals(kvp.Value, imagemAtual) Then
                    Try
                        kvp.Value.Dispose()
                    Catch
                        ' Ignorar erros de dispose
                    End Try
                End If
            Next

            cacheImagens.Clear()
            cacheStatusImagens.Clear()

            ' Limpar outros caches
            InvalidarCacheCompleto()

            ' Forçar coleta de lixo
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

            MessageBox.Show("Cache global limpo com sucesso!", "Limpeza de Cache",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.LimparCacheGlobal")
            MessageBox.Show($"Erro ao limpar cache: {ex.Message}", "Erro",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Limpeza ao descarregar
    Public Sub LimparRecursos()
        Try
            ' Parar timers
            If debounceTimer IsNot Nothing Then
                debounceTimer.Stop()
                debounceTimer.Dispose()
            End If

            If filtroTimer IsNot Nothing Then
                filtroTimer.Stop()
                filtroTimer.Dispose()
            End If

            ' Limpar caches
            InvalidarCacheCompleto()

            ' Limpar imagem atual
            If pbProduto.Image IsNot Nothing Then
                pbProduto.Image.Dispose()
                pbProduto.Image = Nothing
            End If

            LogErros.RegistrarInfo("Recursos liberados", "UcReposicaoEstoque.LimparRecursos")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "UcReposicaoEstoque.LimparRecursos")
        End Try
    End Sub

    Private Sub UcReposicaoEstoque_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        LimparRecursos()
    End Sub

End Class
