Imports System.Drawing
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports WinFormsApp = System.Windows.Forms.Application

Public Class MainForm
    Private ucReposicaoEstoque As UcReposicaoEstoque
    Private timerStatusBar As System.Windows.Forms.Timer
    Private moduloAtual As String = ""
    Private isClosingControlled As Boolean = False
    Private isDisposing As Boolean = False
    Private WithEvents btnTestes As Button
    Private carregamentoInicial As Boolean = True

    Public Sub New()
        InitializeComponent()
        ConfigurarFormulario()
        InicializarStatusBar()
        ' NÃO carregar UserControl inicial aqui - será carregado sob demanda
    End Sub

    Private Sub ConfigurarFormulario()
        Try
            ' Configurações básicas e rápidas
            Me.Text = "Gestão de Estoque - Difemaq"
            Me.Size = New Size(1400, 900)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.WindowState = FormWindowState.Maximized
            Me.MinimumSize = New Size(1200, 800)
            Me.ShowInTaskbar = True
            Me.KeyPreview = True

            ' Configurar evento de teclas
            AddHandler Me.KeyDown, AddressOf MainForm_KeyDown

            ' Criar botão de testes
            CriarBotaoTestes()

            LogErros.RegistrarInfo("Formulário principal configurado rapidamente", "MainForm.ConfigurarFormulario")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.ConfigurarFormulario")
        End Try
    End Sub

    ' Carregar UserControl apenas quando solicitado (lazy loading)
    Private Sub CarregarUserControlInicial()
        Try
            If carregamentoInicial Then
                carregamentoInicial = False

                ' Mostrar mensagem de carregamento
                AtualizarStatus("Carregando módulo de Reposição de Estoque...")

                ' Carregar em background para não travar a UI
                Task.Run(Sub()
                             Try
                                 Me.Invoke(Sub()
                                               CarregarUserControl(GetType(UcReposicaoEstoque))
                                               HighlightButtonMenu(btnReposicaoEstoque)
                                               moduloAtual = "Reposição de Estoque"
                                               AtualizarStatus("Sistema pronto")
                                           End Sub)
                             Catch ex As Exception
                                 LogErros.RegistrarErro(ex, "MainForm.CarregarUserControlInicial.Background")
                                 Me.Invoke(Sub()
                                               AtualizarStatus("Erro ao carregar módulo")
                                               MessageBox.Show($"Erro ao carregar interface inicial: {ex.Message}",
                                                             "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                           End Sub)
                             End Try
                         End Sub)
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.CarregarUserControlInicial")
        End Try
    End Sub

    Public Sub CarregarUserControl(tipoUserControl As Type)
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

            ' Mostrar cursor de carregamento apenas se necessário
            Dim mostrarCursor = pnlConteudo.Controls.Count > 0
            If mostrarCursor Then
                Me.Cursor = Cursors.WaitCursor
                AtualizarStatus("Carregando módulo...")
            End If

            ' Limpar painel atual
            LimparPainelConteudo()

            ' Criar instância do UserControl
            Dim novoUserControl As UserControl = CType(Activator.CreateInstance(tipoUserControl), UserControl)

            ' Configurar e adicionar ao painel
            With novoUserControl
                .Dock = DockStyle.Fill
                .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            End With

            pnlConteudo.Controls.Add(novoUserControl)

            ' Armazenar referência se for UcReposicaoEstoque
            If TypeOf novoUserControl Is UcReposicaoEstoque Then
                ucReposicaoEstoque = CType(novoUserControl, UcReposicaoEstoque)
            End If

            ' Forçar redesenho apenas se necessário
            If mostrarCursor Then
                pnlConteudo.Refresh()
                Me.Refresh()
            End If

            LogErros.RegistrarInfo($"UserControl carregado: {tipoUserControl.Name}", "MainForm.CarregarUserControl")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"MainForm.CarregarUserControl({tipoUserControl.Name})")
            MessageBox.Show($"Erro ao carregar módulo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            If pnlConteudo.Controls.Count > 0 Then
                AtualizarStatus("Pronto")
            End If
        End Try
    End Sub

    Private Sub btnReposicaoEstoque_Click(sender As Object, e As EventArgs) Handles btnReposicaoEstoque.Click
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

            ' Carregar inicial se necessário
            If carregamentoInicial Then
                CarregarUserControlInicial()
                Return
            End If

            ' Apenas carregar se não for o módulo atual
            If moduloAtual <> "Reposição de Estoque" Then
                AtualizarStatus("Carregando Reposição de Estoque...")
                HighlightButtonMenu(btnReposicaoEstoque)
                CarregarUserControl(GetType(UcReposicaoEstoque))
                moduloAtual = "Reposição de Estoque"
                AtualizarStatus("Módulo Reposição de Estoque carregado")
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.btnReposicaoEstoque_Click")
            AtualizarStatus("Erro ao carregar módulo")
        End Try
    End Sub

    ' Resto dos métodos permanecem iguais, mas otimizados onde possível...

    Private Sub LimparPainelConteudo()
        Try
            If pnlConteudo Is Nothing OrElse pnlConteudo.IsDisposed Then Return

            ' Suspender layout para melhor performance
            pnlConteudo.SuspendLayout()

            Try
                ' Limpar controles de forma otimizada
                For Each control As Control In pnlConteudo.Controls.Cast(Of Control).ToArray()
                    If TypeOf control Is UserControl Then
                        control.Dispose()
                    End If
                    pnlConteudo.Controls.Remove(control)
                Next

                ucReposicaoEstoque = Nothing

            Finally
                pnlConteudo.ResumeLayout(True)
            End Try

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.LimparPainelConteudo")
        End Try
    End Sub

    Private Sub InicializarStatusBar()
        Try
            ' Timer otimizado para status bar
            timerStatusBar = New System.Windows.Forms.Timer()
            timerStatusBar.Interval = 1000
            AddHandler timerStatusBar.Tick, AddressOf TimerStatusBar_Tick
            timerStatusBar.Start()

            AtualizarStatus("Sistema inicializado - Clique em 'Reposição Estoque' para começar")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.InicializarStatusBar")
        End Try
    End Sub

    Private Sub TimerStatusBar_Tick(sender As Object, e As EventArgs)
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

            If lblDataHora IsNot Nothing AndAlso Not lblDataHora.IsDisposed Then
                lblDataHora.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.TimerStatusBar_Tick")
        End Try
    End Sub

    Public Sub AtualizarStatus(mensagem As String)
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

            If lblStatus IsNot Nothing AndAlso Not lblStatus.IsDisposed Then
                lblStatus.Text = mensagem
            End If

            ' Processar eventos apenas se necessário
            If Not carregamentoInicial Then
                WinFormsApp.DoEvents()
            End If

            LogErros.RegistrarInfo($"Status atualizado: {mensagem}", "MainForm.AtualizarStatus")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.AtualizarStatus")
        End Try
    End Sub

    Private Sub CriarBotaoTestes()
        Try
            btnTestes = New Button()

            With btnTestes
                .Name = "btnTestes"
                .Text = "🧪 Testes"
                .Size = New Size(100, 40)
                .Location = New Point(430, 10)
                .BackColor = Color.Transparent
                .FlatStyle = FlatStyle.Flat
                .Font = New Font("Segoe UI", 10.0!, FontStyle.Bold)
                .TabIndex = 3
                .UseVisualStyleBackColor = False
                .Visible = False

                With .FlatAppearance
                    .BorderSize = 0
                End With
            End With

            If pnlMenu IsNot Nothing Then
                pnlMenu.Controls.Add(btnTestes)
            End If

            AddHandler btnTestes.Click, AddressOf btnTestes_Click

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.CriarBotaoTestes")
        End Try
    End Sub

    ' Métodos de controle de janela otimizados
    Private Sub btnMinimizar_Click(sender As Object, e As EventArgs) Handles btnMinimizar.Click
        Try
            If isDisposing OrElse Me.IsDisposed Then Return
            Me.WindowState = FormWindowState.Minimized
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.btnMinimizar_Click")
        End Try
    End Sub

    Private Sub btnMaximizar_Click(sender As Object, e As EventArgs) Handles btnMaximizar.Click
        Try
            If isDisposing OrElse Me.IsDisposed Then Return
            If btnMaximizar Is Nothing OrElse btnMaximizar.IsDisposed Then Return

            If Me.WindowState = FormWindowState.Maximized Then
                Me.WindowState = FormWindowState.Normal
                btnMaximizar.Text = "□"
            Else
                Me.WindowState = FormWindowState.Maximized
                btnMaximizar.Text = "❐"
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.btnMaximizar_Click")
        End Try
    End Sub

    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Try
            FecharAplicacaoSegura()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.btnFechar_Click")
        End Try
    End Sub

    Private Sub FecharAplicacaoSegura()
        Try
            isClosingControlled = True
            isDisposing = True
            Me.Close()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.FecharAplicacaoSegura")
        End Try
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            isDisposing = True

            If Not isClosingControlled Then
                Dim resultado As DialogResult = MessageBox.Show(
                    "Deseja realmente fechar o sistema?",
                    "Confirmar Fechamento",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2)

                If resultado = DialogResult.No Then
                    isDisposing = False
                    e.Cancel = True
                    Return
                End If
            End If

            LimparRecursos()

            Try
                If Globals.ThisWorkbook IsNot Nothing Then
                    Dim thisWb As ThisWorkbook = CType(Globals.ThisWorkbook, ThisWorkbook)
                    If thisWb IsNot Nothing Then
                        thisWb.FecharAplicacao()
                    End If
                End If
            Catch ex As Exception
                LogErros.RegistrarErro(ex, "MainForm.MainForm_FormClosing - Erro ao fechar workbook")
            End Try

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.MainForm_FormClosing")
        End Try
    End Sub

    Private Sub LimparRecursos()
        Try
            LogErros.RegistrarInfo("Limpando recursos do MainForm", "MainForm.LimparRecursos")

            If timerStatusBar IsNot Nothing Then
                timerStatusBar.Stop()
                timerStatusBar.Dispose()
                timerStatusBar = Nothing
            End If

            LimparPainelConteudo()

            LogErros.RegistrarInfo("Recursos do MainForm limpos", "MainForm.LimparRecursos")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.LimparRecursos")
        End Try
    End Sub

    ' Eventos de menu otimizados
    Private Sub btnRelatorios_Click(sender As Object, e As EventArgs) Handles btnRelatorios.Click
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

            AtualizarStatus("Módulo de Relatórios em desenvolvimento...")
            HighlightButtonMenu(btnRelatorios)

            ' Não travar a UI com MessageBox longo
            Task.Run(Sub()
                         Thread.Sleep(100) ' Pequena pausa para UI responder
                         Me.Invoke(Sub()
                                       Dim mensagem = $"Módulo de Relatórios{Environment.NewLine}{Environment.NewLine}Funcionalidades planejadas:{Environment.NewLine}• Relatório de estoque baixo{Environment.NewLine}• Análise de vendas por período{Environment.NewLine}• Histórico de compras{Environment.NewLine}• Gráficos de movimentação{Environment.NewLine}• Exportação para Excel/PDF{Environment.NewLine}{Environment.NewLine}Será implementado em versão futura."

                                       MessageBox.Show(mensagem, "Relatórios - Em Desenvolvimento",
                                                     MessageBoxButtons.OK, MessageBoxIcon.Information)

                                       If moduloAtual = "Reposição de Estoque" Then
                                           HighlightButtonMenu(btnReposicaoEstoque)
                                       End If

                                       AtualizarStatus("Pronto")
                                   End Sub)
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.btnRelatorios_Click")
        End Try
    End Sub

    Private Sub btnConfiguracoes_Click(sender As Object, e As EventArgs) Handles btnConfiguracoes.Click
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

            AtualizarStatus("Módulo de Configurações em desenvolvimento...")
            HighlightButtonMenu(btnConfiguracoes)

            Task.Run(Sub()
                         Thread.Sleep(100)
                         Me.Invoke(Sub()
                                       Dim mensagem = $"Módulo de Configurações{Environment.NewLine}{Environment.NewLine}Funcionalidades planejadas:{Environment.NewLine}• Configuração de conexões Power Query{Environment.NewLine}• Personalização de interface{Environment.NewLine}• Configuração de caminhos de arquivos{Environment.NewLine}• Backup e restauração{Environment.NewLine}• Logs do sistema{Environment.NewLine}{Environment.NewLine}Será implementado em versão futura."

                                       MessageBox.Show(mensagem, "Configurações - Em Desenvolvimento",
                                                     MessageBoxButtons.OK, MessageBoxIcon.Information)

                                       If moduloAtual = "Reposição de Estoque" Then
                                           HighlightButtonMenu(btnReposicaoEstoque)
                                       End If

                                       AtualizarStatus("Pronto")
                                   End Sub)
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.btnConfiguracoes_Click")
        End Try
    End Sub

    Private Sub HighlightButtonMenu(botaoAtivo As Button)
        Try
            If isDisposing OrElse Me.IsDisposed Then Return
            If pnlMenu Is Nothing OrElse pnlMenu.IsDisposed Then Return

            ' Resetar cores de forma otimizada
            For Each control As Control In pnlMenu.Controls
                If TypeOf control Is Button AndAlso control IsNot botaoAtivo AndAlso Not control.IsDisposed Then
                    control.BackColor = Color.Transparent
                    control.ForeColor = Color.Black
                End If
            Next

            ' Destacar botão ativo
            If botaoAtivo IsNot Nothing AndAlso Not botaoAtivo.IsDisposed Then
                With botaoAtivo
                    .BackColor = ConfiguracaoApp.ObterCorHeader()
                    .ForeColor = Color.White
                End With
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.HighlightButtonMenu")
        End Try
    End Sub

    Private Sub MainForm_KeyDown(sender As Object, e As KeyEventArgs)
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

            If e.Control Then
                Select Case e.KeyCode
                    Case Keys.F5
                        If ucReposicaoEstoque IsNot Nothing Then
                            AtualizarStatus("Atualizando dados via atalho...")
                        End If
                    Case Keys.Q
                        FecharAplicacaoSegura()
                    Case Keys.D1
                        If btnReposicaoEstoque IsNot Nothing AndAlso Not btnReposicaoEstoque.IsDisposed Then
                            btnReposicaoEstoque.PerformClick()
                        End If
                    Case Keys.D2
                        If btnRelatorios IsNot Nothing AndAlso Not btnRelatorios.IsDisposed Then
                            btnRelatorios.PerformClick()
                        End If
                    Case Keys.D3
                        If btnConfiguracoes IsNot Nothing AndAlso Not btnConfiguracoes.IsDisposed Then
                            btnConfiguracoes.PerformClick()
                        End If
                    Case Keys.T
                        If btnTestes IsNot Nothing AndAlso Not btnTestes.IsDisposed Then
                            btnTestes.Visible = Not btnTestes.Visible
                            If btnTestes.Visible Then
                                AtualizarStatus("Modo de desenvolvimento ativado")
                            End If
                        End If

                ' ✅ ADICIONAR ESTAS LINHAS:
                    Case Keys.I
                        VerificarConfiguracaoImagens()
                        e.Handled = True
                    Case Keys.P
                        Dim produto = InputBox("Digite o código do produto para testar:", "Teste de Imagem", "101")
                        If Not String.IsNullOrEmpty(produto) Then
                            TestarImagemProduto(produto)
                        End If
                        e.Handled = True
                End Select
            End If

            If e.KeyCode = Keys.F1 Then
                MostrarAjuda()
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.MainForm_KeyDown")
        End Try
    End Sub

    Private Sub MostrarAjuda()
        Try
            Dim ajuda = $"Atalhos de Teclado:{Environment.NewLine}{Environment.NewLine}Ctrl+1 - Reposição de Estoque{Environment.NewLine}Ctrl+2 - Relatórios{Environment.NewLine}Ctrl+3 - Configurações{Environment.NewLine}Ctrl+F5 - Atualizar dados{Environment.NewLine}Ctrl+Q - Fechar aplicação{Environment.NewLine}F1 - Esta ajuda{Environment.NewLine}{Environment.NewLine}Navegação:{Environment.NewLine}• Use os botões do menu superior para navegar entre módulos{Environment.NewLine}• Clique duplo em produtos para ver detalhes{Environment.NewLine}• Use o filtro para localizar produtos rapidamente"

            MessageBox.Show(ajuda, "Ajuda - Gestão de Estoque Difemaq", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.MostrarAjuda")
        End Try
    End Sub

    Private Sub btnTestes_Click(sender As Object, e As EventArgs)
        Try
            AtualizarStatus("Abrindo módulo de testes...")

            Dim testForm As New TestRunnerForm()
            testForm.ShowDialog()

            AtualizarStatus("Pronto")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.btnTestes_Click")
            MessageBox.Show($"Erro ao abrir testes: {ex.Message}", "Erro",
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            LogErros.RegistrarInfo("MainForm fechado", "MainForm.OnFormClosed")

            If timerStatusBar IsNot Nothing Then
                timerStatusBar.Stop()
                timerStatusBar.Dispose()
                timerStatusBar = Nothing
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.OnFormClosed")
        Finally
            MyBase.OnFormClosed(e)
        End Try
    End Sub

    Public Sub FecharControlado()
        Try
            isClosingControlled = True
            isDisposing = True
            Me.Close()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.FecharControlado")
        End Try
    End Sub

    ' Propriedades otimizadas
    Public ReadOnly Property TemDadosCarregados As Boolean
        Get
            Try
                Return ucReposicaoEstoque IsNot Nothing AndAlso Not ucReposicaoEstoque.IsDisposed
            Catch
                Return False
            End Try
        End Get
    End Property

    Public Function ObterInformacoesSistema() As Dictionary(Of String, String)
        Try
            Dim info As New Dictionary(Of String, String)

            info.Add("Módulo Atual", moduloAtual)
            info.Add("Estado Janela", Me.WindowState.ToString())
            info.Add("Dados Carregados", TemDadosCarregados.ToString())
            info.Add("Uptime", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"))

            Return info

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.ObterInformacoesSistema")
            Return New Dictionary(Of String, String)
        End Try
    End Function

    Public Sub VerificarConfiguracaoImagens()
        Try
            Dim relatorio As New System.Text.StringBuilder()
            relatorio.AppendLine("=== RELATÓRIO DE IMAGENS ===")
            relatorio.AppendLine($"Data: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            relatorio.AppendLine()

            ' 1. Verificar diretório
            relatorio.AppendLine($"📁 Diretório configurado: {ConfiguracaoApp.CAMINHO_IMAGENS}")

            If Not System.IO.Directory.Exists(ConfiguracaoApp.CAMINHO_IMAGENS) Then
                relatorio.AppendLine("❌ ERRO: Diretório não existe!")
                relatorio.AppendLine("💡 SOLUÇÃO: Criar o diretório ou alterar o caminho")
            Else
                relatorio.AppendLine("✅ Diretório existe")

                ' Verificar permissões
                Try
                    Dim testFile = System.IO.Path.Combine(ConfiguracaoApp.CAMINHO_IMAGENS, "test.tmp")
                    System.IO.File.WriteAllText(testFile, "teste")
                    System.IO.File.Delete(testFile)
                    relatorio.AppendLine("✅ Permissões de escrita OK")
                Catch
                    relatorio.AppendLine("⚠️ AVISO: Sem permissão de escrita")
                End Try
            End If

            relatorio.AppendLine()

            ' 2. Verificar extensões suportadas
            relatorio.AppendLine("📸 Extensões suportadas:")
            For Each ext In ConfiguracaoApp.EXTENSOES_IMAGEM
                relatorio.AppendLine($"   • {ext}")
            Next

            ' 3. Verificar tamanho máximo
            Dim tamanhoMB = ConfiguracaoApp.TAMANHO_MAXIMO_IMAGEM / 1024 / 1024
            relatorio.AppendLine($"📏 Tamanho máximo: {tamanhoMB}MB")
            relatorio.AppendLine()

            ' 4. Analisar imagens existentes
            If System.IO.Directory.Exists(ConfiguracaoApp.CAMINHO_IMAGENS) Then
                relatorio.AppendLine("🔍 ANÁLISE DAS IMAGENS:")

                Dim arquivos = System.IO.Directory.GetFiles(ConfiguracaoApp.CAMINHO_IMAGENS)
                relatorio.AppendLine($"📊 Total de arquivos encontrados: {arquivos.Length}")

                Dim imagensValidas As Integer = 0
                Dim imagensInvalidas As Integer = 0
                Dim exemplos As New List(Of String)

                For Each arquivo In arquivos
                    Dim nomeArquivo = System.IO.Path.GetFileName(arquivo)
                    Dim extensao = System.IO.Path.GetExtension(arquivo).ToLower()
                    Dim tamanho = New System.IO.FileInfo(arquivo).Length

                    Dim valida = ConfiguracaoApp.EXTENSOES_IMAGEM.Contains(extensao) AndAlso
                               tamanho <= ConfiguracaoApp.TAMANHO_MAXIMO_IMAGEM

                    If valida Then
                        imagensValidas += 1
                        If exemplos.Count < 5 Then
                            exemplos.Add($"✅ {nomeArquivo} ({(tamanho / 1024):F0}KB)")
                        End If
                    Else
                        imagensInvalidas += 1
                        If exemplos.Count < 5 Then
                            Dim motivo = If(Not ConfiguracaoApp.EXTENSOES_IMAGEM.Contains(extensao),
                                          "extensão não suportada", "arquivo muito grande")
                            exemplos.Add($"❌ {nomeArquivo} ({motivo})")
                        End If
                    End If
                Next

                relatorio.AppendLine($"✅ Imagens válidas: {imagensValidas}")
                relatorio.AppendLine($"❌ Imagens inválidas: {imagensInvalidas}")
                relatorio.AppendLine()

                If exemplos.Count > 0 Then
                    relatorio.AppendLine("📋 Exemplos:")
                    For Each exemplo In exemplos
                        relatorio.AppendLine($"   {exemplo}")
                    Next
                    If arquivos.Length > 5 Then
                        relatorio.AppendLine($"   ... e mais {arquivos.Length - 5} arquivos")
                    End If
                End If
            End If

            ' 5. Exemplo de nomenclatura
            relatorio.AppendLine()
            relatorio.AppendLine("💡 EXEMPLO DE NOMENCLATURA CORRETA:")
            relatorio.AppendLine("   Código do produto: ABC123")
            relatorio.AppendLine("   Nome do arquivo: ABC123.jpg")
            relatorio.AppendLine("   Caminho completo: C:\ImagesEstoque\ABC123.jpg")
            relatorio.AppendLine()
            relatorio.AppendLine("⚠️ IMPORTANTE:")
            relatorio.AppendLine("   • Nome do arquivo = Código EXATO do produto")
            relatorio.AppendLine("   • Diferenciar maiúsculas/minúsculas")
            relatorio.AppendLine("   • Não usar espaços ou caracteres especiais")

            ' Exibir relatório
            MessageBox.Show(relatorio.ToString(), "Verificação de Imagens",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Salvar relatório
            Try
                Dim caminhoRelatorio = System.IO.Path.Combine(ConfiguracaoApp.CAMINHO_LOG,
                                                             $"RelatorioImagens_{DateTime.Now:yyyyMMdd_HHmmss}.txt")
                System.IO.File.WriteAllText(caminhoRelatorio, relatorio.ToString())
                MessageBox.Show($"Relatório salvo em: {caminhoRelatorio}", "Relatório Salvo")
            Catch
                ' Ignorar se não conseguir salvar
            End Try

        Catch ex As Exception
            MessageBox.Show($"Erro na verificação: {ex.Message}", "Erro",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' TESTADOR RÁPIDO - Testar com um produto específico
    Public Sub TestarImagemProduto(codigoProduto As String)
        Try
            Dim imagemEncontrada As Boolean = False
            Dim caminhoEncontrado As String = ""

            For Each extensao In ConfiguracaoApp.EXTENSOES_IMAGEM
                Dim caminho = System.IO.Path.Combine(ConfiguracaoApp.CAMINHO_IMAGENS, $"{codigoProduto}{extensao}")
                If System.IO.File.Exists(caminho) Then
                    imagemEncontrada = True
                    caminhoEncontrado = caminho
                    Exit For
                End If
            Next

            If imagemEncontrada Then
                Dim tamanho = New System.IO.FileInfo(caminhoEncontrado).Length
                Dim tamanhoKB = tamanho / 1024
                MessageBox.Show($"✅ IMAGEM ENCONTRADA!" & vbCrLf & vbCrLf &
                              $"Produto: {codigoProduto}" & vbCrLf &
                              $"Arquivo: {System.IO.Path.GetFileName(caminhoEncontrado)}" & vbCrLf &
                              $"Tamanho: {tamanhoKB:F1}KB" & vbCrLf &
                              $"Caminho: {caminhoEncontrado}", "Teste de Imagem")
            Else
                MessageBox.Show($"❌ IMAGEM NÃO ENCONTRADA!" & vbCrLf & vbCrLf &
                              $"Produto: {codigoProduto}" & vbCrLf &
                              $"Procurado em: {ConfiguracaoApp.CAMINHO_IMAGENS}" & vbCrLf &
                              $"Extensões testadas: {String.Join(", ", ConfiguracaoApp.EXTENSOES_IMAGEM)}", "Teste de Imagem")
            End If

        Catch ex As Exception
            MessageBox.Show($"Erro no teste: {ex.Message}", "Erro")
        End Try
    End Sub

End Class