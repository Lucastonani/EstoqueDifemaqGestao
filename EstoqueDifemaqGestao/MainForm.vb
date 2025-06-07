Imports System.Drawing
Imports System.Windows.Forms

Public Class MainForm
    Private ucReposicaoEstoque As UcReposicaoEstoque
    Private timerStatusBar As Timer
    Private moduloAtual As String = ""
    Private isClosingControlled As Boolean = False
    Private isDisposing As Boolean = False

    Public Sub New()
        InitializeComponent()
        ConfigurarFormulario()
        InicializarStatusBar()
        CarregarUserControlInicial()
    End Sub

    Private Sub ConfigurarFormulario()
        Try
            ' Configurações do formulário
            Me.Text = "Gestão de Estoque - Difemaq"
            Me.Size = New Size(1400, 900)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.WindowState = FormWindowState.Maximized
            Me.MinimumSize = New Size(1200, 800)

            ' Configurações de exibição
            Me.ShowInTaskbar = True
            Me.KeyPreview = True

            ' Configurar evento de teclas para atalhos
            AddHandler Me.KeyDown, AddressOf MainForm_KeyDown

            LogErros.RegistrarInfo("Formulário principal configurado", "MainForm.ConfigurarFormulario")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.ConfigurarFormulario")
        End Try
    End Sub

    Private Sub InicializarStatusBar()
        Try
            ' Configurar timer para atualizar data/hora
            timerStatusBar = New Timer()
            timerStatusBar.Interval = 1000 ' 1 segundo
            AddHandler timerStatusBar.Tick, AddressOf TimerStatusBar_Tick
            timerStatusBar.Start()

            ' Atualizar status inicial
            AtualizarStatus("Sistema iniciado - Pronto para uso")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.InicializarStatusBar")
        End Try
    End Sub

    Private Sub CarregarUserControlInicial()
        Try
            ' Carregar UserControl de Reposição de Estoque como padrão
            CarregarUserControl(GetType(UcReposicaoEstoque))
            HighlightButtonMenu(btnReposicaoEstoque)
            moduloAtual = "Reposição de Estoque"

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.CarregarUserControlInicial")
            MessageBox.Show(String.Format("Erro ao carregar interface inicial: {0}", ex.Message), "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub CarregarUserControl(tipoUserControl As Type)
        Try
            ' Verificar se está em processo de fechamento
            If isDisposing OrElse Me.IsDisposed Then Return

            ' Mostrar indicador de carregamento
            AtualizarStatus("Carregando módulo...")
            Me.Cursor = Cursors.WaitCursor

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

            ' Forçar redesenho
            pnlConteudo.Refresh()
            Me.Refresh()

            LogErros.RegistrarInfo(String.Format("UserControl carregado: {0}", tipoUserControl.Name), "MainForm.CarregarUserControl")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, String.Format("MainForm.CarregarUserControl({0})", tipoUserControl.Name))
            MessageBox.Show(String.Format("Erro ao carregar módulo: {0}", ex.Message), "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            AtualizarStatus("Pronto")
        End Try
    End Sub

    Private Sub LimparPainelConteudo()
        Try
            ' Verificar se o painel ainda existe
            If pnlConteudo Is Nothing OrElse pnlConteudo.IsDisposed Then Return

            ' Limpar controles existentes
            For Each control As Control In pnlConteudo.Controls.Cast(Of Control).ToArray()
                If TypeOf control Is UserControl Then
                    control.Dispose()
                End If
                pnlConteudo.Controls.Remove(control)
            Next

            ' Limpar referências
            ucReposicaoEstoque = Nothing

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.LimparPainelConteudo")
        End Try
    End Sub

    Private Sub MainForm_KeyDown(sender As Object, e As KeyEventArgs)
        Try
            ' Verificar se está em processo de fechamento
            If isDisposing OrElse Me.IsDisposed Then Return

            ' Atalhos de teclado
            If e.Control Then
                Select Case e.KeyCode
                    Case Keys.F5
                        ' Ctrl+F5 - Atualizar dados
                        If ucReposicaoEstoque IsNot Nothing Then
                            AtualizarStatus("Atualizando dados via atalho...")
                        End If
                    Case Keys.Q
                        ' Ctrl+Q - Fechar aplicação
                        FecharAplicacaoSegura()
                    Case Keys.D1
                        ' Ctrl+1 - Ir para Reposição de Estoque
                        If btnReposicaoEstoque IsNot Nothing AndAlso Not btnReposicaoEstoque.IsDisposed Then
                            btnReposicaoEstoque.PerformClick()
                        End If
                    Case Keys.D2
                        ' Ctrl+2 - Ir para Relatórios
                        If btnRelatorios IsNot Nothing AndAlso Not btnRelatorios.IsDisposed Then
                            btnRelatorios.PerformClick()
                        End If
                    Case Keys.D3
                        ' Ctrl+3 - Ir para Configurações
                        If btnConfiguracoes IsNot Nothing AndAlso Not btnConfiguracoes.IsDisposed Then
                            btnConfiguracoes.PerformClick()
                        End If
                End Select
            End If

            ' F1 - Ajuda
            If e.KeyCode = Keys.F1 Then
                MostrarAjuda()
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.MainForm_KeyDown")
        End Try
    End Sub

    Private Sub MostrarAjuda()
        Try
            Dim ajuda As String = String.Format("Atalhos de Teclado:{0}{0}Ctrl+1 - Reposição de Estoque{0}Ctrl+2 - Relatórios{0}Ctrl+3 - Configurações{0}Ctrl+F5 - Atualizar dados{0}Ctrl+Q - Fechar aplicação{0}F1 - Esta ajuda{0}{0}Navegação:{0}• Use os botões do menu superior para navegar entre módulos{0}• Clique duplo em produtos para ver detalhes{0}• Use o filtro para localizar produtos rapidamente", Environment.NewLine)

            MessageBox.Show(ajuda, "Ajuda - Gestão de Estoque Difemaq", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.MostrarAjuda")
        End Try
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            ' Marcar que está em processo de fechamento
            isDisposing = True

            If Not isClosingControlled Then
                ' Confirmar fechamento apenas se não foi controlado
                Dim resultado As DialogResult = MessageBox.Show(
                    "Deseja realmente fechar o sistema?",
                    "Confirmar Fechamento",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2)

                If resultado = DialogResult.No Then
                    isDisposing = False ' Cancelar processo de fechamento
                    e.Cancel = True
                    Return
                End If
            End If

            ' Limpar recursos
            LimparRecursos()

            ' Fechar workbook Excel de forma controlada
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

    Public Sub FecharControlado()
        ' Método para fechar o formulário sem confirmação
        Try
            isClosingControlled = True
            isDisposing = True
            Me.Close()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.FecharControlado")
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

    Private Sub LimparRecursos()
        Try
            LogErros.RegistrarInfo("Limpando recursos do MainForm", "MainForm.LimparRecursos")

            ' Parar timer
            If timerStatusBar IsNot Nothing Then
                timerStatusBar.Stop()
                timerStatusBar.Dispose()
                timerStatusBar = Nothing
            End If

            ' Limpar UserControls
            LimparPainelConteudo()

            LogErros.RegistrarInfo("Recursos do MainForm limpos", "MainForm.LimparRecursos")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.LimparRecursos")
        End Try
    End Sub

    Private Sub TimerStatusBar_Tick(sender As Object, e As EventArgs)
        Try
            ' Verificar se está em processo de fechamento
            If isDisposing OrElse Me.IsDisposed Then Return

            ' Verificar se o label ainda existe
            If lblDataHora IsNot Nothing AndAlso Not lblDataHora.IsDisposed Then
                lblDataHora.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.TimerStatusBar_Tick")
        End Try
    End Sub

    Public Sub AtualizarStatus(mensagem As String)
        Try
            ' Verificar se está em processo de fechamento
            If isDisposing OrElse Me.IsDisposed Then Return

            ' Verificar se os controles ainda existem
            If lblStatus IsNot Nothing AndAlso Not lblStatus.IsDisposed Then
                lblStatus.Text = mensagem
            End If

            If StatusStrip IsNot Nothing AndAlso Not StatusStrip.IsDisposed Then
                StatusStrip.Refresh()
            End If

            System.Windows.Forms.Application.DoEvents()

            LogErros.RegistrarInfo(String.Format("Status atualizado: {0}", mensagem), "MainForm.AtualizarStatus")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.AtualizarStatus")
        End Try
    End Sub

    Private Sub btnReposicaoEstoque_Click(sender As Object, e As EventArgs) Handles btnReposicaoEstoque.Click
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

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

    Private Sub btnRelatorios_Click(sender As Object, e As EventArgs) Handles btnRelatorios.Click
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

            AtualizarStatus("Módulo de Relatórios em desenvolvimento...")
            HighlightButtonMenu(btnRelatorios)

            Dim mensagem As String = String.Format("Módulo de Relatórios{0}{0}Funcionalidades planejadas:{0}• Relatório de estoque baixo{0}• Análise de vendas por período{0}• Histórico de compras{0}• Gráficos de movimentação{0}• Exportação para Excel/PDF{0}{0}Será implementado em versão futura.", Environment.NewLine)

            MessageBox.Show(mensagem, "Relatórios - Em Desenvolvimento", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Voltar para o módulo anterior
            If moduloAtual = "Reposição de Estoque" Then
                HighlightButtonMenu(btnReposicaoEstoque)
            End If

            AtualizarStatus("Pronto")
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.btnRelatorios_Click")
        End Try
    End Sub

    Private Sub btnConfiguracoes_Click(sender As Object, e As EventArgs) Handles btnConfiguracoes.Click
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

            AtualizarStatus("Módulo de Configurações em desenvolvimento...")
            HighlightButtonMenu(btnConfiguracoes)

            Dim mensagem As String = String.Format("Módulo de Configurações{0}{0}Funcionalidades planejadas:{0}• Configuração de conexões Power Query{0}• Personalização de interface{0}• Configuração de caminhos de arquivos{0}• Backup e restauração{0}• Logs do sistema{0}{0}Será implementado em versão futura.", Environment.NewLine)

            MessageBox.Show(mensagem, "Configurações - Em Desenvolvimento", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Voltar para o módulo anterior
            If moduloAtual = "Reposição de Estoque" Then
                HighlightButtonMenu(btnReposicaoEstoque)
            End If

            AtualizarStatus("Pronto")
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "MainForm.btnConfiguracoes_Click")
        End Try
    End Sub

    Private Sub HighlightButtonMenu(botaoAtivo As Button)
        Try
            If isDisposing OrElse Me.IsDisposed Then Return
            If pnlMenu Is Nothing OrElse pnlMenu.IsDisposed Then Return

            ' Resetar cores de todos os botões
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

    Private Sub btnMinimizar_Click(sender As Object, e As EventArgs) Handles btnMinimizar.Click
        Try
            If isDisposing OrElse Me.IsDisposed Then Return

            Me.WindowState = FormWindowState.Minimized
            LogErros.RegistrarInfo("Janela minimizada", "MainForm.btnMinimizar_Click")
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
                LogErros.RegistrarInfo("Janela restaurada", "MainForm.btnMaximizar_Click")
            Else
                Me.WindowState = FormWindowState.Maximized
                btnMaximizar.Text = "❐"
                LogErros.RegistrarInfo("Janela maximizada", "MainForm.btnMaximizar_Click")
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

    ' Propriedade para verificar se há dados carregados
    Public ReadOnly Property TemDadosCarregados As Boolean
        Get
            Try
                Return ucReposicaoEstoque IsNot Nothing AndAlso Not ucReposicaoEstoque.IsDisposed
            Catch
                Return False
            End Try
        End Get
    End Property

    ' Método para obter informações do sistema
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



End Class