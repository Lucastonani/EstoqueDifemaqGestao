Imports System.Drawing
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Imports WinFormsApp = System.Windows.Forms.Application

Public Class ThisWorkbook
    Private mainForm As MainForm
    Private WithEvents appEvents As Microsoft.Office.Interop.Excel.Application
    Private originalWindowState As Microsoft.Office.Interop.Excel.XlWindowState
    Private originalVisible As Boolean
    Private isShuttingDown As Boolean = False
    Private isInitialized As Boolean = False
    Private preloadTimer As System.Windows.Forms.Timer
    Private loadingForm As Form
    Private splashForm As SplashForm


    Private Sub ThisWorkbook_Startup() Handles Me.Startup
        Dim splash As SplashForm = Nothing

        Try
            LogErros.RegistrarInfo("Iniciando aplicação GestãoEstoqueDifemaq", "ThisWorkbook.ThisWorkbook_Startup")

            ' Tentar criar splash - se falhar, continuar sem
            Try
                splash = New SplashForm()
                splash.DefinirVersao("1.0.0")
                splash.Show()
                splash.AtualizarStatus("Iniciando sistema...", 0)
                Application.DoEvents()
                Thread.Sleep(300)
            Catch splashEx As Exception
                LogErros.RegistrarErro(splashEx, "ThisWorkbook.ThisWorkbook_Startup - Erro no splash")
                splash = Nothing ' Continuar sem splash
            End Try

            ' Configurar eventos da aplicação
            If splash IsNot Nothing Then splash.AtualizarStatus("Configurando eventos...", 15)
            appEvents = Me.Application
            Thread.Sleep(200)

            ' Salvar estado original do Excel
            If splash IsNot Nothing Then splash.AtualizarStatus("Salvando configurações...", 30)
            ConfigurarAplicacaoExcelRapido()
            Thread.Sleep(300)

            ' Configurações da aplicação Excel
            If splash IsNot Nothing Then splash.AtualizarStatus("Otimizando Excel...", 50)
            InicializarFormularioPrincipalRapido()
            Thread.Sleep(400)

            ' Verificar se as tabelas necessárias existem
            If splash IsNot Nothing Then splash.AtualizarStatus("Verificando dados...", 70)
            IniciarVerificacoesPosCarregamento()
            Thread.Sleep(500)

            ' Criar e exibir o formulário principal
            If splash IsNot Nothing Then splash.AtualizarStatus("Carregando interface...", 90)
            Thread.Sleep(300)

            ' Finalizar
            If splash IsNot Nothing Then splash.AtualizarStatus("Sistema pronto!", 100)
            Thread.Sleep(500)

            isInitialized = True
            LogErros.RegistrarInfo("Aplicação iniciada com sucesso", "ThisWorkbook.ThisWorkbook_Startup")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.ThisWorkbook_Startup")

            ' Atualizar splash com erro
            If splash IsNot Nothing Then
                Try
                    splash.AtualizarStatus("Erro na inicialização!", 0)
                    Thread.Sleep(1000)
                Catch
                    ' Ignorar erro no splash
                End Try
            End If

            ' Tentar restaurar Excel em caso de erro
            Try
                RestaurarConfiguracoes()
            Catch
                ' Ignorar erros na restauração
            End Try

            MessageBox.Show($"Erro ao inicializar aplicação: {ex.Message}{Environment.NewLine}{Environment.NewLine}Verifique se:{Environment.NewLine}- As consultas Power Query estão funcionando{Environment.NewLine}- As tabelas necessárias existem no workbook{Environment.NewLine}- O diretório de imagens está acessível", "Erro de Inicialização", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Fechar splash de forma segura
            If splash IsNot Nothing Then
                Try
                    splash.FecharSuave()
                Catch
                    Try
                        splash.Close()
                    Catch
                        ' Ignorar erro final
                    End Try
                End Try
            End If
        End Try
    End Sub

    Private Sub MostrarSplashScreen()
        Try
            loadingForm = New Form()
            With loadingForm
                .Text = "Carregando Gestão de Estoque..."
                .Size = New Size(400, 200)
                .StartPosition = FormStartPosition.CenterScreen
                .FormBorderStyle = FormBorderStyle.None
                .BackColor = Color.FromArgb(46, 134, 171)
                .TopMost = True
                .ShowInTaskbar = False
            End With

            Dim lblTitle = New Label()
            With lblTitle
                .Text = "Gestão de Estoque - Difemaq"
                .ForeColor = Color.White
                .Font = New Font("Segoe UI", 16, FontStyle.Bold)
                .Location = New Point(50, 50)
                .Size = New Size(300, 30)
                .TextAlign = ContentAlignment.MiddleCenter
            End With

            Dim lblStatus = New Label()
            With lblStatus
                .Text = "Carregando..."
                .ForeColor = Color.White
                .Font = New Font("Segoe UI", 10)
                .Location = New Point(50, 100)
                .Size = New Size(300, 20)
                .TextAlign = ContentAlignment.MiddleCenter
            End With

            Dim progressBar = New ProgressBar()
            With progressBar
                .Location = New Point(50, 130)
                .Size = New Size(300, 20)
                .Style = ProgressBarStyle.Marquee
                .MarqueeAnimationSpeed = 50
            End With

            loadingForm.Controls.AddRange({lblTitle, lblStatus, progressBar})
            loadingForm.Show()
            WinFormsApp.DoEvents()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.MostrarSplashScreen")
        End Try
    End Sub

    Private Sub FecharSplashScreen()
        Try
            If loadingForm IsNot Nothing AndAlso Not loadingForm.IsDisposed Then
                loadingForm.Close()
                loadingForm.Dispose()
                loadingForm = Nothing
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.FecharSplashScreen")
        End Try
    End Sub

    Private Sub ConfigurarAplicacaoExcelRapido()
        Try
            ' Apenas configurações essenciais e rápidas
            With Me.Application
                .Visible = False
                .WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized
                .DisplayAlerts = False
                .ScreenUpdating = False
                .EnableEvents = True
                .Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual ' Manual para acelerar
            End With

            ' Configurar eventos da aplicação
            appEvents = Me.Application

            ' Salvar estado original
            originalWindowState = Me.Application.WindowState
            originalVisible = Me.Application.Visible

            LogErros.RegistrarInfo("Configurações básicas do Excel aplicadas", "ThisWorkbook.ConfigurarAplicacaoExcelRapido")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.ConfigurarAplicacaoExcelRapido")
        End Try
    End Sub

    Private Sub InicializarFormularioPrincipalRapido()
        Try
            ' Criar formulário sem carregar dados inicialmente
            mainForm = New MainForm()

            With mainForm
                .StartPosition = FormStartPosition.CenterScreen
                .WindowState = FormWindowState.Maximized
                .ShowInTaskbar = True
                .TopMost = False
            End With

            ' Mostrar formulário imediatamente
            mainForm.Show()
            mainForm.BringToFront()
            mainForm.Activate()

            ' Fechar splash screen
            FecharSplashScreen()

            ' Forçar processamento de mensagens
            WinFormsApp.DoEvents()

            LogErros.RegistrarInfo("Formulário principal mostrado rapidamente", "ThisWorkbook.InicializarFormularioPrincipalRapido")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.InicializarFormularioPrincipalRapido")
            FecharSplashScreen()
            Throw
        End Try
    End Sub

    Private Sub IniciarVerificacoesPosCarregamento()
        Try
            ' Usar timer para executar verificações não críticas após 2 segundos
            preloadTimer = New System.Windows.Forms.Timer()
            preloadTimer.Interval = 2000 ' 2 segundos
            AddHandler preloadTimer.Tick, AddressOf PreloadTimer_Tick
            preloadTimer.Start()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.IniciarVerificacoesPosCarregamento")
        End Try
    End Sub

    Private Sub PreloadTimer_Tick(sender As Object, e As EventArgs)
        Try
            ' Parar o timer
            preloadTimer.Stop()
            preloadTimer.Dispose()
            preloadTimer = Nothing

            ' Executar verificações em background
            Task.Run(Sub()
                         Try
                             ' Verificar pré-requisitos em background
                             VerificarPreRequisitosBackground()

                             ' Pré-carregar dados essenciais
                             PreCarregarDadosEssenciais()

                             ' Restaurar cálculo automático
                             Me.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic

                         Catch ex As Exception
                             LogErros.RegistrarErro(ex, "ThisWorkbook.PreloadTimer_Tick.Background")
                         End Try
                     End Sub)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.PreloadTimer_Tick")
        End Try
    End Sub

    Private Sub VerificarPreRequisitosBackground()
        Try
            Dim workbookInterface As Microsoft.Office.Interop.Excel.Workbook = ObterWorkbook()
            If workbookInterface Is Nothing Then Return

            ' Verificar rapidamente se as tabelas principais existem
            Dim powerQueryManager As New PowerQueryManager(workbookInterface)
            Dim tabelasEncontradas As List(Of String) = powerQueryManager.ListarTabelas()

            ' Verificar diretório de imagens
            If Not System.IO.Directory.Exists(ConfiguracaoApp.CAMINHO_IMAGENS) Then
                System.IO.Directory.CreateDirectory(ConfiguracaoApp.CAMINHO_IMAGENS)
            End If

            LogErros.RegistrarInfo("Verificações de background concluídas", "ThisWorkbook.VerificarPreRequisitosBackground")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.VerificarPreRequisitosBackground")
        End Try
    End Sub

    Private Sub PreCarregarDadosEssenciais()
        Try
            ' Informar o MainForm que pode carregar dados quando necessário
            If mainForm IsNot Nothing AndAlso Not mainForm.IsDisposed Then
                mainForm.Invoke(Sub()
                                    mainForm.AtualizarStatus("Sistema pronto - Dados carregados sob demanda")
                                End Sub)
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.PreCarregarDadosEssenciais")
        End Try
    End Sub

    ' Resto dos métodos permanecem iguais...
    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown
        Try
            isShuttingDown = True
            LogErros.RegistrarInfo("Iniciando shutdown da aplicação", "ThisWorkbook.ThisWorkbook_Shutdown")

            ' Limpar timer se ainda estiver ativo
            If preloadTimer IsNot Nothing Then
                preloadTimer.Stop()
                preloadTimer.Dispose()
                preloadTimer = Nothing
            End If

            ' Fechar splash screen se ainda estiver aberto
            FecharSplashScreen()

            ' Fechar formulário principal
            If mainForm IsNot Nothing AndAlso Not mainForm.IsDisposed Then
                mainForm.Close()
                mainForm.Dispose()
                mainForm = Nothing
            End If

            ' Restaurar configurações do Excel
            RestaurarConfiguracoes()

            LogErros.RegistrarInfo("Shutdown concluído", "ThisWorkbook.ThisWorkbook_Shutdown")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.ThisWorkbook_Shutdown")
        End Try
    End Sub

    Private Sub RestaurarConfiguracoes()
        Try
            If Me.Application IsNot Nothing AndAlso Not isShuttingDown Then
                With Me.Application
                    .Visible = originalVisible
                    .WindowState = originalWindowState
                    .DisplayAlerts = True
                    .ScreenUpdating = True
                    .EnableEvents = True
                    .Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic
                    .DisplayFullScreen = False
                    .DisplayFormulaBar = True
                    .DisplayStatusBar = True
                End With

                LogErros.RegistrarInfo("Configurações do Excel restauradas", "ThisWorkbook.RestaurarConfiguracoes")
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.RestaurarConfiguracoes")
        End Try
    End Sub

    ' Eventos da aplicação permanecem iguais...
    Private Sub appEvents_WorkbookBeforeClose(Wb As Microsoft.Office.Interop.Excel.Workbook, ByRef Cancel As Boolean) Handles appEvents.WorkbookBeforeClose
        Try
            Dim thisWorkbookInterface As Microsoft.Office.Interop.Excel.Workbook = ObterWorkbook()

            If thisWorkbookInterface IsNot Nothing AndAlso Wb Is thisWorkbookInterface AndAlso Not isShuttingDown Then
                LogErros.RegistrarInfo("Evento WorkbookBeforeClose disparado", "ThisWorkbook.appEvents_WorkbookBeforeClose")

                If mainForm IsNot Nothing AndAlso Not mainForm.IsDisposed Then
                    Cancel = True
                    mainForm.Close()
                End If
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.appEvents_WorkbookBeforeClose")
        End Try
    End Sub

    Private Sub appEvents_WorkbookBeforeSave(Wb As Microsoft.Office.Interop.Excel.Workbook, SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles appEvents.WorkbookBeforeSave
        Try
            Dim thisWorkbookInterface As Microsoft.Office.Interop.Excel.Workbook = ObterWorkbook()

            If thisWorkbookInterface IsNot Nothing AndAlso Wb Is thisWorkbookInterface Then
                LogErros.RegistrarInfo($"Workbook sendo salvo (SaveAsUI: {SaveAsUI})", "ThisWorkbook.appEvents_WorkbookBeforeSave")
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.appEvents_WorkbookBeforeSave")
        End Try
    End Sub

    ' Métodos públicos permanecem iguais...
    Public Sub FecharAplicacao()
        Try
            LogErros.RegistrarInfo("Fechamento controlado da aplicação solicitado", "ThisWorkbook.FecharAplicacao")
            isShuttingDown = True

            Dim thisWorkbookInterface As Microsoft.Office.Interop.Excel.Workbook = ObterWorkbook()
            If thisWorkbookInterface IsNot Nothing Then
                thisWorkbookInterface.Close(False)
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.FecharAplicacao")
        End Try
    End Sub

    Public Function ObterWorkbook() As Microsoft.Office.Interop.Excel.Workbook
        Try
            If isShuttingDown Then Return Nothing
            If Me.Application Is Nothing Then Return Nothing

            Try
                Dim wb = DirectCast(Me.InnerObject, Microsoft.Office.Interop.Excel.Workbook)
                If wb IsNot Nothing Then Return wb
            Catch
            End Try

            Try
                If Me.Application.Workbooks.Count > 0 Then
                    For Each wb As Microsoft.Office.Interop.Excel.Workbook In Me.Application.Workbooks
                        If wb.Name = Me.Name Then Return wb
                    Next
                End If
            Catch
            End Try

            Return Nothing

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.ObterWorkbook")
            Return Nothing
        End Try
    End Function

    ' Propriedades permanecem iguais...
    Public ReadOnly Property FormularioPrincipal As MainForm
        Get
            Return mainForm
        End Get
    End Property

    Public ReadOnly Property AplicacaoAtiva As Boolean
        Get
            Return mainForm IsNot Nothing AndAlso Not mainForm.IsDisposed AndAlso mainForm.Visible
        End Get
    End Property

    Public ReadOnly Property EstaInicializado As Boolean
        Get
            Return isInitialized AndAlso Not isShuttingDown
        End Get
    End Property

    Public Sub ExecutarTestes()
        UnitTests.RunAllTests()
    End Sub

    Public Sub VerificarSistema()
        DeploymentChecklist.RunAllChecks()
    End Sub

End Class