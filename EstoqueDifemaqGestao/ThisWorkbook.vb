Imports System.Windows.Forms

Public Class ThisWorkbook
    Private mainForm As MainForm
    Private WithEvents appEvents As Microsoft.Office.Interop.Excel.Application
    Private originalWindowState As Microsoft.Office.Interop.Excel.XlWindowState
    Private originalVisible As Boolean
    Private isShuttingDown As Boolean = False
    Private isInitialized As Boolean = False

    Private Sub ThisWorkbook_Startup() Handles Me.Startup
        Try
            LogErros.RegistrarInfo("Iniciando aplicação GestãoEstoqueDifemaq", "ThisWorkbook.ThisWorkbook_Startup")

            ' Configurar eventos da aplicação
            appEvents = Me.Application

            ' Salvar estado original do Excel
            SalvarEstadoOriginal()

            ' Configurações da aplicação Excel
            ConfigurarAplicacaoExcel()

            ' Verificar se as tabelas necessárias existem
            VerificarPreRequisitos()

            ' Criar e exibir o formulário principal
            InicializarFormularioPrincipal()

            isInitialized = True
            LogErros.RegistrarInfo("Aplicação iniciada com sucesso", "ThisWorkbook.ThisWorkbook_Startup")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.ThisWorkbook_Startup")

            ' Tentar restaurar Excel em caso de erro
            Try
                RestaurarConfiguracoes()
            Catch
                ' Ignorar erros na restauração
            End Try

            MessageBox.Show(String.Format("Erro ao inicializar aplicação: {0}{1}{1}Verifique se:{1}- As consultas Power Query estão funcionando{1}- As tabelas necessárias existem no workbook{1}- O diretório de imagens está acessível", ex.Message, Environment.NewLine), "Erro de Inicialização", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SalvarEstadoOriginal()
        Try
            With Me.Application
                originalWindowState = .WindowState
                originalVisible = .Visible
            End With
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.SalvarEstadoOriginal")
        End Try
    End Sub

    Private Sub VerificarPreRequisitos()
        Try
            ' Usar a interface Workbook em vez de casting direto
            Dim workbookInterface As Microsoft.Office.Interop.Excel.Workbook = ObterWorkbook()
            If workbookInterface Is Nothing Then
                LogErros.RegistrarInfo("Não foi possível obter interface do workbook", "ThisWorkbook.VerificarPreRequisitos")
                Return
            End If

            Dim powerQueryManager As New PowerQueryManager(workbookInterface)

            Dim tabelasNecessarias As String() = {
                ConfiguracaoApp.TABELA_PRODUTOS,
                ConfiguracaoApp.TABELA_ESTOQUE,
                ConfiguracaoApp.TABELA_COMPRAS,
                ConfiguracaoApp.TABELA_VENDAS
            }

            Dim tabelasEncontradas As List(Of String) = powerQueryManager.ListarTabelas()
            Dim tabelasFaltando As New List(Of String)()

            For Each tabelaNecessaria As String In tabelasNecessarias
                Dim encontrada As Boolean = False
                For Each tabelaEncontrada As String In tabelasEncontradas
                    If tabelaEncontrada.Contains(tabelaNecessaria) Then
                        encontrada = True
                        Exit For
                    End If
                Next

                If Not encontrada Then
                    tabelasFaltando.Add(tabelaNecessaria)
                End If
            Next

            If tabelasFaltando.Count > 0 Then
                Dim mensagem As String = String.Format("As seguintes tabelas não foram encontradas:{0}{1}{0}{0}Verifique se as consultas Power Query estão configuradas corretamente.", Environment.NewLine, String.Join(Environment.NewLine, tabelasFaltando))

                LogErros.RegistrarInfo(String.Format("Tabelas faltando: {0}", String.Join(", ", tabelasFaltando)), "ThisWorkbook.VerificarPreRequisitos")

                MessageBox.Show(mensagem, "Aviso - Tabelas Não Encontradas", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                LogErros.RegistrarInfo("Todas as tabelas necessárias foram encontradas", "ThisWorkbook.VerificarPreRequisitos")
            End If

            ' Verificar diretório de imagens
            If Not System.IO.Directory.Exists(ConfiguracaoApp.CAMINHO_IMAGENS) Then
                Try
                    System.IO.Directory.CreateDirectory(ConfiguracaoApp.CAMINHO_IMAGENS)
                    LogErros.RegistrarInfo(String.Format("Diretório de imagens criado: {0}", ConfiguracaoApp.CAMINHO_IMAGENS), "ThisWorkbook.VerificarPreRequisitos")
                Catch dirEx As Exception
                    LogErros.RegistrarErro(dirEx, "ThisWorkbook.VerificarPreRequisitos - Erro ao criar diretório de imagens")
                    MessageBox.Show(String.Format("Não foi possível criar o diretório de imagens em:{0}{1}{0}{0}As imagens dos produtos não serão exibidas.", Environment.NewLine, ConfiguracaoApp.CAMINHO_IMAGENS), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.VerificarPreRequisitos")
        End Try
    End Sub

    Private Sub ConfigurarAplicacaoExcel()
        Try
            With Me.Application
                .Visible = False
                .WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized
                .DisplayAlerts = False
                .ScreenUpdating = False
                .EnableEvents = True ' Manter eventos habilitados para funcionalidade
                .Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic ' Manter automático para Power Query

                ' Configurações adicionais de interface
                .DisplayFullScreen = False
                .DisplayFormulaBar = False
                .DisplayStatusBar = True
            End With

            LogErros.RegistrarInfo("Configurações do Excel aplicadas", "ThisWorkbook.ConfigurarAplicacaoExcel")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.ConfigurarAplicacaoExcel")
        End Try
    End Sub

    Private Sub InicializarFormularioPrincipal()
        Try
            ' Garantir que estamos no thread principal
            If System.Windows.Forms.Application.OpenForms.Count = 0 OrElse mainForm Is Nothing OrElse mainForm.IsDisposed Then
                mainForm = New MainForm()
            End If

            ' Configurar o formulário
            With mainForm
                .StartPosition = FormStartPosition.CenterScreen
                .WindowState = FormWindowState.Maximized
                .ShowInTaskbar = True
                .TopMost = False
            End With

            ' Mostrar o formulário
            mainForm.Show()

            ' Garantir que o formulário está visível e em foco
            mainForm.BringToFront()
            mainForm.Activate()
            mainForm.Focus()

            ' Forçar o Windows Forms a processar mensagens
            System.Windows.Forms.Application.DoEvents()

            LogErros.RegistrarInfo("Formulário principal inicializado", "ThisWorkbook.InicializarFormularioPrincipal")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.InicializarFormularioPrincipal")
            Throw
        End Try
    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown
        Try
            isShuttingDown = True
            LogErros.RegistrarInfo("Iniciando shutdown da aplicação", "ThisWorkbook.ThisWorkbook_Shutdown")

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

    Private Sub appEvents_WorkbookBeforeClose(Wb As Microsoft.Office.Interop.Excel.Workbook, ByRef Cancel As Boolean) Handles appEvents.WorkbookBeforeClose
        Try
            ' Comparar usando InnerObject para evitar problemas de casting
            Dim thisWorkbookInterface As Microsoft.Office.Interop.Excel.Workbook = ObterWorkbook()

            If thisWorkbookInterface IsNot Nothing AndAlso Wb Is thisWorkbookInterface AndAlso Not isShuttingDown Then
                LogErros.RegistrarInfo("Evento WorkbookBeforeClose disparado", "ThisWorkbook.appEvents_WorkbookBeforeClose")

                ' Fechar aplicação quando workbook for fechado
                If mainForm IsNot Nothing AndAlso Not mainForm.IsDisposed Then
                    ' Cancelar o fechamento do workbook temporariamente
                    Cancel = True

                    ' Fechar o formulário principal (que vai fechar o workbook corretamente)
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
                ' Permitir salvamento automático, mas logar a ação
                LogErros.RegistrarInfo(String.Format("Workbook sendo salvo (SaveAsUI: {0})", SaveAsUI), "ThisWorkbook.appEvents_WorkbookBeforeSave")
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.appEvents_WorkbookBeforeSave")
        End Try
    End Sub

    Private Sub appEvents_WorkbookOpen(Wb As Microsoft.Office.Interop.Excel.Workbook) Handles appEvents.WorkbookOpen
        Try
            Dim thisWorkbookInterface As Microsoft.Office.Interop.Excel.Workbook = ObterWorkbook()

            If thisWorkbookInterface IsNot Nothing AndAlso Wb IsNot thisWorkbookInterface Then
                LogErros.RegistrarInfo(String.Format("Outro workbook foi aberto: {0}", Wb.Name), "ThisWorkbook.appEvents_WorkbookOpen")
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.appEvents_WorkbookOpen")
        End Try
    End Sub

    Private Sub appEvents_SheetCalculate(Sh As Object) Handles appEvents.SheetCalculate
        Try
            ' Log apenas se necessário para debug, pode gerar muitas entradas
            ' LogErros.RegistrarInfo(String.Format("Planilha calculada: {0}", Sh.Name), "ThisWorkbook.appEvents_SheetCalculate")
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.appEvents_SheetCalculate")
        End Try
    End Sub

    ' Método público para fechar a aplicação de forma controlada
    Public Sub FecharAplicacao()
        Try
            LogErros.RegistrarInfo("Fechamento controlado da aplicação solicitado", "ThisWorkbook.FecharAplicacao")
            isShuttingDown = True

            ' Fechar o workbook usando InnerObject
            Dim thisWorkbookInterface As Microsoft.Office.Interop.Excel.Workbook = ObterWorkbook()
            If thisWorkbookInterface IsNot Nothing Then
                thisWorkbookInterface.Close(False)
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.FecharAplicacao")
        End Try
    End Sub

    ' Método para obter referência do formulário principal (se necessário)
    Public ReadOnly Property FormularioPrincipal As MainForm
        Get
            Return mainForm
        End Get
    End Property

    ' Método para verificar se a aplicação está rodando
    Public ReadOnly Property AplicacaoAtiva As Boolean
        Get
            Return mainForm IsNot Nothing AndAlso Not mainForm.IsDisposed AndAlso mainForm.Visible
        End Get
    End Property

    ' Propriedade para verificar se foi inicializado
    Public ReadOnly Property EstaInicializado As Boolean
        Get
            Return isInitialized AndAlso Not isShuttingDown
        End Get
    End Property

    ' Método público para obter a interface Workbook corretamente
    Public Function ObterWorkbook() As Microsoft.Office.Interop.Excel.Workbook
        Try
            ' Tentar múltiplas formas de obter o workbook
            Dim workbook As Microsoft.Office.Interop.Excel.Workbook = Nothing

            ' 1ª tentativa: DirectCast com InnerObject
            Try
                workbook = DirectCast(Me.InnerObject, Microsoft.Office.Interop.Excel.Workbook)
                If workbook IsNot Nothing Then Return workbook
            Catch
                ' Continuar para próxima tentativa
            End Try

            ' 2ª tentativa: Através da aplicação
            Try
                If Me.Application IsNot Nothing Then
                    workbook = Me.Application.ActiveWorkbook
                    If workbook IsNot Nothing Then Return workbook
                End If
            Catch
                ' Continuar para próxima tentativa
            End Try

            ' 3ª tentativa: Tentar cast direto (pode falhar)
            Try
                workbook = CType(Me, Microsoft.Office.Interop.Excel.Workbook)
            Catch
                ' Última tentativa falhou
            End Try

            Return workbook

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ThisWorkbook.ObterWorkbook")
            Return Nothing
        End Try
    End Function

End Class