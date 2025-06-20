Imports Microsoft.Office.Interop.Excel
Imports System.Threading.Tasks

Public Class PowerQueryManagerOtimizado
    Private workbook As Workbook
    Private timeoutSegundos As Integer = 15  ' Reduzido drasticamente
    Private maxTentativas As Integer = 2     ' Reduzido para ser mais rápido
    Private ultimaAtualizacao As DateTime = DateTime.MinValue
    Private Shared ReadOnly cacheValidacao As New Dictionary(Of String, DateTime)

    Public Sub New(wb As Workbook)
        If wb Is Nothing Then
            Throw New ArgumentNullException("workbook", "Workbook não pode ser nulo")
        End If
        workbook = wb
        LogErros.RegistrarInfo("PowerQueryManagerOtimizado inicializado", "PowerQueryManagerOtimizado.New")
    End Sub

    ' NOVO: Verificar se atualização é realmente necessária
    Public Function PrecisaAtualizar() As Boolean
        Try
            ' Só atualizar se passou mais de 10 minutos da última atualização
            If DateTime.Now.Subtract(ultimaAtualizacao).TotalMinutes > 10 Then
                LogErros.RegistrarInfo("Atualização necessária - tempo expirado", "PrecisaAtualizar")
                Return True
            End If

            LogErros.RegistrarInfo("Atualização não necessária - cache válido", "PrecisaAtualizar")
            Return False
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManagerOtimizado.PrecisaAtualizar")
            Return True ' Se der erro, atualizar por segurança
        End Try
    End Function

    ' PRINCIPAL: Atualização condicional e rápida
    Public Async Function AtualizarSeNecessarioAsync() As Task
        Try
            If Not PrecisaAtualizar() Then
                LogErros.RegistrarInfo("Pulando atualização Power Query - cache válido", "AtualizarSeNecessarioAsync")
                Return
            End If

            LogErros.RegistrarInfo("Iniciando atualização Power Query otimizada", "AtualizarSeNecessarioAsync")

            ' Atualizar em background thread
            Await Task.Run(Sub() AtualizarRapido())

            ultimaAtualizacao = DateTime.Now
            LogErros.RegistrarInfo("Atualização Power Query concluída", "AtualizarSeNecessarioAsync")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManagerOtimizado.AtualizarSeNecessarioAsync")
            ' Não fazer throw - deixar a aplicação continuar mesmo se Power Query falhar
            LogErros.RegistrarInfo("Continuando sem atualização Power Query", "AtualizarSeNecessarioAsync")
        End Try
    End Function

    ' OTIMIZADO: Atualização mais rápida e seletiva
    Private Sub AtualizarRapido()
        Try
            Dim startTime As DateTime = DateTime.Now
            LogErros.RegistrarInfo("Iniciando atualização rápida", "AtualizarRapido")

            ' Salvar estado para restaurar depois
            Dim estado = SalvarEstadoAplicacao()

            Try
                ' Atualizar apenas conexões essenciais para o projeto
                Dim conexoesAtualizadas As Integer = 0

                For Each connection As WorkbookConnection In workbook.Connections
                    Try
                        If EhConnectionEssencial(connection) Then
                            LogErros.RegistrarInfo($"Atualizando conexão essencial: {connection.Name}", "AtualizarRapido")
                            connection.Refresh()
                            conexoesAtualizadas += 1
                        End If
                    Catch connEx As Exception
                        LogErros.RegistrarErro(connEx, $"Erro na conexão {connection.Name} - continuando")
                        Continue For ' Continuar mesmo se uma conexão falhar
                    End Try
                Next

                LogErros.RegistrarInfo($"Conexões atualizadas: {conexoesAtualizadas}", "AtualizarRapido")

                ' Aguardar conclusão com timeout muito menor
                AguardarConclusaoRapida()

                Dim tempoTotal = DateTime.Now.Subtract(startTime).TotalSeconds
                LogErros.RegistrarInfo($"Atualização rápida concluída em {tempoTotal:F2}s", "AtualizarRapido")

            Finally
                RestaurarEstadoAplicacao(estado)
            End Try

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManagerOtimizado.AtualizarRapido")
            Throw
        End Try
    End Sub

    ' NOVO: Verificar apenas conexões essenciais para o projeto
    Private Function EhConnectionEssencial(connection As WorkbookConnection) As Boolean
        Try
            ' Lista das tabelas essenciais para o funcionamento da aplicação
            Dim tabelasEssenciais() As String = {
                "tblProdutos", "Produtos", "PRODUTOS",
                "tblEstoque", "tblEstoqueVisao", "Estoque", "ESTOQUE",
                "tblCompras", "Compras", "COMPRAS",
                "tblVendas", "Vendas", "VENDAS"
            }

            Dim nomeConnection = connection.Name.ToUpper()

            For Each tabela In tabelasEssenciais
                If nomeConnection.Contains(tabela.ToUpper()) Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "EhConnectionEssencial")
            Return False
        End Try
    End Function

    ' OTIMIZADO: Aguardar conclusão com timeout muito menor
    Private Sub AguardarConclusaoRapida()
        Try
            Dim startTime As DateTime = DateTime.Now
            Dim timeoutReduzido As Integer = 10 ' Apenas 10 segundos de timeout
            Dim intervalosCheck As Integer = 0
            Dim maxIntervalos As Integer = 200 ' 200 * 50ms = 10 segundos

            Do While workbook.Application.CalculationState <> XlCalculationState.xlDone AndAlso intervalosCheck < maxIntervalos
                System.Threading.Thread.Sleep(50) ' Check mais frequente
                System.Windows.Forms.Application.DoEvents()
                intervalosCheck += 1

                ' Log de progresso apenas se demorar mais que 3 segundos
                If intervalosCheck Mod 60 = 0 Then ' A cada 3 segundos (60 * 50ms)
                    Dim tempoDecorrido = DateTime.Now.Subtract(startTime).TotalSeconds
                    LogErros.RegistrarInfo($"Aguardando conclusão... {tempoDecorrido:F1}s", "AguardarConclusaoRapida")
                End If
            Loop

            If intervalosCheck >= maxIntervalos Then
                LogErros.RegistrarInfo("Timeout atingido - forçando continuação", "AguardarConclusaoRapida")
            Else
                Dim tempoTotal = DateTime.Now.Subtract(startTime).TotalSeconds
                LogErros.RegistrarInfo($"Cálculos concluídos em {tempoTotal:F2}s", "AguardarConclusaoRapida")
            End If

            ' Aguardar um pouquinho mais para garantir estabilidade
            System.Threading.Thread.Sleep(200)
            System.Windows.Forms.Application.DoEvents()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManagerOtimizado.AguardarConclusaoRapida")
        End Try
    End Sub

    ' Métodos auxiliares - mantidos do original mas otimizados
    Private Function SalvarEstadoAplicacao() As Dictionary(Of String, Object)
        Dim estado As New Dictionary(Of String, Object)

        Try
            With workbook.Application
                estado("DisplayAlerts") = .DisplayAlerts
                estado("ScreenUpdating") = .ScreenUpdating
                estado("EnableEvents") = .EnableEvents
                estado("Calculation") = .Calculation

                .DisplayAlerts = False
                .ScreenUpdating = False
                .EnableEvents = False
                .Calculation = XlCalculation.xlCalculationManual
            End With
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManagerOtimizado.SalvarEstadoAplicacao")
        End Try

        Return estado
    End Function

    Private Sub RestaurarEstadoAplicacao(estado As Dictionary(Of String, Object))
        Try
            If estado IsNot Nothing AndAlso workbook.Application IsNot Nothing Then
                With workbook.Application
                    If estado.ContainsKey("DisplayAlerts") Then .DisplayAlerts = CBool(estado("DisplayAlerts"))
                    If estado.ContainsKey("ScreenUpdating") Then .ScreenUpdating = CBool(estado("ScreenUpdating"))
                    If estado.ContainsKey("EnableEvents") Then .EnableEvents = CBool(estado("EnableEvents"))
                    If estado.ContainsKey("Calculation") Then .Calculation = CType(estado("Calculation"), XlCalculation)
                End With
            End If
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManagerOtimizado.RestaurarEstadoAplicacao")
        End Try
    End Sub

    ' Métodos públicos mantidos para compatibilidade
    Public Function ObterTabela(nomeTabela As String) As ListObject
        Try
            If String.IsNullOrEmpty(nomeTabela) Then
                LogErros.RegistrarInfo("Nome da tabela está vazio", "ObterTabela")
                Return Nothing
            End If

            LogErros.RegistrarInfo($"Procurando tabela: {nomeTabela}", "ObterTabela")

            For Each worksheet As Worksheet In workbook.Worksheets
                Try
                    For Each tabela As ListObject In worksheet.ListObjects
                        If tabela.Name.Equals(nomeTabela, StringComparison.OrdinalIgnoreCase) Then
                            LogErros.RegistrarInfo($"Tabela encontrada: {nomeTabela} na planilha {worksheet.Name}", "ObterTabela")
                            Return tabela
                        End If
                    Next
                Catch wsEx As Exception
                    LogErros.RegistrarErro(wsEx, $"Erro na planilha {worksheet.Name}")
                    Continue For
                End Try
            Next

            LogErros.RegistrarInfo($"Tabela não encontrada: {nomeTabela}", "ObterTabela")
            Return Nothing

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"PowerQueryManagerOtimizado.ObterTabela({nomeTabela})")
            Return Nothing
        End Try
    End Function

    Public Function ListarTabelas() As List(Of String)
        Try
            Dim tabelas As New List(Of String)()

            For Each worksheet As Worksheet In workbook.Worksheets
                Try
                    For Each tabela As ListObject In worksheet.ListObjects
                        tabelas.Add($"{tabela.Name} ({worksheet.Name})")
                    Next
                Catch wsEx As Exception
                    LogErros.RegistrarErro(wsEx, $"Erro na planilha {worksheet.Name}")
                    Continue For
                End Try
            Next

            LogErros.RegistrarInfo($"Total de tabelas encontradas: {tabelas.Count}", "ListarTabelas")
            Return tabelas

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManagerOtimizado.ListarTabelas")
            Return New List(Of String)()
        End Try
    End Function

    ' NOVO: Método para forçar atualização (para uso manual)
    Public Async Function ForcarAtualizacaoAsync() As Task
        Try
            LogErros.RegistrarInfo("Forçando atualização Power Query", "ForcarAtualizacaoAsync")
            ultimaAtualizacao = DateTime.MinValue ' Invalidar cache
            Await AtualizarSeNecessarioAsync()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ForcarAtualizacaoAsync")
            Throw
        End Try
    End Function

    ' NOVO: Verificar status da última atualização
    Public Function ObterStatusUltimaAtualizacao() As String
        Try
            If ultimaAtualizacao = DateTime.MinValue Then
                Return "Nunca atualizado"
            End If

            Dim tempoDecorrido = DateTime.Now.Subtract(ultimaAtualizacao)
            Return $"Última atualização: {ultimaAtualizacao:HH:mm:ss} (há {tempoDecorrido.TotalMinutes:F1} minutos)"

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ObterStatusUltimaAtualizacao")
            Return "Erro ao verificar status"
        End Try
    End Function
End Class