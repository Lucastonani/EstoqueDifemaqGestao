Imports Microsoft.Office.Interop.Excel
Imports System.Threading
Imports WinFormsApp = System.Windows.Forms.Application

Public Class PowerQueryManager
    Private workbook As Workbook
    Private timeoutSegundos As Integer
    Private maxTentativas As Integer = 2 ' Reduzido de 3 para 2

    ' Cache para otimizar acesso às tabelas
    Private tabelasCache As New Dictionary(Of String, ListObject)
    Private cacheValido As DateTime = DateTime.MinValue
    Private Const CACHE_TIMEOUT_MINUTES As Integer = 5

    Public Sub New(wb As Workbook)
        If wb Is Nothing Then
            Throw New ArgumentNullException("workbook", "Workbook não pode ser nulo")
        End If
        workbook = wb
        timeoutSegundos = Math.Min(ConfiguracaoApp.TIMEOUT_POWERQUERY, 30) ' Máximo 30 segundos
    End Sub

    Public Sub AtualizarTodasConsultas()
        Try
            Dim startTime As DateTime = DateTime.Now
            Dim tentativa As Integer = 1

            LogErros.RegistrarInfo("Iniciando atualização otimizada de consultas Power Query", "PowerQueryManager.AtualizarTodasConsultas")

            ' Configurações otimizadas do Excel
            Dim estadoAnterior = ConfigurarExcelParaAtualizacao()

            Try
                While tentativa <= maxTentativas
                    Try
                        LogErros.RegistrarInfo($"Tentativa {tentativa} de {maxTentativas}", "PowerQueryManager.AtualizarTodasConsultas")

                        ' Atualizar de forma otimizada
                        AtualizarConexoesPowerQueryOtimizado()

                        ' Aguardar com timeout reduzido
                        AguardarConclusaoAtualizacaoOtimizada(startTime)

                        LogErros.RegistrarInfo($"Consultas atualizadas com sucesso na tentativa {tentativa}", "PowerQueryManager.AtualizarTodasConsultas")

                        ' Invalidar cache após atualização bem-sucedida
                        InvalidarCache()
                        Exit While

                    Catch ex As TimeoutException
                        LogErros.RegistrarErro(ex, $"PowerQueryManager.AtualizarTodasConsultas - Timeout na tentativa {tentativa}")
                        tentativa += 1
                        If tentativa <= maxTentativas Then
                            Thread.Sleep(1000) ' Reduzido de 2000ms para 1000ms
                            startTime = DateTime.Now
                        End If

                    Catch ex As Exception
                        LogErros.RegistrarErro(ex, $"PowerQueryManager.AtualizarTodasConsultas - Erro na tentativa {tentativa}")
                        tentativa += 1
                        If tentativa <= maxTentativas Then
                            Thread.Sleep(500) ' Reduzido de 1000ms para 500ms
                            startTime = DateTime.Now
                        End If
                    End Try
                End While

                If tentativa > maxTentativas Then
                    Throw New Exception($"Falha ao atualizar consultas após {maxTentativas} tentativas")
                End If

            Finally
                RestaurarConfiguracoesExcel(estadoAnterior)
            End Try

            LogErros.RegistrarInfo("Atualização otimizada de consultas Power Query concluída", "PowerQueryManager.AtualizarTodasConsultas")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.AtualizarTodasConsultas")
            Throw New Exception($"Erro ao atualizar consultas Power Query: {ex.Message}")
        End Try
    End Sub

    Private Function ConfigurarExcelParaAtualizacao() As Dictionary(Of String, Object)
        Dim estado As New Dictionary(Of String, Object)

        Try
            With workbook.Application
                estado("DisplayAlerts") = .DisplayAlerts
                estado("ScreenUpdating") = .ScreenUpdating
                estado("EnableEvents") = .EnableEvents
                estado("Calculation") = .Calculation

                ' Configurações otimizadas para atualização
                .DisplayAlerts = False
                .ScreenUpdating = False
                .EnableEvents = False
                .Calculation = XlCalculation.xlCalculationManual
            End With
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ConfigurarExcelParaAtualizacao")
        End Try

        Return estado
    End Function

    Private Sub RestaurarConfiguracoesExcel(estado As Dictionary(Of String, Object))
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
            LogErros.RegistrarErro(ex, "PowerQueryManager.RestaurarConfiguracoesExcel")
        End Try
    End Sub

    Private Sub AtualizarConexoesPowerQueryOtimizado()
        Try
            Dim conexoesAtualizadas As Integer = 0
            Dim conexoesComErro As Integer = 0

            ' Atualizar apenas conexões Power Query essenciais
            For Each connection As WorkbookConnection In workbook.Connections
                Try
                    If EhConnectionPowerQueryEssencial(connection) Then
                        LogErros.RegistrarInfo($"Atualizando conexão essencial: {connection.Name}", "PowerQueryManager.AtualizarConexoesPowerQueryOtimizado")
                        connection.Refresh()
                        conexoesAtualizadas += 1
                    End If
                Catch connEx As Exception
                    conexoesComErro += 1
                    LogErros.RegistrarErro(connEx, $"PowerQueryManager.AtualizarConexoesPowerQueryOtimizado - Erro na conexão {connection.Name}")
                End Try
            Next

            ' Atualizar tabelas críticas apenas
            AtualizarTabelasCriticas()

            LogErros.RegistrarInfo($"Conexões atualizadas: {conexoesAtualizadas}, Com erro: {conexoesComErro}", "PowerQueryManager.AtualizarConexoesPowerQueryOtimizado")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.AtualizarConexoesPowerQueryOtimizado")
            Throw
        End Try
    End Sub

    Private Sub AtualizarTabelasCriticas()
        Try
            ' Lista de tabelas críticas para o funcionamento do sistema
            Dim tabelasCriticas As String() = {
                ConfiguracaoApp.TABELA_PRODUTOS,
                ConfiguracaoApp.TABELA_ESTOQUE,
                ConfiguracaoApp.TABELA_COMPRAS,
                ConfiguracaoApp.TABELA_VENDAS
            }

            For Each nomeTabela In tabelasCriticas
                Try
                    Dim tabela = ObterTabela(nomeTabela)
                    If tabela IsNot Nothing AndAlso tabela.QueryTable IsNot Nothing Then
                        tabela.QueryTable.Refresh()
                        LogErros.RegistrarInfo($"Tabela crítica atualizada: {nomeTabela}", "PowerQueryManager.AtualizarTabelasCriticas")
                    End If
                Catch tableEx As Exception
                    LogErros.RegistrarErro(tableEx, $"PowerQueryManager.AtualizarTabelasCriticas - Erro na tabela {nomeTabela}")
                End Try
            Next

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.AtualizarTabelasCriticas")
        End Try
    End Sub

    Private Function EhConnectionPowerQueryEssencial(connection As WorkbookConnection) As Boolean
        Try
            ' Verificar se é uma conexão essencial para o sistema
            Dim nomeConnection = connection.Name.ToLower()

            ' Lista de padrões de nomes de conexões essenciais
            Dim padroesCriticos As String() = {
                "produtos", "estoque", "compras", "vendas",
                "tblprodutos", "tblestoque", "tblcompras", "tblvendas"
            }

            For Each padrao In padroesCriticos
                If nomeConnection.Contains(padrao) Then
                    Return True
                End If
            Next

            ' Verificar tipos de conexão
            Return connection.Type = XlConnectionType.xlConnectionTypeOLEDB OrElse
                   connection.Type = XlConnectionType.xlConnectionTypeODBC OrElse
                   connection.Type = XlConnectionType.xlConnectionTypeTEXT

        Catch
            Return False
        End Try
    End Function

    Private Sub AguardarConclusaoAtualizacaoOtimizada(startTime As DateTime)
        Try
            Dim timeoutReduzido = Math.Min(timeoutSegundos, 20) ' Máximo 20 segundos
            Dim intervalosVerificacao As Integer = 0
            Const maxIntervalos As Integer = 50 ' Reduzido de 100 para 50

            Do While workbook.Application.CalculationState <> XlCalculationState.xlDone
                WinFormsApp.DoEvents()
                Thread.Sleep(200) ' Aumentado de 100ms para 200ms para reduzir CPU
                intervalosVerificacao += 1

                ' Verificar timeout com tempo reduzido
                If DateTime.Now.Subtract(startTime).TotalSeconds > timeoutReduzido Then
                    Throw New TimeoutException($"Timeout de {timeoutReduzido} segundos excedido ao aguardar atualização das consultas")
                End If

                ' Sair mais rapidamente se necessário
                If intervalosVerificacao > maxIntervalos Then
                    LogErros.RegistrarInfo($"Muitas verificações ({intervalosVerificacao}), forçando saída", "PowerQueryManager.AguardarConclusaoAtualizacaoOtimizada")
                    Exit Do
                End If
            Loop

            ' Aguardar final reduzido
            Thread.Sleep(200) ' Reduzido de 500ms para 200ms
            WinFormsApp.DoEvents()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.AguardarConclusaoAtualizacaoOtimizada")
            Throw
        End Try
    End Sub

    Public Function ObterTabela(nomeTabela As String) As ListObject
        Try
            If String.IsNullOrEmpty(nomeTabela) Then
                Return Nothing
            End If

            ' Verificar cache primeiro
            If CacheEstaValido() AndAlso tabelasCache.ContainsKey(nomeTabela) Then
                Dim tabelaCache = tabelasCache(nomeTabela)
                If tabelaCache IsNot Nothing Then
                    Return tabelaCache
                End If
            End If

            LogErros.RegistrarInfo($"Procurando tabela: {nomeTabela}", "PowerQueryManager.ObterTabela")

            ' Buscar tabela otimizada
            For Each worksheet As Worksheet In workbook.Worksheets
                Try
                    For Each tabela As ListObject In worksheet.ListObjects
                        If tabela.Name.Equals(nomeTabela, StringComparison.OrdinalIgnoreCase) Then
                            ' Armazenar no cache
                            tabelasCache(nomeTabela) = tabela
                            cacheValido = DateTime.Now

                            LogErros.RegistrarInfo($"Tabela encontrada: {nomeTabela} na planilha {worksheet.Name}", "PowerQueryManager.ObterTabela")
                            Return tabela
                        End If
                    Next
                Catch wsEx As Exception
                    LogErros.RegistrarErro(wsEx, $"PowerQueryManager.ObterTabela - Erro na planilha {worksheet.Name}")
                    Continue For
                End Try
            Next

            LogErros.RegistrarInfo($"Tabela não encontrada: {nomeTabela}", "PowerQueryManager.ObterTabela")
            Return Nothing

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"PowerQueryManager.ObterTabela({nomeTabela})")
            Return Nothing
        End Try
    End Function

    Private Sub InvalidarCache()
        tabelasCache.Clear()
        cacheValido = DateTime.MinValue
    End Sub

    Private Function CacheEstaValido() As Boolean
        Return DateTime.Now.Subtract(cacheValido).TotalMinutes < CACHE_TIMEOUT_MINUTES
    End Function

    Public Function ListarTabelas() As List(Of String)
        Try
            Dim tabelas As New List(Of String)()

            For Each worksheet As Worksheet In workbook.Worksheets
                Try
                    For Each tabela As ListObject In worksheet.ListObjects
                        tabelas.Add($"{tabela.Name} ({worksheet.Name})")
                        LogErros.RegistrarInfo($"Tabela encontrada: {tabela.Name} na planilha {worksheet.Name}", "PowerQueryManager.ListarTabelas")
                    Next
                Catch wsEx As Exception
                    LogErros.RegistrarErro(wsEx, $"PowerQueryManager.ListarTabelas - Erro na planilha {worksheet.Name}")
                    Continue For
                End Try
            Next

            LogErros.RegistrarInfo($"Total de tabelas encontradas: {tabelas.Count}", "PowerQueryManager.ListarTabelas")
            Return tabelas

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ListarTabelas")
            Return New List(Of String)()
        End Try
    End Function

    Public Function VerificarStatusConexoes() As Dictionary(Of String, String)
        Try
            Dim status As New Dictionary(Of String, String)()

            For Each connection As WorkbookConnection In workbook.Connections
                Try
                    Dim tipoConexao As String = connection.Type.ToString()
                    Dim isPowerQuery As Boolean = EhConnectionPowerQueryEssencial(connection)

                    status.Add(connection.Name, $"Ativa ({tipoConexao}){If(isPowerQuery, " [Power Query]", "")}")

                Catch connEx As Exception
                    status.Add(connection.Name, $"Erro: {connEx.Message}")
                    LogErros.RegistrarErro(connEx, $"PowerQueryManager.VerificarStatusConexoes - Erro na conexão {connection.Name}")
                End Try
            Next

            LogErros.RegistrarInfo($"Status verificado para {status.Count} conexões", "PowerQueryManager.VerificarStatusConexoes")
            Return status

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.VerificarStatusConexoes")
            Return New Dictionary(Of String, String)()
        End Try
    End Function

    Public Function ObterInformacoesTabela(nomeTabela As String) As Dictionary(Of String, Object)
        Try
            Dim info As New Dictionary(Of String, Object)()
            Dim tabela As ListObject = ObterTabela(nomeTabela)

            If tabela IsNot Nothing Then
                info.Add("Nome", tabela.Name)
                info.Add("Planilha", tabela.Parent.Name)
                info.Add("Linhas", tabela.ListRows.Count)
                info.Add("Colunas", tabela.ListColumns.Count)
                info.Add("TemCabeçalho", tabela.ShowHeaders)
                info.Add("Endereço", tabela.Range.Address)

                If tabela.QueryTable IsNot Nothing Then
                    info.Add("TemQuery", True)
                    info.Add("TipoSource", "Power Query")
                Else
                    info.Add("TemQuery", False)
                    info.Add("TipoSource", "Dados manuais")
                End If
            End If

            Return info

        Catch ex As Exception
            LogErros.RegistrarErro(ex, $"PowerQueryManager.ObterInformacoesTabela({nomeTabela})")
            Return New Dictionary(Of String, Object)()
        End Try
    End Function

End Class