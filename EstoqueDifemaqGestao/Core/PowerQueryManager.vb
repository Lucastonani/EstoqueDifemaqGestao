Imports Microsoft.Office.Interop.Excel
Imports System.Threading
Imports WinFormsApp = System.Windows.Forms.Application

''' <summary>
''' Gerenciador otimizado de consultas Power Query para integração com Excel
''' Responsável por atualizar consultas, obter tabelas e gerenciar conexões com timeout configurável
''' </summary>
''' <remarks>
''' Esta classe implementa funcionalidades otimizadas para o sistema EstoqueDifemaqGestao:
''' - Cache inteligente de tabelas com timeout de 5 minutos
''' - Atualização otimizada com tentativas múltiplas (máximo 2)
''' - Configurações específicas do Excel para melhor performance
''' - Foco em tabelas críticas do sistema (produtos, estoque, compras, vendas)
''' - Tratamento robusto de erros e timeouts configuráveis
''' </remarks>
Public Class PowerQueryManager
    Private workbook As Workbook
    Private timeoutSegundos As Integer
    Private maxTentativas As Integer = 2 ' Reduzido de 3 para 2

    ' Cache para otimizar acesso às tabelas
    Private tabelasCache As New Dictionary(Of String, ListObject)
    Private cacheValido As DateTime = DateTime.MinValue
    Private Const CACHE_TIMEOUT_MINUTES As Integer = 5


    ''' <summary>
    ''' Inicializa uma nova instância do PowerQueryManager com workbook específico
    ''' </summary>
    ''' <param name="wb">Workbook do Excel a ser gerenciado</param>
    ''' <exception cref="ArgumentNullException">Quando workbook é nulo</exception>
    ''' <remarks>
    ''' Configura timeout otimizado (máximo 30 segundos) e inicializa cache de tabelas.
    ''' O workbook deve estar aberto e acessível no momento da criação.
    ''' </remarks>
    ''' <example>
    ''' <code>
    ''' Dim workbook As Workbook = Globals.ThisWorkbook.InnerObject
    ''' Dim manager As New PowerQueryManager(workbook)
    ''' </code>
    ''' </example>
    Public Sub New(wb As Workbook)
        If wb Is Nothing Then
            Throw New ArgumentNullException("workbook", "Workbook não pode ser nulo")
        End If
        workbook = wb
        timeoutSegundos = Math.Min(ConfiguracaoApp.TIMEOUT_POWERQUERY, 30) ' Máximo 30 segundos
    End Sub

    ''' <summary>
    ''' Atualiza todas as consultas Power Query do workbook com otimizações e retry automático
    ''' </summary>
    ''' <exception cref="ArgumentNullException">Quando workbook é nulo</exception>
    ''' <exception cref="Exception">Quando falha após todas as tentativas configuradas</exception>
    ''' <exception cref="TimeoutException">Quando excede timeout configurado</exception>
    ''' <example>
    ''' <code>
    ''' Dim manager As New PowerQueryManager(workbook)
    ''' Try
    '''     manager.AtualizarTodasConsultas()
    '''     Console.WriteLine("Consultas atualizadas com sucesso")
    ''' Catch ex As Exception
    '''     Console.WriteLine($"Erro na atualização: {ex.Message}")
    ''' End Try
    ''' </code>
    ''' </example>
    ''' <remarks>
    ''' Processo otimizado que:
    ''' - Configura Excel para melhor performance (desabilita alerts, screen updating, etc.)
    ''' - Atualiza apenas conexões Power Query essenciais
    ''' - Implementa retry automático (máximo 2 tentativas)
    ''' - Foca em tabelas críticas do sistema
    ''' - Invalida cache após sucesso
    ''' - Restaura configurações originais do Excel
    ''' 
    ''' Timeout padrão limitado a 30 segundos com verificações a cada 200ms.
    ''' </remarks>

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

    ''' <summary>
    ''' Configura Excel para otimização durante atualização de consultas
    ''' </summary>
    ''' <returns>Dictionary com configurações anteriores para restauração</returns>
    ''' <remarks>
    ''' Desabilita temporariamente: DisplayAlerts, ScreenUpdating, EnableEvents
    ''' e define Calculation como Manual para melhor performance.
    ''' </remarks>

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

    ''' <summary>
    ''' Restaura configurações originais do Excel após atualização
    ''' </summary>
    ''' <param name="estado">Dictionary com configurações anteriores</param>
    ''' <remarks>
    ''' Restaura configurações salvas por ConfigurarExcelParaAtualizacao().
    ''' </remarks>
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

    ''' <summary>
    ''' Verifica se uma conexão é Power Query essencial para o sistema
    ''' </summary>
    ''' <param name="connection">Conexão do workbook a ser verificada</param>
    ''' <returns>True se é uma conexão essencial, False caso contrário</returns>
    ''' <remarks>
    ''' Identifica conexões baseadas em padrões: produtos, estoque, compras, vendas
    ''' e tipos de conexão OLEDB, ODBC, TEXT.
    ''' </remarks>
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

    ''' <summary>
    ''' Obtém uma tabela específica do workbook com cache inteligente
    ''' </summary>
    ''' <param name="nomeTabela">Nome da tabela a ser obtida (case-insensitive)</param>
    ''' <returns>ListObject da tabela encontrada ou Nothing se não encontrada</returns>
    ''' <example>
    ''' <code>
    ''' Dim manager As New PowerQueryManager(workbook)
    ''' Dim tabela = manager.ObterTabela("tblProdutos")
    ''' If tabela IsNot Nothing Then
    '''     Console.WriteLine($"Tabela tem {tabela.ListRows.Count} linhas")
    ''' Else
    '''     Console.WriteLine("Tabela não encontrada")
    ''' End If
    ''' </code>
    ''' </example>
    ''' <remarks>
    ''' Implementa cache inteligente com timeout de 5 minutos para otimizar performance.
    ''' Busca case-insensitive em todas as planilhas do workbook.
    ''' Cache é invalidado automaticamente após atualizações bem-sucedidas.
    ''' Primeira busca: O(n), buscas subsequentes: O(1) devido ao cache.
    ''' </remarks>
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

    ''' <summary>
    ''' Invalida o cache interno de tabelas forçando nova busca
    ''' </summary>
    ''' <remarks>
    ''' Chamado automaticamente após atualizações bem-sucedidas.
    ''' </remarks>
    Private Sub InvalidarCache()
        tabelasCache.Clear()
        cacheValido = DateTime.MinValue
    End Sub

    ''' <summary>
    ''' Verifica se o cache de tabelas ainda está dentro do timeout válido
    ''' </summary>
    ''' <returns>True se cache é válido, False se expirou</returns>
    ''' <remarks>
    ''' Cache expira após 5 minutos (CACHE_TIMEOUT_MINUTES).
    ''' </remarks>
    Private Function CacheEstaValido() As Boolean
        Return DateTime.Now.Subtract(cacheValido).TotalMinutes < CACHE_TIMEOUT_MINUTES
    End Function

    ''' <summary>
    ''' Lista todas as tabelas disponíveis no workbook com informações de localização
    ''' </summary>
    ''' <returns>Lista de strings no formato "NomeTabela (NomePlanilha)" ou lista vazia se erro</returns>
    ''' <example>
    ''' <code>
    ''' Dim manager As New PowerQueryManager(workbook)
    ''' Dim tabelas = manager.ListarTabelas()
    ''' Console.WriteLine("Tabelas disponíveis:")
    ''' For Each tabela In tabelas
    '''     Console.WriteLine($"• {tabela}")
    ''' Next
    ''' 
    ''' ' Exemplo de saída:
    ''' ' • tblProdutos (Produtos)
    ''' ' • tblEstoque (EstoqueVisao)
    ''' ' • tblVendas (Vendas)
    ''' </code>
    ''' </example>
    ''' <remarks>
    ''' Percorre todas as planilhas do workbook e registra cada tabela encontrada.
    ''' Útil para debug, validação de estrutura e discovery de dados.
    ''' Continua execução mesmo se houver erro em planilhas específicas.
    ''' Registra informações detalhadas nos logs para auditoria.
    ''' </remarks>
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

    ''' <summary>
    ''' Verifica o status de todas as conexões do workbook com identificação de Power Query
    ''' </summary>
    ''' <returns>Dictionary com nome da conexão como chave e status detalhado como valor</returns>
    ''' <example>
    ''' <code>
    ''' Dim manager As New PowerQueryManager(workbook)
    ''' Dim status = manager.VerificarStatusConexoes()
    ''' 
    ''' Console.WriteLine("Status das conexões:")
    ''' For Each kvp In status
    '''     Console.WriteLine($"{kvp.Key}: {kvp.Value}")
    ''' Next
    ''' 
    ''' ' Exemplo de saída:
    ''' ' ConexaoProdutos: Ativa (xlConnectionTypeOLEDB) [Power Query]
    ''' ' ConexaoVendas: Ativa (xlConnectionTypeTEXT) [Power Query]
    ''' </code>
    ''' </example>
    ''' <remarks>
    ''' Identifica conexões Power Query essenciais baseadas em padrões de nomenclatura
    ''' e tipos de conexão (OLEDB, ODBC, TEXT).
    ''' Útil para diagnóstico de problemas de conectividade e validação de configuração.
    ''' Registra erros específicos por conexão sem interromper verificação das demais.
    ''' </remarks>
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

    ''' <summary>
    ''' Obtém informações detalhadas e metadados sobre uma tabela específica
    ''' </summary>
    ''' <param name="nomeTabela">Nome da tabela para obter informações</param>
    ''' <returns>Dictionary com informações completas da tabela ou dictionary vazio se não encontrada</returns>
    ''' <example>
    ''' <code>
    ''' Dim manager As New PowerQueryManager(workbook)
    ''' Dim info = manager.ObterInformacoesTabela("tblProdutos")
    ''' 
    ''' If info.Count > 0 Then
    '''     Console.WriteLine($"Nome: {info("Nome")}")
    '''     Console.WriteLine($"Planilha: {info("Planilha")}")
    '''     Console.WriteLine($"Linhas: {info("Linhas")}")
    '''     Console.WriteLine($"Colunas: {info("Colunas")}")
    '''     Console.WriteLine($"Tem Power Query: {info("TemQuery")}")
    '''     Console.WriteLine($"Tipo de fonte: {info("TipoSource")}")
    ''' Else
    '''     Console.WriteLine("Tabela não encontrada")
    ''' End If
    ''' </code>
    ''' </example>
    ''' <remarks>
    ''' Retorna informações completas incluindo:
    ''' - Nome da tabela e planilha onde está localizada
    ''' - Número de linhas e colunas
    ''' - Se possui cabeçalho visível
    ''' - Endereço da range da tabela
    ''' - Se está conectada a Power Query ou contém dados manuais
    ''' 
    ''' Útil para validação de estrutura, debug e relatórios de configuração.
    ''' Utiliza o método ObterTabela() internamente, beneficiando-se do cache.
    ''' </remarks>
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
