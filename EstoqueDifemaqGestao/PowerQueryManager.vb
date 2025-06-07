Imports Microsoft.Office.Interop.Excel
Imports System.Threading

Public Class PowerQueryManager
    Private workbook As Workbook
    Private timeoutSegundos As Integer  ' Removido ReadOnly para evitar conflito
    Private maxTentativas As Integer = 3  ' Removido ReadOnly para evitar conflito

    Public Sub New(wb As Workbook)
        If wb Is Nothing Then
            Throw New ArgumentNullException("workbook", "Workbook não pode ser nulo")
        End If
        workbook = wb
        timeoutSegundos = ConfiguracaoApp.TIMEOUT_POWERQUERY
    End Sub

    Public Sub AtualizarTodasConsultas()
        Try
            Dim startTime As DateTime = DateTime.Now
            Dim tentativa As Integer = 1

            LogErros.RegistrarInfo("Iniciando atualização de consultas Power Query", "PowerQueryManager.AtualizarTodasConsultas")

            ' Desabilitar alertas e eventos temporariamente
            Dim estadoAnterior = SalvarEstadoAplicacao()

            Try
                ' Tentar atualizar até o máximo de tentativas
                While tentativa <= maxTentativas
                    Try
                        LogErros.RegistrarInfo(String.Format("Tentativa {0} de {1}", tentativa, maxTentativas), "PowerQueryManager.AtualizarTodasConsultas")

                        ' Atualizar todas as conexões Power Query
                        AtualizarConexoesPowerQuery()

                        ' Aguardar conclusão com timeout
                        AguardarConclusaoAtualizacao(startTime)

                        ' Se chegou até aqui, sucesso
                        LogErros.RegistrarInfo(String.Format("Consultas atualizadas com sucesso na tentativa {0}", tentativa), "PowerQueryManager.AtualizarTodasConsultas")
                        Exit While

                    Catch ex As TimeoutException
                        LogErros.RegistrarErro(ex, String.Format("PowerQueryManager.AtualizarTodasConsultas - Timeout na tentativa {0}", tentativa))
                        tentativa += 1
                        If tentativa <= maxTentativas Then
                            ' Aguardar antes da próxima tentativa
                            Thread.Sleep(2000)
                            startTime = DateTime.Now
                        End If

                    Catch ex As Exception
                        LogErros.RegistrarErro(ex, String.Format("PowerQueryManager.AtualizarTodasConsultas - Erro na tentativa {0}", tentativa))
                        tentativa += 1
                        If tentativa <= maxTentativas Then
                            Thread.Sleep(1000)
                            startTime = DateTime.Now
                        End If
                    End Try
                End While

                ' Se todas as tentativas falharam
                If tentativa > maxTentativas Then
                    Throw New Exception(String.Format("Falha ao atualizar consultas após {0} tentativas", maxTentativas))
                End If

            Finally
                ' Restaurar configuração anterior
                RestaurarEstadoAplicacao(estadoAnterior)
            End Try

            LogErros.RegistrarInfo("Atualização de consultas Power Query concluída", "PowerQueryManager.AtualizarTodasConsultas")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.AtualizarTodasConsultas")
            Throw New Exception(String.Format("Erro ao atualizar consultas Power Query: {0}", ex.Message))
        End Try
    End Sub

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
            LogErros.RegistrarErro(ex, "PowerQueryManager.SalvarEstadoAplicacao")
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
            LogErros.RegistrarErro(ex, "PowerQueryManager.RestaurarEstadoAplicacao")
        End Try
    End Sub

    Private Sub AtualizarConexoesPowerQuery()
        Try
            Dim conexoesAtualizadas As Integer = 0
            Dim conexoesComErro As Integer = 0

            ' Atualizar workbook connections
            For Each connection As WorkbookConnection In workbook.Connections
                Try
                    If EhConnectionPowerQuery(connection) Then
                        LogErros.RegistrarInfo(String.Format("Atualizando conexão: {0}", connection.Name), "PowerQueryManager.AtualizarConexoesPowerQuery")
                        connection.Refresh()
                        conexoesAtualizadas += 1
                    End If
                Catch connEx As Exception
                    conexoesComErro += 1
                    LogErros.RegistrarErro(connEx, String.Format("PowerQueryManager.AtualizarConexoesPowerQuery - Erro na conexão {0}", connection.Name))
                End Try
            Next

            ' Verificar se há queries nas planilhas
            AtualizarQueriesPlanilhas()

            LogErros.RegistrarInfo(String.Format("Conexões atualizadas: {0}, Com erro: {1}", conexoesAtualizadas, conexoesComErro), "PowerQueryManager.AtualizarConexoesPowerQuery")

            If conexoesAtualizadas = 0 Then
                LogErros.RegistrarInfo("Nenhuma conexão Power Query encontrada", "PowerQueryManager.AtualizarConexoesPowerQuery")
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.AtualizarConexoesPowerQuery")
            Throw
        End Try
    End Sub

    Private Sub AtualizarQueriesPlanilhas()
        Try
            For Each worksheet As Worksheet In workbook.Worksheets
                Try
                    ' Atualizar ListObjects (tabelas) que podem ter consultas
                    For Each listObj As ListObject In worksheet.ListObjects
                        Try
                            If listObj.QueryTable IsNot Nothing Then
                                listObj.QueryTable.Refresh()
                                LogErros.RegistrarInfo(String.Format("Tabela atualizada: {0} na planilha {1}", listObj.Name, worksheet.Name), "PowerQueryManager.AtualizarQueriesPlanilhas")
                            End If
                        Catch tableEx As Exception
                            LogErros.RegistrarErro(tableEx, String.Format("PowerQueryManager.AtualizarQueriesPlanilhas - Erro na tabela {0}", listObj.Name))
                        End Try
                    Next

                    ' Atualizar PivotTables se houver
                    For Each pivotTable As PivotTable In worksheet.PivotTables
                        Try
                            pivotTable.RefreshTable()
                            LogErros.RegistrarInfo(String.Format("PivotTable atualizada: {0} na planilha {1}", pivotTable.Name, worksheet.Name), "PowerQueryManager.AtualizarQueriesPlanilhas")
                        Catch pivotEx As Exception
                            LogErros.RegistrarErro(pivotEx, String.Format("PowerQueryManager.AtualizarQueriesPlanilhas - Erro na PivotTable {0}", pivotTable.Name))
                        End Try
                    Next

                Catch wsEx As Exception
                    LogErros.RegistrarErro(wsEx, String.Format("PowerQueryManager.AtualizarQueriesPlanilhas - Erro na planilha {0}", worksheet.Name))
                    Continue For
                End Try
            Next
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.AtualizarQueriesPlanilhas")
        End Try
    End Sub

    Private Function EhConnectionPowerQuery(connection As WorkbookConnection) As Boolean
        Try
            ' Verificar tipos de conexão que podem ser Power Query
            Return connection.Type = XlConnectionType.xlConnectionTypeOLEDB OrElse
                   connection.Type = XlConnectionType.xlConnectionTypeODBC OrElse
                   connection.Type = XlConnectionType.xlConnectionTypeTEXT OrElse
                   connection.Type = XlConnectionType.xlConnectionTypeWEB OrElse
                   connection.Type = XlConnectionType.xlConnectionTypeXMLMAP
        Catch
            Return False
        End Try
    End Function

    Private Sub AguardarConclusaoAtualizacao(startTime As DateTime)
        Try
            Dim ultimoCheck As DateTime = DateTime.Now
            Dim intervalosVerificacao As Integer = 0
            Const maxIntervalos As Integer = 100 ' Máximo de verificações

            Do While workbook.Application.CalculationState <> XlCalculationState.xlDone
                System.Windows.Forms.Application.DoEvents()
                Thread.Sleep(100)
                intervalosVerificacao += 1

                ' Log de progresso a cada 5 segundos
                If DateTime.Now.Subtract(ultimoCheck).TotalSeconds >= 5 Then
                    LogErros.RegistrarInfo(String.Format("Aguardando conclusão... ({0:F1}s)", DateTime.Now.Subtract(startTime).TotalSeconds), "PowerQueryManager.AguardarConclusaoAtualizacao")
                    ultimoCheck = DateTime.Now
                End If

                ' Verificar timeout
                If DateTime.Now.Subtract(startTime).TotalSeconds > timeoutSegundos Then
                    Throw New TimeoutException(String.Format("Timeout de {0} segundos excedido ao aguardar atualização das consultas", timeoutSegundos))
                End If

                ' Verificar se não está travado em loop infinito
                If intervalosVerificacao > maxIntervalos Then
                    LogErros.RegistrarInfo(String.Format("Muitas verificações ({0}), forçando saída", intervalosVerificacao), "PowerQueryManager.AguardarConclusaoAtualizacao")
                    Exit Do
                End If
            Loop

            ' Aguardar um pouco mais para garantir que tudo foi processado
            Thread.Sleep(500)
            System.Windows.Forms.Application.DoEvents()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.AguardarConclusaoAtualizacao")
            Throw
        End Try
    End Sub

    Public Function ObterTabela(nomeTabela As String) As ListObject
        Try
            If String.IsNullOrEmpty(nomeTabela) Then
                LogErros.RegistrarInfo("Nome da tabela está vazio", "PowerQueryManager.ObterTabela")
                Return Nothing
            End If

            LogErros.RegistrarInfo(String.Format("Procurando tabela: {0}", nomeTabela), "PowerQueryManager.ObterTabela")

            For Each worksheet As Worksheet In workbook.Worksheets
                Try
                    For Each tabela As ListObject In worksheet.ListObjects
                        If tabela.Name.Equals(nomeTabela, StringComparison.OrdinalIgnoreCase) Then
                            LogErros.RegistrarInfo(String.Format("Tabela encontrada: {0} na planilha {1}", nomeTabela, worksheet.Name), "PowerQueryManager.ObterTabela")
                            Return tabela
                        End If
                    Next
                Catch wsEx As Exception
                    LogErros.RegistrarErro(wsEx, String.Format("PowerQueryManager.ObterTabela - Erro na planilha {0}", worksheet.Name))
                    Continue For
                End Try
            Next

            LogErros.RegistrarInfo(String.Format("Tabela não encontrada: {0}", nomeTabela), "PowerQueryManager.ObterTabela")
            Return Nothing

        Catch ex As Exception
            LogErros.RegistrarErro(ex, String.Format("PowerQueryManager.ObterTabela({0})", nomeTabela))
            Return Nothing
        End Try
    End Function

    Public Function ListarTabelas() As List(Of String)
        Try
            Dim tabelas As New List(Of String)()

            For Each worksheet As Worksheet In workbook.Worksheets
                Try
                    For Each tabela As ListObject In worksheet.ListObjects
                        tabelas.Add(String.Format("{0} ({1})", tabela.Name, worksheet.Name))
                        LogErros.RegistrarInfo(String.Format("Tabela encontrada: {0} na planilha {1}", tabela.Name, worksheet.Name), "PowerQueryManager.ListarTabelas")
                    Next
                Catch wsEx As Exception
                    LogErros.RegistrarErro(wsEx, String.Format("PowerQueryManager.ListarTabelas - Erro na planilha {0}", worksheet.Name))
                    Continue For
                End Try
            Next

            LogErros.RegistrarInfo(String.Format("Total de tabelas encontradas: {0}", tabelas.Count), "PowerQueryManager.ListarTabelas")
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
                    Dim isPowerQuery As Boolean = EhConnectionPowerQuery(connection)

                    status.Add(connection.Name, String.Format("Ativa ({0}){1}", tipoConexao, If(isPowerQuery, " [Power Query]", "")))

                Catch connEx As Exception
                    status.Add(connection.Name, String.Format("Erro: {0}", connEx.Message))
                    LogErros.RegistrarErro(connEx, String.Format("PowerQueryManager.VerificarStatusConexoes - Erro na conexão {0}", connection.Name))
                End Try
            Next

            LogErros.RegistrarInfo(String.Format("Status verificado para {0} conexões", status.Count), "PowerQueryManager.VerificarStatusConexoes")
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
            LogErros.RegistrarErro(ex, String.Format("PowerQueryManager.ObterInformacoesTabela({0})", nomeTabela))
            Return New Dictionary(Of String, Object)()
        End Try
    End Function

End Class