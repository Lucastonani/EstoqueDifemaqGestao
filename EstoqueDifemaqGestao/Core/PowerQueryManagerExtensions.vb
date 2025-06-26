Imports Microsoft.Office.Interop.Excel

Partial Public Class PowerQueryManager

    ''' <summary>
    ''' Método singleton para obter instância do PowerQueryManager
    ''' </summary>
    Private Shared _instance As PowerQueryManager
    Private Shared _instanceLock As New Object()

    Public Shared Function GetInstance() As PowerQueryManager
        If _instance Is Nothing Then
            SyncLock _instanceLock
                If _instance Is Nothing Then
                    Try
                        ' Tenta obter o workbook ativo
                        Dim wb As Workbook = Nothing

                        ' Verifica se estamos em ambiente VSTO
                        If Globals.ThisWorkbook IsNot Nothing Then
                            wb = Globals.ThisWorkbook.InnerObject
                        End If

                        If wb IsNot Nothing Then
                            _instance = New PowerQueryManager(wb)
                        End If
                    Catch ex As Exception
                        LogErros.RegistrarErro(ex, "PowerQueryManager.GetInstance")
                    End Try
                End If
            End SyncLock
        End If

        Return _instance
    End Function

    ''' <summary>
    ''' Atualiza dados das consultas Power Query
    ''' </summary>
    Public Sub AtualizarDados()
        Try
            AtualizarTodasConsultas()
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.AtualizarDados")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Obtém lista de produtos da tabela tblProdutos
    ''' </summary>
    Public Function ObterProdutos() As System.Data.DataTable
        Try
            Dim dt As New System.Data.DataTable()
            Dim tabela = ObterTabela("tblProdutos")

            If tabela IsNot Nothing Then
                ' Criar colunas
                For Each col As ListColumn In tabela.ListColumns
                    dt.Columns.Add(col.Name, GetType(String))
                Next

                ' Adicionar dados
                For Each row As ListRow In tabela.ListRows
                    Dim dataRow = dt.NewRow()
                    For i = 0 To tabela.ListColumns.Count - 1
                        Dim valor = row.Range(i + 1).Value
                        dataRow(i) = If(valor IsNot Nothing, valor.ToString(), "")
                    Next
                    dt.Rows.Add(dataRow)
                Next
            End If

            Return dt

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ObterProdutos")
            Return New System.Data.DataTable()
        End Try
    End Function

    ''' <summary>
    ''' Obtém dados de estoque para um produto específico
    ''' </summary>
    Public Function ObterEstoqueProduto(codigoProduto As String) As System.Data.DataTable
        Try
            LogErros.RegistrarInfo($"🔍 Buscando estoque para produto: {codigoProduto}", "ObterEstoqueProduto")
            
            Dim dt As New System.Data.DataTable()
            Dim tabela = ObterTabela("tblEstoqueVisao")

            If tabela IsNot Nothing Then
                LogErros.RegistrarInfo($"✅ Tabela tblEstoqueVisao encontrada com {tabela.ListRows.Count} linhas", "ObterEstoqueProduto")

                ' ✅ OTIMIZAÇÃO: Usar busca rápida para tabelas grandes (limiar muito reduzido)
                If tabela.ListRows.Count > 20000 Then
                    LogErros.RegistrarInfo("🚀 Usando busca ultra-rápida para tabela grande", "ObterEstoqueProduto")
                    Try
                        Return PowerQueryManagerOptimizedSearch.BuscarEstoqueRapido(tabela, codigoProduto)
                    Catch searchEx As Exception
                        LogErros.RegistrarErro(searchEx, "ObterEstoqueProduto.BuscaRapida")
                        ' Fallback para método limitado
                    End Try
                End If

                ' ✅ OTIMIZAÇÃO: Usar AutoFilter para filtrar no Excel antes de iterar
                Dim worksheet = tabela.Parent
                Dim originalAutoFilterMode = worksheet.AutoFilterMode

                Try
                    ' Aplicar filtro no Excel para reduzir dados
                    If Not tabela.ShowAutoFilter Then
                        tabela.ShowAutoFilter = True
                    End If

                    ' Filtrar pela primeira coluna (código do produto)
                    tabela.Range.AutoFilter(1, codigoProduto)

                    ' Criar colunas no DataTable
                    For Each col As ListColumn In tabela.ListColumns
                        dt.Columns.Add(col.Name, GetType(Object))
                    Next
                    LogErros.RegistrarInfo($"✅ Criadas {dt.Columns.Count} colunas no DataTable", "ObterEstoqueProduto")

                    ' ✅ OTIMIZAÇÃO: Processar apenas linhas visíveis (filtradas) com limite
                    Dim registrosEncontrados = 0
                    Dim maxRegistros = 100 ' Limitar para evitar OutOfMemory

                    Try
                        Dim visibleCells = tabela.Range.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible)

                        If visibleCells IsNot Nothing Then
                            For Each area In visibleCells.Areas
                                If registrosEncontrados >= maxRegistros Then Exit For

                                For rowIndex = 2 To Math.Min(area.Rows.Count, maxRegistros + 1) ' Pular cabeçalho
                                    Try
                                        Dim dataRow = dt.NewRow()
                                        For i = 0 To tabela.ListColumns.Count - 1
                                            dataRow(i) = area.Cells(rowIndex, i + 1).Value
                                        Next
                                        dt.Rows.Add(dataRow)
                                        registrosEncontrados += 1

                                        If registrosEncontrados >= maxRegistros Then Exit For
                                    Catch cellEx As Exception
                                        ' Ignorar erro de célula específica
                                    End Try
                                Next
                            Next
                        End If
                    Catch visibleEx As Exception
                        LogErros.RegistrarErro(visibleEx, "ObterEstoqueProduto.SpecialCells")
                        ' Fallback: busca limitada linha por linha
                        For i = 1 To Math.Min(tabela.ListRows.Count, maxRegistros)
                            Try
                                Dim codigo = tabela.ListRows(i).Range(1).Value?.ToString()
                                If codigo = codigoProduto Then
                                    Dim dataRow = dt.NewRow()
                                    For j = 0 To tabela.ListColumns.Count - 1
                                        dataRow(j) = tabela.ListRows(i).Range(j + 1).Value
                                    Next
                                    dt.Rows.Add(dataRow)
                                    registrosEncontrados += 1
                                End If
                            Catch rowEx As Exception
                                ' Ignorar linha com erro
                            End Try
                        Next
                    End Try

                    LogErros.RegistrarInfo($"✅ Estoque encontrado: {registrosEncontrados} registros para produto {codigoProduto}", "ObterEstoqueProduto")

                Finally
                    ' Remover filtro para não afetar outras operações
                    Try
                        If tabela.ShowAutoFilter Then
                            tabela.Range.AutoFilter()
                        End If
                        worksheet.AutoFilterMode = originalAutoFilterMode
                    Catch
                        ' Ignorar erros de limpeza
                    End Try
                End Try
            Else
                LogErros.RegistrarInfo("⚠️ Tabela tblEstoqueVisao não encontrada", "ObterEstoqueProduto")
            End If

            Return dt

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ObterEstoqueProduto")
            Return New System.Data.DataTable()
        End Try
    End Function

    ''' <summary>
    ''' Obtém dados de compras para um produto específico
    ''' </summary>
    Public Function ObterComprasProduto(codigoProduto As String) As System.Data.DataTable
        Try
            Dim dt As New System.Data.DataTable()
            Dim tabela = ObterTabela("tblCompras")

            If tabela IsNot Nothing Then
                ' Criar colunas
                For Each col As ListColumn In tabela.ListColumns
                    dt.Columns.Add(col.Name, GetType(Object))
                Next

                ' Filtrar e adicionar dados
                For Each row As ListRow In tabela.ListRows
                    Dim codigo = row.Range(1).Value?.ToString()
                    If codigo = codigoProduto Then
                        Dim dataRow = dt.NewRow()
                        For i = 0 To tabela.ListColumns.Count - 1
                            dataRow(i) = row.Range(i + 1).Value
                        Next
                        dt.Rows.Add(dataRow)
                    End If
                Next
            End If

            Return dt

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ObterComprasProduto")
            Return New System.Data.DataTable()
        End Try
    End Function

    ''' <summary>
    ''' Obtém dados de vendas para um produto específico
    ''' </summary>
    Public Function ObterVendasProduto(codigoProduto As String) As System.Data.DataTable
        Try
            Dim dt As New System.Data.DataTable()
            Dim tabela = ObterTabela("tblVendas")

            If tabela IsNot Nothing Then
                ' Criar colunas
                For Each col As ListColumn In tabela.ListColumns
                    dt.Columns.Add(col.Name, GetType(Object))
                Next

                ' Filtrar e adicionar dados
                For Each row As ListRow In tabela.ListRows
                    Dim codigo = row.Range(1).Value?.ToString()
                    If codigo = codigoProduto Then
                        Dim dataRow = dt.NewRow()
                        For i = 0 To tabela.ListColumns.Count - 1
                            dataRow(i) = row.Range(i + 1).Value
                        Next
                        dt.Rows.Add(dataRow)
                    End If
                Next
            End If

            Return dt

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ObterVendasProduto")
            Return New System.Data.DataTable()
        End Try
    End Function

    ''' <summary>
    ''' Verifica se um código de produto já existe
    ''' </summary>
    Public Function VerificarCodigoExistente(codigo As String) As Boolean
        Try
            Dim tabela = ObterTabela("tblProdutos")

            If tabela IsNot Nothing Then
                For Each row As ListRow In tabela.ListRows
                    If row.Range(1).Value?.ToString() = codigo Then
                        Return True
                    End If
                Next
            End If

            ' Verificar também na tabela de produtos manuais
            Dim tabelaManual = ObterTabela("tblProdutosManual")
            If tabelaManual IsNot Nothing Then
                For Each row As ListRow In tabelaManual.ListRows
                    If row.Range(1).Value?.ToString() = codigo Then
                        Return True
                    End If
                Next
            End If

            Return False

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.VerificarCodigoExistente")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Insere um produto manual na tabela tblProdutosManual
    ''' </summary>
    Public Sub InserirProdutoManual(dadosProduto As Dictionary(Of String, Object))
        Try
            Dim tabela = ObterTabela("tblProdutosManual")

            If tabela IsNot Nothing Then
                ' Adicionar nova linha
                Dim novaLinha = tabela.ListRows.Add()

                ' Preencher dados
                novaLinha.Range(1).Value = dadosProduto("Codigo")
                novaLinha.Range(2).Value = dadosProduto("Descricao")
                novaLinha.Range(3).Value = dadosProduto("Fabricante")
                novaLinha.Range(4).Value = dadosProduto("QuantidadeInicial")
                novaLinha.Range(5).Value = dadosProduto("Data")
            Else
                Throw New Exception("Tabela tblProdutosManual não encontrada")
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.InserirProdutoManual")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Obtém lista de lojas distintas da tabela tblEstoqueVisao
    ''' </summary>
    Public Function ObterLojasDistintas() As List(Of String)
        Try
            LogErros.RegistrarInfo("🏢 Obtendo lojas distintas", "ObterLojasDistintas")

            Dim lojas As New HashSet(Of String)
            Dim tabela = ObterTabela("tblEstoqueVisao")

            If tabela IsNot Nothing Then
                LogErros.RegistrarInfo($"✅ Tabela tblEstoqueVisao encontrada com {tabela.ListRows.Count} linhas", "ObterLojasDistintas")

                ' ✅ OTIMIZAÇÃO: Ler valores da coluna diretamente em array
                Try
                    Dim lojaColumn = tabela.ListColumns(2).Range

                    ' Ler valores da coluna
                    Dim valores = lojaColumn.Value
                    If TypeOf valores Is Object(,) Then
                        Dim array2D As Object(,) = valores
                        For i = 2 To array2D.GetLength(0) ' Pular cabeçalho
                            Dim loja = array2D(i, 1)?.ToString()
                            If Not String.IsNullOrEmpty(loja) Then
                                lojas.Add(loja)
                                If lojas.Count >= 50 Then ' Limitar a 50 lojas diferentes
                                    Exit For
                                End If
                            End If
                        Next
                    ElseIf valores IsNot Nothing Then
                        Dim loja = valores.ToString()
                        If Not String.IsNullOrEmpty(loja) Then
                            lojas.Add(loja)
                        End If
                    End If

                Catch optimizationEx As Exception
                    LogErros.RegistrarErro(optimizationEx, "ObterLojasDistintas.Otimizacao")

                    ' ✅ FALLBACK: Método mais seguro com limite
                    Dim maxProcessar = Math.Min(1000, tabela.ListRows.Count) ' Limitar a 1000 registros
                    For i = 1 To maxProcessar
                        Try
                            Dim loja = tabela.ListRows(i).Range(2).Value?.ToString()
                            If Not String.IsNullOrEmpty(loja) Then
                                lojas.Add(loja)
                                If lojas.Count >= 50 Then ' Limitar a 50 lojas diferentes
                                    Exit For
                                End If
                            End If
                        Catch rowEx As Exception
                            ' Ignorar erro de linha específica
                        End Try
                    Next
                End Try

                LogErros.RegistrarInfo($"✅ Lojas encontradas: {lojas.Count}", "ObterLojasDistintas")
            Else
                LogErros.RegistrarInfo("⚠️ Tabela tblEstoqueVisao não encontrada", "ObterLojasDistintas")
            End If

            Return lojas.ToList()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ObterLojasDistintas")
            ' Retornar lista padrão em caso de erro
            Return New List(Of String) From {"Cariacica", "Vila Velha", "Serra"}
        End Try
    End Function

    ''' <summary>
    ''' Obtém histórico de compras agrupado por mês
    ''' </summary>
    Public Function ObterHistoricoComprasPorMes(codigoProduto As String, dataInicio As Date, dataFim As Date) As Dictionary(Of Date, Decimal)
        Try
            LogErros.RegistrarInfo($"📈 Buscando histórico de compras para produto: {codigoProduto}", "ObterHistoricoComprasPorMes")

            Dim resultado As New Dictionary(Of Date, Decimal)
            Dim tabela = ObterTabela("tblCompras")

            If tabela IsNot Nothing Then
                LogErros.RegistrarInfo($"✅ Tabela tblCompras encontrada com {tabela.ListRows.Count} linhas", "ObterHistoricoComprasPorMes")

                ' ✅ OTIMIZAÇÃO: Usar busca ultra-rápida para tabelas grandes (limiar reduzido)
                If tabela.ListRows.Count > 10000 Then
                    LogErros.RegistrarInfo("🚀 Usando busca ultra-rápida com fórmulas para tabela grande", "ObterHistoricoComprasPorMes")
                    Try
                        Return PowerQueryManagerOptimizedSearch.BuscarHistoricoRapido(tabela, codigoProduto, dataInicio, dataFim)
                    Catch searchEx As Exception
                        LogErros.RegistrarErro(searchEx, "ObterHistoricoComprasPorMes.BuscaRapida")
                        ' Fallback para método limitado
                    End Try
                End If

                ' Criar estrutura para todos os meses no período
                Dim dataAtual = New Date(dataInicio.Year, dataInicio.Month, 1)
                While dataAtual <= dataFim
                    resultado(dataAtual) = 0
                    dataAtual = dataAtual.AddMonths(1)
                End While
                LogErros.RegistrarInfo($"✅ Criados {resultado.Count} meses no período", "ObterHistoricoComprasPorMes")

                ' ✅ OTIMIZAÇÃO: Usar Range.Find para buscar mais eficientemente com limite
                Dim worksheet = tabela.Parent
                Dim codigoRange = tabela.ListColumns(1).Range
                Dim registrosProcessados = 0
                Dim maxRegistros = 1000 ' Limitar para evitar OutOfMemory

                Try
                    ' Buscar primeira ocorrência
                    Dim foundCell = codigoRange.Find(codigoProduto, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                                                    LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlWhole)

                    If foundCell IsNot Nothing Then
                        Dim firstAddress = foundCell.Address
                        Do
                            Try
                                If registrosProcessados >= maxRegistros Then Exit Do

                                Dim rowIndex = foundCell.Row - tabela.Range.Row + 1
                                If rowIndex > 1 AndAlso rowIndex <= tabela.ListRows.Count + 1 Then
                                    Dim dataCompra = CDate(tabela.ListColumns(2).Range.Cells(rowIndex).Value)
                                    Dim quantidade = CDec(tabela.ListColumns(3).Range.Cells(rowIndex).Value)

                                    If dataCompra >= dataInicio AndAlso dataCompra <= dataFim Then
                                        Dim mesChave = New Date(dataCompra.Year, dataCompra.Month, 1)
                                        If resultado.ContainsKey(mesChave) Then
                                            resultado(mesChave) += quantidade
                                            registrosProcessados += 1
                                        End If
                                    End If
                                End If
                            Catch cellEx As Exception
                                ' Ignorar erro de célula específica
                            End Try

                            ' Buscar próxima ocorrência
                            foundCell = codigoRange.FindNext(foundCell)
                        Loop While foundCell IsNot Nothing AndAlso foundCell.Address <> firstAddress
                    End If

                Catch findEx As Exception
                    LogErros.RegistrarErro(findEx, "ObterHistoricoComprasPorMes.Find")
                End Try

                LogErros.RegistrarInfo($"✅ Compras processadas: {registrosProcessados} registros para produto {codigoProduto}", "ObterHistoricoComprasPorMes")
            Else
                LogErros.RegistrarInfo("⚠️ Tabela tblCompras não encontrada", "ObterHistoricoComprasPorMes")
            End If

            Return resultado

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ObterHistoricoComprasPorMes")
            Return New Dictionary(Of Date, Decimal)
        End Try
    End Function

    ''' <summary>
    ''' Obtém histórico de vendas agrupado por mês
    ''' </summary>
    Public Function ObterHistoricoVendasPorMes(codigoProduto As String, dataInicio As Date, dataFim As Date) As Dictionary(Of Date, Decimal)
        Try
            LogErros.RegistrarInfo($"📉 Buscando histórico de vendas para produto: {codigoProduto}", "ObterHistoricoVendasPorMes")

            Dim resultado As New Dictionary(Of Date, Decimal)
            Dim tabela = ObterTabela("tblVendas")

            If tabela IsNot Nothing Then
                LogErros.RegistrarInfo($"✅ Tabela tblVendas encontrada com {tabela.ListRows.Count} linhas", "ObterHistoricoVendasPorMes")

                ' ✅ OTIMIZAÇÃO: Usar busca ultra-rápida para tabelas grandes (limiar reduzido)
                If tabela.ListRows.Count > 10000 Then
                    LogErros.RegistrarInfo("🚀 Usando busca ultra-rápida com fórmulas para tabela grande", "ObterHistoricoVendasPorMes")
                    Try
                        Return PowerQueryManagerOptimizedSearch.BuscarHistoricoRapido(tabela, codigoProduto, dataInicio, dataFim)
                    Catch searchEx As Exception
                        LogErros.RegistrarErro(searchEx, "ObterHistoricoVendasPorMes.BuscaRapida")
                        ' Fallback para método limitado
                    End Try
                End If

                ' Criar estrutura para todos os meses no período
                Dim dataAtual = New Date(dataInicio.Year, dataInicio.Month, 1)
                While dataAtual <= dataFim
                    resultado(dataAtual) = 0
                    dataAtual = dataAtual.AddMonths(1)
                End While
                LogErros.RegistrarInfo($"✅ Criados {resultado.Count} meses no período", "ObterHistoricoVendasPorMes")

                ' ✅ OTIMIZAÇÃO: Usar Range.Find para buscar mais eficientemente com limite
                Dim worksheet = tabela.Parent
                Dim codigoRange = tabela.ListColumns(1).Range
                Dim registrosProcessados = 0
                Dim maxRegistros = 1000 ' Limitar para evitar OutOfMemory
                
                Try
                    ' Buscar primeira ocorrência
                    Dim foundCell = codigoRange.Find(codigoProduto, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, 
                                                    LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlWhole)
                    
                    If foundCell IsNot Nothing Then
                        Dim firstAddress = foundCell.Address
                        Do
                            Try
                                If registrosProcessados >= maxRegistros Then Exit Do
                                
                                Dim rowIndex = foundCell.Row - tabela.Range.Row + 1
                                If rowIndex > 1 AndAlso rowIndex <= tabela.ListRows.Count + 1 Then
                                    Dim dataVenda = CDate(tabela.ListColumns(2).Range.Cells(rowIndex).Value)
                                    Dim quantidade = CDec(tabela.ListColumns(3).Range.Cells(rowIndex).Value)

                                    If dataVenda >= dataInicio AndAlso dataVenda <= dataFim Then
                                        Dim mesChave = New Date(dataVenda.Year, dataVenda.Month, 1)
                                        If resultado.ContainsKey(mesChave) Then
                                            resultado(mesChave) += quantidade
                                            registrosProcessados += 1
                                        End If
                                    End If
                                End If
                            Catch cellEx As Exception
                                ' Ignorar erro de célula específica
                            End Try
                            
                            ' Buscar próxima ocorrência
                            foundCell = codigoRange.FindNext(foundCell)
                        Loop While foundCell IsNot Nothing AndAlso foundCell.Address <> firstAddress
                    End If
                    
                Catch findEx As Exception
                    LogErros.RegistrarErro(findEx, "ObterHistoricoVendasPorMes.Find")
                End Try
                
                LogErros.RegistrarInfo($"✅ Vendas processadas: {registrosProcessados} registros para produto {codigoProduto}", "ObterHistoricoVendasPorMes")
            Else
                LogErros.RegistrarInfo("⚠️ Tabela tblVendas não encontrada", "ObterHistoricoVendasPorMes")
            End If

            Return resultado

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ObterHistoricoVendasPorMes")
            Return New Dictionary(Of Date, Decimal)
        End Try
    End Function

End Class
