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
            Dim dt As New System.Data.DataTable()
            Dim tabela = ObterTabela("tblEstoqueVisao")

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
            Dim lojas As New HashSet(Of String)
            Dim tabela = ObterTabela("tblEstoqueVisao")

            If tabela IsNot Nothing Then
                ' Assumindo que a coluna Loja é a segunda coluna (índice 2)
                For Each row As ListRow In tabela.ListRows
                    Dim loja = row.Range(2).Value?.ToString()
                    If Not String.IsNullOrEmpty(loja) Then
                        lojas.Add(loja)
                    End If
                Next
            End If

            Return lojas.ToList()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ObterLojasDistintas")
            Return New List(Of String)
        End Try
    End Function

    ''' <summary>
    ''' Obtém histórico de compras agrupado por mês
    ''' </summary>
    Public Function ObterHistoricoComprasPorMes(codigoProduto As String, dataInicio As Date, dataFim As Date) As Dictionary(Of Date, Decimal)
        Try
            Dim resultado As New Dictionary(Of Date, Decimal)
            Dim tabela = ObterTabela("tblCompras")

            If tabela IsNot Nothing Then
                ' Criar estrutura para todos os meses no período
                Dim dataAtual = New Date(dataInicio.Year, dataInicio.Month, 1)
                While dataAtual <= dataFim
                    resultado(dataAtual) = 0
                    dataAtual = dataAtual.AddMonths(1)
                End While

                ' Processar dados (assumindo colunas: Código, Data, Quantidade)
                For Each row As ListRow In tabela.ListRows
                    Try
                        Dim codigo = row.Range(1).Value?.ToString()
                        If codigo = codigoProduto Then
                            Dim dataCompra = CDate(row.Range(2).Value)
                            Dim quantidade = CDec(row.Range(3).Value)

                            If dataCompra >= dataInicio AndAlso dataCompra <= dataFim Then
                                Dim mesChave = New Date(dataCompra.Year, dataCompra.Month, 1)
                                If resultado.ContainsKey(mesChave) Then
                                    resultado(mesChave) += quantidade
                                End If
                            End If
                        End If
                    Catch
                        ' Ignorar linhas com erro
                    End Try
                Next
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
            Dim resultado As New Dictionary(Of Date, Decimal)
            Dim tabela = ObterTabela("tblVendas")

            If tabela IsNot Nothing Then
                ' Criar estrutura para todos os meses no período
                Dim dataAtual = New Date(dataInicio.Year, dataInicio.Month, 1)
                While dataAtual <= dataFim
                    resultado(dataAtual) = 0
                    dataAtual = dataAtual.AddMonths(1)
                End While

                ' Processar dados (assumindo colunas: Código, Data, Quantidade)
                For Each row As ListRow In tabela.ListRows
                    Try
                        Dim codigo = row.Range(1).Value?.ToString()
                        If codigo = codigoProduto Then
                            Dim dataVenda = CDate(row.Range(2).Value)
                            Dim quantidade = CDec(row.Range(3).Value)

                            If dataVenda >= dataInicio AndAlso dataVenda <= dataFim Then
                                Dim mesChave = New Date(dataVenda.Year, dataVenda.Month, 1)
                                If resultado.ContainsKey(mesChave) Then
                                    resultado(mesChave) += quantidade
                                End If
                            End If
                        End If
                    Catch
                        ' Ignorar linhas com erro
                    End Try
                Next
            End If

            Return resultado

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManager.ObterHistoricoVendasPorMes")
            Return New Dictionary(Of Date, Decimal)
        End Try
    End Function

End Class
