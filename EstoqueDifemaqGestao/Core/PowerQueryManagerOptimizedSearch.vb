Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Classe auxiliar para buscas otimizadas em tabelas Excel grandes
''' Implementa estratégias específicas para lidar com grandes volumes de dados
''' </summary>
Public Class PowerQueryManagerOptimizedSearch

    ''' <summary>
    ''' Busca otimizada usando XLOOKUP ou INDEX/MATCH para encontrar dados de estoque
    ''' Muito mais eficiente que iterar linha por linha
    ''' </summary>
    Public Shared Function BuscarEstoqueRapido(tabela As ListObject, codigoProduto As String) As System.Data.DataTable
        Try
            LogErros.RegistrarInfo($"🚀 Busca rápida de estoque para produto: {codigoProduto}", "BuscarEstoqueRapido")
            
            Dim dt As New System.Data.DataTable()
            Dim worksheet = tabela.Parent
            
            ' Criar colunas no DataTable
            For Each col As ListColumn In tabela.ListColumns
                dt.Columns.Add(col.Name, GetType(Object))
            Next
            
            ' Usar fórmula XLOOKUP ou INDEX/MATCH para busca eficiente
            Dim searchRange = tabela.ListColumns(1).Range ' Coluna de código
            Dim dataRange = tabela.Range
            
            ' Criar range temporário para fórmula
            Dim tempCell = worksheet.Cells(1, tabela.Range.Columns.Count + 2)
            
            Try
                ' Tentar XLOOKUP primeiro (Excel 365)
                For colIndex = 1 To tabela.ListColumns.Count
                    Dim formula = $"=XLOOKUP(""{codigoProduto}"",{searchRange.Address},{tabela.ListColumns(colIndex).Range.Address})"
                    tempCell.Formula = formula
                    
                    Try
                        If tempCell.Value IsNot Nothing Then
                            ' Produto encontrado, buscar linha completa
                            Return BuscarLinhaCompleta(tabela, codigoProduto)
                        End If
                    Catch
                        ' Fórmula não funcionou, continuar
                    End Try
                Next
                
            Catch xlookupEx As Exception
                ' XLOOKUP não disponível, usar INDEX/MATCH
                LogErros.RegistrarInfo("XLOOKUP não disponível, usando INDEX/MATCH", "BuscarEstoqueRapido")
                
                Dim matchFormula = $"=MATCH(""{codigoProduto}"",{searchRange.Address},0)"
                tempCell.Formula = matchFormula
                
                Try
                    If tempCell.Value IsNot Nothing AndAlso IsNumeric(tempCell.Value) Then
                        Dim rowIndex = CInt(tempCell.Value)
                        Return ExtrairLinhaEspecifica(tabela, rowIndex)
                    End If
                Catch
                    ' Fórmula MATCH não funcionou
                End Try
                
            Finally
                ' Limpar célula temporária
                tempCell.Clear()
            End Try
            
            LogErros.RegistrarInfo($"❌ Produto {codigoProduto} não encontrado na busca rápida", "BuscarEstoqueRapido")
            Return dt
            
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManagerOptimizedSearch.BuscarEstoqueRapido")
            Return New System.Data.DataTable()
        End Try
    End Function
    
    Private Shared Function BuscarLinhaCompleta(tabela As ListObject, codigoProduto As String) As System.Data.DataTable
        Try
            Dim dt As New System.Data.DataTable()
            
            ' Criar colunas
            For Each col As ListColumn In tabela.ListColumns
                dt.Columns.Add(col.Name, GetType(Object))
            Next
            
            ' Usar AutoFilter para filtrar rapidamente
            Dim originalFilter = tabela.ShowAutoFilter
            tabela.ShowAutoFilter = True
            
            Try
                ' Aplicar filtro
                tabela.Range.AutoFilter(1, codigoProduto)
                
                ' Obter apenas células visíveis
                Dim visibleCells = tabela.Range.SpecialCells(XlCellType.xlCellTypeVisible)
                
                For Each area In visibleCells.Areas
                    If area.Rows.Count > 1 Then ' Pular cabeçalho
                        For rowIdx = 2 To area.Rows.Count
                            Dim dataRow = dt.NewRow()
                            For colIdx = 1 To tabela.ListColumns.Count
                                dataRow(colIdx - 1) = area.Cells(rowIdx, colIdx).Value
                            Next
                            dt.Rows.Add(dataRow)
                        Next
                    End If
                Next
                
            Finally
                ' Remover filtro
                tabela.Range.AutoFilter()
                tabela.ShowAutoFilter = originalFilter
            End Try
            
            Return dt
            
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "BuscarLinhaCompleta")
            Return New System.Data.DataTable()
        End Try
    End Function
    
    Private Shared Function ExtrairLinhaEspecifica(tabela As ListObject, rowIndex As Integer) As System.Data.DataTable
        Try
            Dim dt As New System.Data.DataTable()
            
            ' Criar colunas
            For Each col As ListColumn In tabela.ListColumns
                dt.Columns.Add(col.Name, GetType(Object))
            Next
            
            ' Extrair linha específica
            If rowIndex > 0 AndAlso rowIndex <= tabela.ListRows.Count Then
                Dim dataRow = dt.NewRow()
                Dim targetRow = tabela.ListRows(rowIndex)
                
                For i = 0 To tabela.ListColumns.Count - 1
                    dataRow(i) = targetRow.Range(i + 1).Value
                Next
                
                dt.Rows.Add(dataRow)
            End If
            
            Return dt
            
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ExtrairLinhaEspecifica")
            Return New System.Data.DataTable()
        End Try
    End Function
    
    ''' <summary>
    ''' Busca otimizada para histórico usando fórmulas de array
    ''' Mais eficiente para grandes volumes de dados
    ''' </summary>
    Public Shared Function BuscarHistoricoRapido(tabela As ListObject, codigoProduto As String, dataInicio As Date, dataFim As Date) As Dictionary(Of Date, Decimal)
        Try
            LogErros.RegistrarInfo($"🚀 Busca rápida de histórico para produto: {codigoProduto}", "BuscarHistoricoRapido")
            
            Dim resultado As New Dictionary(Of Date, Decimal)
            Dim worksheet = tabela.Parent
            
            ' Criar estrutura de meses
            Dim dataAtual = New Date(dataInicio.Year, dataInicio.Month, 1)
            While dataAtual <= dataFim
                resultado(dataAtual) = 0
                dataAtual = dataAtual.AddMonths(1)
            End While
            
            ' Usar SUMIFS para somar por critério (muito mais rápido)
            Dim codigoRange = tabela.ListColumns(1).Range.Address
            Dim dataRange = If(tabela.ListColumns.Count >= 2, tabela.ListColumns(2).Range.Address, "")
            Dim quantidadeRange = If(tabela.ListColumns.Count >= 3, tabela.ListColumns(3).Range.Address, "")
            
            If String.IsNullOrEmpty(dataRange) OrElse String.IsNullOrEmpty(quantidadeRange) Then
                LogErros.RegistrarInfo("⚠️ Estrutura de tabela inadequada para busca rápida", "BuscarHistoricoRapido")
                Return resultado
            End If
            
            ' Célula temporária para fórmulas
            Dim tempCell = worksheet.Cells(1, tabela.Range.Columns.Count + 10)
            
            Try
                For Each mesChave In resultado.Keys.ToList()
                    Dim inicioMes = mesChave
                    Dim fimMes = mesChave.AddMonths(1).AddDays(-1)
                    
                    ' Fórmula SUMIFS para somar quantidade onde código = produto E data está no mês
                    Dim formula = $"=SUMIFS({quantidadeRange},{codigoRange},""{codigoProduto}"",{dataRange},"">="" & DATE({inicioMes.Year},{inicioMes.Month},{inicioMes.Day}),{dataRange},""<="" & DATE({fimMes.Year},{fimMes.Month},{fimMes.Day}))"
                    
                    tempCell.Formula = formula
                    
                    Try
                        If tempCell.Value IsNot Nothing AndAlso IsNumeric(tempCell.Value) Then
                            resultado(mesChave) = CDec(tempCell.Value)
                        End If
                    Catch
                        ' Fórmula SUMIFS não funcionou para este mês
                    End Try
                Next
                
            Finally
                tempCell.Clear()
            End Try
            
            Dim totalRegistros = resultado.Values.Sum()
            LogErros.RegistrarInfo($"✅ Busca rápida concluída: {totalRegistros} registros encontrados", "BuscarHistoricoRapido")
            
            Return resultado
            
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "PowerQueryManagerOptimizedSearch.BuscarHistoricoRapido")
            Return New Dictionary(Of Date, Decimal)
        End Try
    End Function
    
End Class