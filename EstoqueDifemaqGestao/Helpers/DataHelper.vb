Imports Microsoft.Office.Interop.Excel

Public Class DataHelper

    Public Shared Function ConvertListObjectToDataTable(listObject As ListObject) As System.Data.DataTable
        Try
            Dim dataTable As New System.Data.DataTable()

            If listObject Is Nothing Then
                Return dataTable
            End If

            ' Criar colunas baseadas no cabeçalho
            Dim headerRange As Range = listObject.HeaderRowRange
            If headerRange IsNot Nothing Then
                For col As Integer = 1 To headerRange.Columns.Count
                    Dim cellValue = headerRange.Cells(1, col).Value
                    Dim columnName As String = If(cellValue IsNot Nothing, cellValue.ToString(), String.Format("Coluna{0}", col))
                    dataTable.Columns.Add(columnName)
                Next
            End If

            ' Adicionar dados do corpo da tabela
            Dim dataBodyRange As Range = listObject.DataBodyRange
            If dataBodyRange IsNot Nothing Then
                For row As Integer = 1 To dataBodyRange.Rows.Count
                    Dim dataRow As System.Data.DataRow = dataTable.NewRow()
                    For col As Integer = 1 To Math.Min(dataBodyRange.Columns.Count, dataTable.Columns.Count)
                        Dim cellValue = dataBodyRange.Cells(row, col).Value
                        dataRow(col - 1) = If(cellValue Is Nothing, "", cellValue.ToString())
                    Next
                    dataTable.Rows.Add(dataRow)
                Next
            End If

            Return dataTable

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "DataHelper.ConvertListObjectToDataTable")
            Throw New Exception(String.Format("Erro ao converter ListObject para DataTable: {0}", ex.Message))
        End Try
    End Function

    Public Shared Function ConvertRangeToDataTable(range As Range, Optional hasHeaders As Boolean = True) As System.Data.DataTable
        Try
            Dim dataTable As New System.Data.DataTable()

            If range Is Nothing OrElse range.Rows.Count = 0 Then
                Return dataTable
            End If

            Dim startRow As Integer = 1

            ' Criar colunas
            If hasHeaders Then
                For col As Integer = 1 To range.Columns.Count
                    Dim cellValue = range.Cells(1, col).Value
                    Dim columnName As String = If(cellValue IsNot Nothing, cellValue.ToString(), String.Format("Coluna{0}", col))
                    dataTable.Columns.Add(columnName)
                Next
                startRow = 2
            Else
                For col As Integer = 1 To range.Columns.Count
                    dataTable.Columns.Add(String.Format("Coluna{0}", col))
                Next
            End If

            ' Adicionar dados
            For row As Integer = startRow To range.Rows.Count
                Dim dataRow As System.Data.DataRow = dataTable.NewRow()
                For col As Integer = 1 To range.Columns.Count
                    Dim cellValue = range.Cells(row, col).Value
                    dataRow(col - 1) = If(cellValue Is Nothing, "", cellValue.ToString())
                Next
                dataTable.Rows.Add(dataRow)
            Next

            Return dataTable

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "DataHelper.ConvertRangeToDataTable")
            Throw New Exception(String.Format("Erro ao converter Range para DataTable: {0}", ex.Message))
        End Try
    End Function

    Public Shared Function FiltrarDataTable(dataTable As System.Data.DataTable, coluna As String, valor As String) As System.Data.DataTable
        Try
            If dataTable Is Nothing OrElse String.IsNullOrEmpty(coluna) Then
                Return New System.Data.DataTable()
            End If

            Dim tabelaFiltrada As System.Data.DataTable = dataTable.Clone()

            ' Verificar se a coluna existe
            If Not dataTable.Columns.Contains(coluna) Then
                Return tabelaFiltrada
            End If

            For Each row As System.Data.DataRow In dataTable.Rows
                Dim valorCelula As String = If(row(coluna) IsNot Nothing, row(coluna).ToString(), "")
                If valorCelula.Equals(valor, StringComparison.OrdinalIgnoreCase) Then
                    tabelaFiltrada.ImportRow(row)
                End If
            Next

            Return tabelaFiltrada

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "DataHelper.FiltrarDataTable")
            Throw New Exception(String.Format("Erro ao filtrar DataTable: {0}", ex.Message))
        End Try
    End Function

    Public Shared Function FiltrarDataTableMultiplosValores(dataTable As System.Data.DataTable, coluna As String, valores As String()) As System.Data.DataTable
        Try
            If dataTable Is Nothing OrElse String.IsNullOrEmpty(coluna) OrElse valores Is Nothing Then
                Return New System.Data.DataTable()
            End If

            Dim tabelaFiltrada As System.Data.DataTable = dataTable.Clone()

            If Not dataTable.Columns.Contains(coluna) Then
                Return tabelaFiltrada
            End If

            For Each row As System.Data.DataRow In dataTable.Rows
                Dim valorCelula As String = If(row(coluna) IsNot Nothing, row(coluna).ToString(), "")
                If valores.Any(Function(v) valorCelula.Equals(v, StringComparison.OrdinalIgnoreCase)) Then
                    tabelaFiltrada.ImportRow(row)
                End If
            Next

            Return tabelaFiltrada

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "DataHelper.FiltrarDataTableMultiplosValores")
            Throw New Exception(String.Format("Erro ao filtrar DataTable com múltiplos valores: {0}", ex.Message))
        End Try
    End Function

    ' Métodos adicionais melhorados

    Public Shared Function FiltrarDataTablePorTexto(dataTable As System.Data.DataTable, coluna As String, texto As String) As System.Data.DataTable
        Try
            If dataTable Is Nothing OrElse String.IsNullOrEmpty(coluna) OrElse String.IsNullOrEmpty(texto) Then
                Return If(dataTable IsNot Nothing, dataTable.Clone(), New System.Data.DataTable())
            End If

            Dim tabelaFiltrada As System.Data.DataTable = dataTable.Clone()

            If Not dataTable.Columns.Contains(coluna) Then
                Return tabelaFiltrada
            End If

            For Each row As System.Data.DataRow In dataTable.Rows
                Dim valorCelula As String = If(row(coluna) IsNot Nothing, row(coluna).ToString(), "")
                If valorCelula.IndexOf(texto, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    tabelaFiltrada.ImportRow(row)
                End If
            Next

            Return tabelaFiltrada

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "DataHelper.FiltrarDataTablePorTexto")
            Return If(dataTable IsNot Nothing, dataTable.Clone(), New System.Data.DataTable())
        End Try
    End Function

    Public Shared Function FiltrarDataTableMultiplasColunas(dataTable As System.Data.DataTable, filtroTexto As String) As System.Data.DataTable
        Try
            If dataTable Is Nothing OrElse String.IsNullOrEmpty(filtroTexto) Then
                Return If(dataTable IsNot Nothing, dataTable.Clone(), New System.Data.DataTable())
            End If

            Dim tabelaFiltrada As System.Data.DataTable = dataTable.Clone()

            For Each row As System.Data.DataRow In dataTable.Rows
                Dim incluirRow As Boolean = False

                ' Verificar em todas as colunas de texto
                For Each column As System.Data.DataColumn In dataTable.Columns
                    If column.DataType = GetType(String) Then
                        Dim valorCelula As String = If(row(column) IsNot Nothing, row(column).ToString(), "")
                        If valorCelula.IndexOf(filtroTexto, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            incluirRow = True
                            Exit For
                        End If
                    End If
                Next

                If incluirRow Then
                    tabelaFiltrada.ImportRow(row)
                End If
            Next

            Return tabelaFiltrada

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "DataHelper.FiltrarDataTableMultiplasColunas")
            Return If(dataTable IsNot Nothing, dataTable.Clone(), New System.Data.DataTable())
        End Try
    End Function

    Public Shared Function ValidarDataTable(dataTable As System.Data.DataTable) As Dictionary(Of String, Object)
        Try
            Dim resultado As New Dictionary(Of String, Object)()

            If dataTable Is Nothing Then
                resultado.Add("Valido", False)
                resultado.Add("Erro", "DataTable é nulo")
                Return resultado
            End If

            resultado.Add("Valido", True)
            resultado.Add("TotalLinhas", dataTable.Rows.Count)
            resultado.Add("TotalColunas", dataTable.Columns.Count)
            resultado.Add("TemDados", dataTable.Rows.Count > 0)
            resultado.Add("NomeColunas", dataTable.Columns.Cast(Of System.Data.DataColumn).Select(Function(c) c.ColumnName).ToArray())

            ' Verificar tipos de dados
            Dim tiposColunas As New Dictionary(Of String, String)()
            For Each column As System.Data.DataColumn In dataTable.Columns
                tiposColunas.Add(column.ColumnName, column.DataType.Name)
            Next
            resultado.Add("TiposColunas", tiposColunas)

            Return resultado

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "DataHelper.ValidarDataTable")
            Dim resultado As New Dictionary(Of String, Object)()
            resultado.Add("Valido", False)
            resultado.Add("Erro", ex.Message)
            Return resultado
        End Try
    End Function

    Public Shared Function ObterResumoDataTable(dataTable As System.Data.DataTable) As String
        Try
            If dataTable Is Nothing Then
                Return "DataTable é nulo"
            End If

            Dim resumo As New System.Text.StringBuilder()
            resumo.AppendLine(String.Format("Total de linhas: {0}", dataTable.Rows.Count))
            resumo.AppendLine(String.Format("Total de colunas: {0}", dataTable.Columns.Count))
            resumo.AppendLine("Colunas:")

            For Each column As System.Data.DataColumn In dataTable.Columns
                resumo.AppendLine(String.Format("  - {0} ({1})", column.ColumnName, column.DataType.Name))
            Next

            If dataTable.Rows.Count > 0 Then
                resumo.AppendLine("Primeira linha (amostra):")
                For i As Integer = 0 To Math.Min(4, dataTable.Columns.Count - 1)
                    Dim valor As String = If(dataTable.Rows(0)(i) IsNot Nothing, dataTable.Rows(0)(i).ToString(), "(vazio)")
                    resumo.AppendLine(String.Format("  {0}: {1}", dataTable.Columns(i).ColumnName, valor))
                Next
            End If

            Return resumo.ToString()

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "DataHelper.ObterResumoDataTable")
            Return String.Format("Erro ao gerar resumo: {0}", ex.Message)
        End Try
    End Function

    Public Shared Function LimparDataTable(dataTable As System.Data.DataTable) As System.Data.DataTable
        Try
            If dataTable Is Nothing Then
                Return New System.Data.DataTable()
            End If

            Dim tabelaLimpa As System.Data.DataTable = dataTable.Clone()

            For Each row As System.Data.DataRow In dataTable.Rows
                Dim rowVazia As Boolean = True

                ' Verificar se a linha tem algum dado não vazio
                For Each item As Object In row.ItemArray
                    If item IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(item.ToString()) Then
                        rowVazia = False
                        Exit For
                    End If
                Next

                ' Adicionar apenas linhas que não estão vazias
                If Not rowVazia Then
                    Dim novaRow As System.Data.DataRow = tabelaLimpa.NewRow()
                    For i As Integer = 0 To row.ItemArray.Length - 1
                        novaRow(i) = row(i)
                    Next
                    tabelaLimpa.Rows.Add(novaRow)
                End If
            Next

            Return tabelaLimpa

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "DataHelper.LimparDataTable")
            Return If(dataTable IsNot Nothing, dataTable.Clone(), New System.Data.DataTable())
        End Try
    End Function

End Class