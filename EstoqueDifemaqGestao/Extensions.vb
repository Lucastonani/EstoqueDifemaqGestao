Imports System.Runtime.CompilerServices
Imports System.Drawing
Imports System.Windows.Forms

Public Module Extensions
    <Extension()>
    Public Sub EstilizarDataGridView(dgv As DataGridView)
        Try
            With dgv
                ' Configurações gerais
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToResizeRows = False
                .ReadOnly = True
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .MultiSelect = False
                .AutoGenerateColumns = True
                .BorderStyle = BorderStyle.Fixed3D
                .EnableHeadersVisualStyles = False
                .AllowUserToOrderColumns = False

                ' Configurações visuais
                .BackgroundColor = Color.White
                .GridColor = Color.LightGray
                .RowHeadersVisible = False
                .ColumnHeadersHeight = 35
                .RowTemplate.Height = ConfiguracaoApp.ALTURA_LINHA_GRID

                ' Configurações de performance
                .VirtualMode = False
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                ' Estilo do cabeçalho
                With .ColumnHeadersDefaultCellStyle
                    .BackColor = ConfiguracaoApp.ObterCorHeader()
                    .ForeColor = Color.White
                    .Font = New Font("Segoe UI", 9, FontStyle.Bold)
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.False
                End With

                ' Estilo das células
                With .DefaultCellStyle
                    .BackColor = Color.White
                    .ForeColor = Color.Black
                    .Font = New Font("Segoe UI", 9)
                    .SelectionBackColor = ConfiguracaoApp.ObterCorSelecao()
                    .SelectionForeColor = Color.White
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .Padding = New Padding(5, 2, 5, 2)
                    .WrapMode = DataGridViewTriState.False
                End With

                ' Estilo das linhas alternadas
                With .AlternatingRowsDefaultCellStyle
                    .BackColor = ConfiguracaoApp.ObterCorAlternada()
                End With

                ' Configurações de scroll
                .ScrollBars = ScrollBars.Both
                .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                ' Configurações de eventos para melhor performance
                .SuspendLayout()
                .ResumeLayout(False)
            End With
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "Extensions.EstilizarDataGridView")
        End Try
    End Sub

    <Extension()>
    Public Sub ConfigurarColunas(dgv As DataGridView, ParamArray configuracoes As ColumnConfig())
        Try
            If configuracoes Is Nothing OrElse configuracoes.Length = 0 Then Return

            dgv.SuspendLayout()

            For Each config As ColumnConfig In configuracoes
                If dgv.Columns.Count > config.Index AndAlso config.Index >= 0 Then
                    With dgv.Columns(config.Index)
                        .HeaderText = config.HeaderText
                        .Width = config.Width
                        .Visible = config.Visible
                        .ReadOnly = config.ReadOnly

                        ' Configurar alinhamento
                        If config.Alignment.HasValue Then
                            .DefaultCellStyle.Alignment = config.Alignment.Value
                        End If

                        ' Configurar formato
                        If Not String.IsNullOrEmpty(config.Format) Then
                            .DefaultCellStyle.Format = config.Format
                        End If

                        ' Configurar largura mínima
                        If config.MinimumWidth.HasValue Then
                            .MinimumWidth = config.MinimumWidth.Value
                        End If

                        ' Configurar modo de auto-dimensionamento
                        If config.AutoSizeMode.HasValue Then
                            .AutoSizeMode = config.AutoSizeMode.Value
                        Else
                            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                        End If

                        ' Configurar modo de ordenação
                        If config.SortMode.HasValue Then
                            .SortMode = config.SortMode.Value
                        Else
                            .SortMode = DataGridViewColumnSortMode.Automatic
                        End If

                        ' Configurar resizable
                        .Resizable = DataGridViewTriState.True
                    End With
                End If
            Next

            dgv.ResumeLayout(True)

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "Extensions.ConfigurarColunas")
        Finally
            Try
                dgv.ResumeLayout(True)
            Catch
                ' Ignorar erro no ResumeLayout
            End Try
        End Try
    End Sub

    <Extension()>
    Public Sub OtimizarPerformance(dgv As DataGridView)
        Try
            With dgv
                .SuspendLayout()

                ' Configurações para melhor performance
                .RowHeadersVisible = False
                .AutoGenerateColumns = True
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToResizeRows = False
                .EnableHeadersVisualStyles = False
                .StandardTab = True

                ' Evitar redesenhos desnecessários
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                .ResumeLayout(True)
            End With
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "Extensions.OtimizarPerformance")
        End Try
    End Sub

    <Extension()>
    Public Sub ExportarParaCSV(dgv As DataGridView, caminhoArquivo As String)
        Try
            If dgv.DataSource Is Nothing Then
                Throw New InvalidOperationException("DataGridView não possui dados para exportar")
            End If

            Using writer As New IO.StreamWriter(caminhoArquivo, False, System.Text.Encoding.UTF8)
                ' Escrever cabeçalhos
                Dim headers As New List(Of String)
                For Each column As DataGridViewColumn In dgv.Columns
                    If column.Visible Then
                        headers.Add($"""{column.HeaderText}""")
                    End If
                Next
                writer.WriteLine(String.Join(",", headers))

                ' Escrever dados
                For Each row As DataGridViewRow In dgv.Rows
                    If Not row.IsNewRow Then
                        Dim values As New List(Of String)
                        For Each column As DataGridViewColumn In dgv.Columns
                            If column.Visible Then
                                Dim valor As String = row.Cells(column.Index).Value?.ToString() ?? ""
                                values.Add($"""{valor.Replace("""", """""")}""")
                            End If
                        Next
                        writer.WriteLine(String.Join(",", values))
                    End If
                Next
            End Using

            LogErros.RegistrarInfo($"Dados exportados para: {caminhoArquivo}", "Extensions.ExportarParaCSV")

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "Extensions.ExportarParaCSV")
            Throw New Exception($"Erro ao exportar dados: {ex.Message}")
        End Try
    End Sub

    <Extension()>
    Public Function ObterDadosSelecionados(dgv As DataGridView) As Dictionary(Of String, Object)
        Try
            Dim dados As New Dictionary(Of String, Object)

            If dgv.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = dgv.SelectedRows(0)

                For Each column As DataGridViewColumn In dgv.Columns
                    If column.Visible Then
                        dados.Add(column.HeaderText, selectedRow.Cells(column.Index).Value)
                    End If
                Next
            End If

            Return dados

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "Extensions.ObterDadosSelecionados")
            Return New Dictionary(Of String, Object)
        End Try
    End Function

    <Extension()>
    Public Sub AplicarFiltroRapido(dgv As DataGridView, filtro As String, colunaIndex As Integer)
        Try
            If dgv.DataSource Is Nothing OrElse TypeOf dgv.DataSource IsNot DataTable Then
                Return
            End If

            Dim dataTable As DataTable = CType(dgv.DataSource, DataTable)
            Dim dataView As DataView = dataTable.DefaultView

            If String.IsNullOrEmpty(filtro) Then
                dataView.RowFilter = ""
            Else
                If colunaIndex >= 0 AndAlso colunaIndex < dataTable.Columns.Count Then
                    Dim nomeColuna As String = dataTable.Columns(colunaIndex).ColumnName
                    dataView.RowFilter = $"[{nomeColuna}] LIKE '%{filtro.Replace("'", "''")}%'"
                End If
            End If

        Catch ex As Exception
            LogErros.RegistrarErro(ex, "Extensions.AplicarFiltroRapido")
        End Try
    End Sub

End Module