' DatabaseConfig.vb
' Preparação para futura migração dos dados do Excel para banco de dados

Public Class DatabaseConfig

    ' Configurações de conexão (para uso futuro)
    Public Const CONNECTION_STRING_TEMPLATE As String = "Data Source={0};Initial Catalog={1};Integrated Security=True"
    Public Const DEFAULT_SERVER As String = "localhost\SQLEXPRESS"
    Public Const DEFAULT_DATABASE As String = "EstoqueDifemaq"

    ' Mapeamento de tabelas Excel -> BD
    Public Shared ReadOnly TABELA_MAPEAMENTO As New Dictionary(Of String, String) From {
        {ConfiguracaoApp.TABELA_PRODUTOS, "Produtos"},
        {ConfiguracaoApp.TABELA_ESTOQUE, "EstoqueAtual"},
        {ConfiguracaoApp.TABELA_COMPRAS, "MovimentoCompras"},
        {ConfiguracaoApp.TABELA_VENDAS, "MovimentoVendas"}
    }

    ' Estrutura das tabelas (para documentação e futura criação)
    Public Class TabelaProdutos
        Public Property Codigo As String
        Public Property Descricao As String
        Public Property Categoria As String
        Public Property Unidade As String
        Public Property PrecoVenda As Decimal
        Public Property PrecoCompra As Decimal
        Public Property FornecedorPrincipal As String
        Public Property Ativo As Boolean
        Public Property DataCadastro As DateTime
        Public Property DataAtualizacao As DateTime
    End Class

    Public Class TabelaEstoque
        Public Property Codigo As String
        Public Property Local As String
        Public Property Quantidade As Decimal
        Public Property QuantidadeMinima As Decimal
        Public Property QuantidadeMaxima As Decimal
        Public Property DataUltimaContagem As DateTime
        Public Property CustoMedio As Decimal
    End Class

    Public Class TabelaMovimentacao
        Public Property Id As Integer
        Public Property Codigo As String
        Public Property Data As DateTime
        Public Property Tipo As String ' "COMPRA" ou "VENDA"
        Public Property Quantidade As Decimal
        Public Property ValorUnitario As Decimal
        Public Property ValorTotal As Decimal
        Public Property Parceiro As String ' Fornecedor ou Cliente
        Public Property NumeroDocumento As String
        Public Property Observacoes As String
    End Class

    ' Métodos para validação de estrutura
    Public Shared Function ValidarEstruturaDados(dataTable As System.Data.DataTable, tipoTabela As String) As List(Of String)
        Dim erros As New List(Of String)()

        Select Case tipoTabela
            Case "Produtos"
                If Not dataTable.Columns.Contains("Codigo") Then erros.Add("Coluna 'Codigo' não encontrada")
                If Not dataTable.Columns.Contains("Descricao") Then erros.Add("Coluna 'Descricao' não encontrada")

            Case "Estoque"
                If Not dataTable.Columns.Contains("Codigo") Then erros.Add("Coluna 'Codigo' não encontrada")
                If Not dataTable.Columns.Contains("Local") Then erros.Add("Coluna 'Local' não encontrada")
                If Not dataTable.Columns.Contains("Quantidade") Then erros.Add("Coluna 'Quantidade' não encontrada")

            Case "Compras", "Vendas"
                If Not dataTable.Columns.Contains("Codigo") Then erros.Add("Coluna 'Codigo' não encontrada")
                If Not dataTable.Columns.Contains("Data") Then erros.Add("Coluna 'Data' não encontrada")
                If Not dataTable.Columns.Contains("Quantidade") Then erros.Add("Coluna 'Quantidade' não encontrada")
        End Select

        Return erros
    End Function

    ' Preparar dados para exportação/migração
    Public Shared Function PrepararDadosParaExportacao(dataTable As System.Data.DataTable) As System.Data.DataTable
        Dim dtExport As System.Data.DataTable = dataTable.Clone()

        ' Adicionar colunas de auditoria se não existirem
        If Not dtExport.Columns.Contains("DataImportacao") Then
            dtExport.Columns.Add("DataImportacao", GetType(DateTime))
        End If

        If Not dtExport.Columns.Contains("UsuarioImportacao") Then
            dtExport.Columns.Add("UsuarioImportacao", GetType(String))
        End If

        ' Copiar dados
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim newRow As System.Data.DataRow = dtExport.NewRow()

            For Each column As System.Data.DataColumn In dataTable.Columns
                newRow(column.ColumnName) = row(column.ColumnName)
            Next

            newRow("DataImportacao") = DateTime.Now
            newRow("UsuarioImportacao") = Environment.UserName

            dtExport.Rows.Add(newRow)
        Next

        Return dtExport
    End Function

End Class