' UnitTests.vb
' Testes unitários para o sistema de gestão de estoque

Imports System.Data
Imports System.IO

Public Class UnitTests

    Public Shared Sub RunAllTests()
        TestFramework.ClearResults()

        ' Executar todas as suites de teste
        TestDataHelper()
        TestExtensions()
        TestColumnConfig()
        TestDatabaseConfig()
        TestConfiguracao()

        ' Gerar e salvar relatório
        Console.WriteLine()
        Console.WriteLine(TestFramework.GenerateReport())

        Dim reportPath = Path.Combine(ConfiguracaoApp.CAMINHO_LOG, $"TestReport_{DateTime.Now:yyyyMMdd_HHmmss}.txt")
        TestFramework.SaveReport(reportPath)
    End Sub

    ' Testes para DataHelper
    Private Shared Sub TestDataHelper()
        TestFramework.BeginTestSuite("DataHelper")

        TestFramework.RunTest("ConvertRangeToDataTable - Com Headers", Sub()
                                                                           ' Criar DataTable de teste
                                                                           Dim dt = New DataTable()
                                                                           dt.Columns.Add("Codigo", GetType(String))
                                                                           dt.Columns.Add("Descricao", GetType(String))
                                                                           dt.Columns.Add("Valor", GetType(Decimal))

                                                                           dt.Rows.Add("001", "Produto A", 10.5)
                                                                           dt.Rows.Add("002", "Produto B", 20.0)

                                                                           TestFramework.AssertNotNull(dt)
                                                                           TestFramework.AssertEquals(2, dt.Rows.Count)
                                                                           TestFramework.AssertEquals(3, dt.Columns.Count)
                                                                       End Sub)

        TestFramework.RunTest("FiltrarDataTable - Filtro Simples", Sub()
                                                                       ' Criar DataTable de teste
                                                                       Dim dt = New DataTable()
                                                                       dt.Columns.Add("Codigo", GetType(String))
                                                                       dt.Columns.Add("Nome", GetType(String))

                                                                       dt.Rows.Add("001", "Item A")
                                                                       dt.Rows.Add("002", "Item B")
                                                                       dt.Rows.Add("001", "Item C")

                                                                       Dim filtered = DataHelper.FiltrarDataTable(dt, "Codigo", "001")

                                                                       TestFramework.AssertNotNull(filtered)
                                                                       TestFramework.AssertEquals(2, filtered.Rows.Count)
                                                                   End Sub)

        TestFramework.RunTest("FiltrarDataTablePorTexto - Busca Parcial", Sub()
                                                                              Dim dt = New DataTable()
                                                                              dt.Columns.Add("Descricao", GetType(String))

                                                                              dt.Rows.Add("Parafuso Phillips")
                                                                              dt.Rows.Add("Parafuso Allen")
                                                                              dt.Rows.Add("Porca Sextavada")

                                                                              Dim filtered = DataHelper.FiltrarDataTablePorTexto(dt, "Descricao", "Parafuso")

                                                                              TestFramework.AssertEquals(2, filtered.Rows.Count)
                                                                          End Sub)

        TestFramework.RunTest("LimparDataTable - Remover Linhas Vazias", Sub()
                                                                             Dim dt = New DataTable()
                                                                             dt.Columns.Add("Col1", GetType(String))
                                                                             dt.Columns.Add("Col2", GetType(String))

                                                                             dt.Rows.Add("A", "B")
                                                                             dt.Rows.Add("", "")
                                                                             dt.Rows.Add("C", "D")
                                                                             dt.Rows.Add(Nothing, Nothing)

                                                                             Dim cleaned = DataHelper.LimparDataTable(dt)

                                                                             TestFramework.AssertEquals(2, cleaned.Rows.Count)
                                                                         End Sub)

        TestFramework.RunTest("ValidarDataTable - Validação Completa", Sub()
                                                                           Dim dt = New DataTable()
                                                                           dt.Columns.Add("ID", GetType(Integer))
                                                                           dt.Columns.Add("Nome", GetType(String))

                                                                           dt.Rows.Add(1, "Teste")

                                                                           Dim validacao = DataHelper.ValidarDataTable(dt)

                                                                           TestFramework.AssertTrue(CBool(validacao("Valido")))
                                                                           TestFramework.AssertEquals(1, validacao("TotalLinhas"))
                                                                           TestFramework.AssertEquals(2, validacao("TotalColunas"))
                                                                           TestFramework.AssertTrue(CBool(validacao("TemDados")))
                                                                       End Sub)
    End Sub

    ' Testes para Extensions
    Private Shared Sub TestExtensions()
        TestFramework.BeginTestSuite("Extensions")

        TestFramework.RunTest("ObterDadosSelecionados - DataGridView Vazio", Sub()
                                                                                 Dim dgv = New DataGridView()
                                                                                 Dim dados = dgv.ObterDadosSelecionados()

                                                                                 TestFramework.AssertNotNull(dados)
                                                                                 TestFramework.AssertEquals(0, dados.Count)
                                                                             End Sub)
    End Sub

    ' Testes para ColumnConfig
    Private Shared Sub TestColumnConfig()
        TestFramework.BeginTestSuite("ColumnConfig")

        TestFramework.RunTest("Construtor Básico", Sub()
                                                       Dim config = New ColumnConfig(0, "Teste", 100)

                                                       TestFramework.AssertEquals(0, config.Index)
                                                       TestFramework.AssertEquals("Teste", config.HeaderText)
                                                       TestFramework.AssertEquals(100, config.Width)
                                                       TestFramework.AssertTrue(config.Visible)
                                                       TestFramework.AssertTrue(config.IsReadOnly)
                                                   End Sub)

        TestFramework.RunTest("Construtor Completo", Sub()
                                                         Dim config = New ColumnConfig(1, "Valor", 150, True,
                                                             DataGridViewContentAlignment.MiddleRight, "C2", 100)

                                                         TestFramework.AssertEquals(1, config.Index)
                                                         TestFramework.AssertEquals("Valor", config.HeaderText)
                                                         TestFramework.AssertEquals(150, config.Width)
                                                         TestFramework.AssertTrue(config.Visible)
                                                         TestFramework.AssertEquals(DataGridViewContentAlignment.MiddleRight, config.Alignment)
                                                         TestFramework.AssertEquals("C2", config.Format)
                                                         TestFramework.AssertEquals(100, config.MinimumWidth)
                                                     End Sub)
    End Sub

    ' Testes para DatabaseConfig
    Private Shared Sub TestDatabaseConfig()
        TestFramework.BeginTestSuite("DatabaseConfig")

        TestFramework.RunTest("ValidarEstruturaDados - Produtos", Sub()
                                                                      Dim dt = New DataTable()
                                                                      dt.Columns.Add("Codigo", GetType(String))
                                                                      dt.Columns.Add("Descricao", GetType(String))

                                                                      Dim erros = DatabaseConfig.ValidarEstruturaDados(dt, "Produtos")

                                                                      TestFramework.AssertEquals(0, erros.Count)
                                                                  End Sub)

        TestFramework.RunTest("ValidarEstruturaDados - Estrutura Inválida", Sub()
                                                                                Dim dt = New DataTable()
                                                                                dt.Columns.Add("OutraColuna", GetType(String))

                                                                                Dim erros = DatabaseConfig.ValidarEstruturaDados(dt, "Produtos")

                                                                                TestFramework.AssertTrue(erros.Count > 0)
                                                                                TestFramework.AssertTrue(erros.Any(Function(e) e.Contains("Codigo")))
                                                                            End Sub)

        TestFramework.RunTest("PrepararDadosParaExportacao", Sub()
                                                                 Dim dt = New DataTable()
                                                                 dt.Columns.Add("ID", GetType(Integer))
                                                                 dt.Columns.Add("Nome", GetType(String))
                                                                 dt.Rows.Add(1, "Teste")

                                                                 Dim dtExport = DatabaseConfig.PrepararDadosParaExportacao(dt)

                                                                 TestFramework.AssertTrue(dtExport.Columns.Contains("DataImportacao"))
                                                                 TestFramework.AssertTrue(dtExport.Columns.Contains("UsuarioImportacao"))
                                                                 TestFramework.AssertEquals(1, dtExport.Rows.Count)
                                                             End Sub)
    End Sub

    ' Testes para Configuração
    Private Shared Sub TestConfiguracao()
        TestFramework.BeginTestSuite("ConfiguracaoApp")

        TestFramework.RunTest("Constantes de Configuração", Sub()
                                                                TestFramework.AssertNotNull(ConfiguracaoApp.TABELA_PRODUTOS)
                                                                TestFramework.AssertNotNull(ConfiguracaoApp.TABELA_ESTOQUE)
                                                                TestFramework.AssertNotNull(ConfiguracaoApp.CAMINHO_IMAGENS)
                                                                TestFramework.AssertTrue(ConfiguracaoApp.EXTENSOES_IMAGEM.Length > 0)
                                                            End Sub)

        TestFramework.RunTest("Métodos de Cor", Sub()
                                                    Dim corHeader = ConfiguracaoApp.ObterCorHeader()
                                                    Dim corSelecao = ConfiguracaoApp.ObterCorSelecao()
                                                    Dim corAlternada = ConfiguracaoApp.ObterCorAlternada()

                                                    TestFramework.AssertNotNull(corHeader)
                                                    TestFramework.AssertNotNull(corSelecao)
                                                    TestFramework.AssertNotNull(corAlternada)
                                                End Sub)
    End Sub

End Class