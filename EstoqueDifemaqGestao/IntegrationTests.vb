' IntegrationTests.vb
' Testes de integração com Excel e Power Query

Imports Excel = Microsoft.Office.Interop.Excel

Public Class IntegrationTests

    Public Shared Sub RunIntegrationTests()
        TestFramework.BeginTestSuite("Integração Excel")

        ' Testes de integração
        TestExcelConnection()
        TestPowerQueryTables()
        TestImageDirectory()
        TestDataLoading()

        ' Gerar relatório
        Console.WriteLine()
        Console.WriteLine(TestFramework.GenerateReport())
    End Sub

    Private Shared Sub TestExcelConnection()
        TestFramework.RunTest("Conexão com Excel", Sub()
                                                       Dim excel As Object = Nothing
                                                       Try
                                                           excel = CreateObject("Excel.Application")
                                                           TestFramework.AssertNotNull(excel, "Excel não pôde ser criado")

                                                           ' Verificar versão
                                                           Dim version As String = excel.Version.ToString()
                                                           TestFramework.AssertNotNull(version, "Versão do Excel não disponível")
                                                           Console.WriteLine($"  Excel versão: {version}")

                                                       Finally
                                                           If excel IsNot Nothing Then
                                                               excel.Quit()
                                                               System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
                                                           End If
                                                       End Try
                                                   End Sub)
    End Sub

    Private Shared Sub TestPowerQueryTables()
        TestFramework.RunTest("Verificar Tabelas Power Query", Sub()
                                                                   ' Este teste precisa ser executado com o workbook aberto
                                                                   If Globals.ThisWorkbook Is Nothing Then
                                                                       Console.WriteLine("  AVISO: Teste requer workbook aberto")
                                                                       Return
                                                                   End If

                                                                   Try
                                                                       Dim thisWb As ThisWorkbook = CType(Globals.ThisWorkbook, ThisWorkbook)
                                                                       Dim wb = thisWb.ObterWorkbook()

                                                                       If wb IsNot Nothing Then
                                                                           Dim pqManager As New PowerQueryManager(wb)
                                                                           Dim tabelas = pqManager.ListarTabelas()

                                                                           TestFramework.AssertTrue(tabelas.Count > 0, "Nenhuma tabela encontrada")
                                                                           Console.WriteLine($"  Tabelas encontradas: {tabelas.Count}")

                                                                           ' Verificar tabelas essenciais
                                                                           Dim tabelasEssenciais As String() = {
                                                                               ConfiguracaoApp.TABELA_PRODUTOS,
                                                                               ConfiguracaoApp.TABELA_ESTOQUE,
                                                                               ConfiguracaoApp.TABELA_COMPRAS,
                                                                               ConfiguracaoApp.TABELA_VENDAS
                                                                           }

                                                                           For Each tabelaNome In tabelasEssenciais
                                                                               Dim encontrada = tabelas.Any(Function(t) t.Contains(tabelaNome))
                                                                               TestFramework.AssertTrue(encontrada, $"Tabela '{tabelaNome}' não encontrada")
                                                                           Next
                                                                       Else
                                                                           Console.WriteLine("  AVISO: Workbook não disponível")
                                                                       End If

                                                                   Catch ex As Exception
                                                                       Console.WriteLine($"  ERRO: {ex.Message}")
                                                                       Throw
                                                                   End Try
                                                               End Sub)
    End Sub

    Private Shared Sub TestImageDirectory()
        TestFramework.RunTest("Diretório de Imagens", Sub()
                                                          ' Verificar se existe
                                                          Dim exists = System.IO.Directory.Exists(ConfiguracaoApp.CAMINHO_IMAGENS)

                                                          If Not exists Then
                                                              ' Tentar criar
                                                              Try
                                                                  System.IO.Directory.CreateDirectory(ConfiguracaoApp.CAMINHO_IMAGENS)
                                                                  exists = System.IO.Directory.Exists(ConfiguracaoApp.CAMINHO_IMAGENS)
                                                                  Console.WriteLine($"  Diretório criado: {ConfiguracaoApp.CAMINHO_IMAGENS}")
                                                              Catch ex As Exception
                                                                  Console.WriteLine($"  Erro ao criar diretório: {ex.Message}")
                                                              End Try
                                                          End If

                                                          TestFramework.AssertTrue(exists, "Diretório de imagens não existe")

                                                          ' Verificar permissões
                                                          If exists Then
                                                              Try
                                                                  Dim testFile = System.IO.Path.Combine(ConfiguracaoApp.CAMINHO_IMAGENS, "test_permission.tmp")
                                                                  System.IO.File.WriteAllText(testFile, "test")
                                                                  System.IO.File.Delete(testFile)
                                                                  Console.WriteLine("  Permissões de escrita: OK")
                                                              Catch ex As Exception
                                                                  TestFramework.AssertTrue(False, $"Sem permissão de escrita: {ex.Message}")
                                                              End Try
                                                          End If
                                                      End Sub)
    End Sub

    Private Shared Sub TestDataLoading()
        TestFramework.RunTest("Carregamento de Dados Simulado", Sub()
                                                                    ' Simular carregamento de dados
                                                                    Dim dt As New System.Data.DataTable()
                                                                    dt.Columns.Add("Codigo", GetType(String))
                                                                    dt.Columns.Add("Descricao", GetType(String))
                                                                    dt.Columns.Add("Preco", GetType(Decimal))

                                                                    ' Adicionar dados de teste
                                                                    For i As Integer = 1 To 100
                                                                        dt.Rows.Add($"PROD{i:000}", $"Produto Teste {i}", i * 10.5D)
                                                                    Next

                                                                    TestFramework.AssertEquals(100, dt.Rows.Count)

                                                                    ' Testar filtro
                                                                    Dim filtered = DataHelper.FiltrarDataTablePorTexto(dt, "Descricao", "Teste 1")
                                                                    TestFramework.AssertTrue(filtered.Rows.Count >= 10, "Filtro não funcionou corretamente")

                                                                    Console.WriteLine($"  Dados carregados: {dt.Rows.Count} registros")
                                                                    Console.WriteLine($"  Dados filtrados: {filtered.Rows.Count} registros")
                                                                End Sub)
    End Sub

End Class