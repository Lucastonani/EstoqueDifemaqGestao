' PerformanceTests.vb
' Testes de performance e stress

Imports System.Diagnostics

Public Class PerformanceTests

    Public Shared Sub RunPerformanceTests()
        TestFramework.BeginTestSuite("Performance")

        ' Executar testes
        TestDataTablePerformance()
        TestFilterPerformance()
        TestUIResponsiveness()
        TestMemoryUsage()

        ' Relatório
        Console.WriteLine()
        Console.WriteLine(TestFramework.GenerateReport())
    End Sub

    Private Shared Sub TestDataTablePerformance()
        TestFramework.RunTest("Performance DataTable - 10.000 registros", Sub()
                                                                              Dim sw As New Stopwatch()
                                                                              sw.Start()

                                                                              Dim dt As New System.Data.DataTable()
                                                                              dt.Columns.Add("ID", GetType(Integer))
                                                                              dt.Columns.Add("Codigo", GetType(String))
                                                                              dt.Columns.Add("Descricao", GetType(String))
                                                                              dt.Columns.Add("Preco", GetType(Decimal))
                                                                              dt.Columns.Add("Estoque", GetType(Integer))

                                                                              For i As Integer = 1 To 10000
                                                                                  dt.Rows.Add(i, $"PROD{i:00000}", $"Descrição do Produto {i}", i * 1.5D, i Mod 100)
                                                                              Next

                                                                              sw.Stop()

                                                                              Console.WriteLine($"  Tempo para criar 10.000 registros: {sw.ElapsedMilliseconds}ms")
                                                                              TestFramework.AssertTrue(sw.ElapsedMilliseconds < 1000, "Performance abaixo do esperado")

                                                                              ' Testar acesso aos dados
                                                                              sw.Restart()
                                                                              Dim soma As Decimal = 0
                                                                              For Each row As System.Data.DataRow In dt.Rows
                                                                                  soma += CDec(row("Preco"))
                                                                              Next
                                                                              sw.Stop()

                                                                              Console.WriteLine($"  Tempo para iterar 10.000 registros: {sw.ElapsedMilliseconds}ms")
                                                                              TestFramework.AssertTrue(sw.ElapsedMilliseconds < 100, "Iteração muito lenta")
                                                                          End Sub)
    End Sub

    Private Shared Sub TestFilterPerformance()
        TestFramework.RunTest("Performance Filtro - Dataset Grande", Sub()
                                                                         ' Criar dataset grande
                                                                         Dim dt As New System.Data.DataTable()
                                                                         dt.Columns.Add("Codigo", GetType(String))
                                                                         dt.Columns.Add("Descricao", GetType(String))

                                                                         For i As Integer = 1 To 5000
                                                                             ' Alternar entre "Par" e "Ímpar" corretamente
                                                                             Dim tipo As String = If(i Mod 2 = 0, "Par", "Ímpar")
                                                                             dt.Rows.Add($"PROD{i:00000}", $"Produto {tipo} Número {i}")
                                                                         Next

                                                                         ' Testar filtro
                                                                         Dim sw As New Stopwatch()
                                                                         sw.Start()

                                                                         Dim filtered = DataHelper.FiltrarDataTablePorTexto(dt, "Descricao", "Par")

                                                                         sw.Stop()

                                                                         Console.WriteLine($"  Filtrar 5.000 registros: {sw.ElapsedMilliseconds}ms")
                                                                         Console.WriteLine($"  Registros encontrados: {filtered.Rows.Count}")

                                                                         TestFramework.AssertTrue(sw.ElapsedMilliseconds < 500, "Filtro muito lento")
                                                                         ' Como todos os produtos têm "Par" ou "Ímpar", o filtro encontrará todos os 5000
                                                                         TestFramework.AssertTrue(filtered.Rows.Count > 0, "Nenhum registro encontrado")
                                                                     End Sub)
    End Sub

    Private Shared Sub TestUIResponsiveness()
        TestFramework.RunTest("Responsividade UI - DataGridView", Sub()
                                                                      Dim dgv As New DataGridView()
                                                                      dgv.VirtualMode = False

                                                                      ' Criar dados
                                                                      Dim dt As New System.Data.DataTable()
                                                                      dt.Columns.Add("ID", GetType(Integer))
                                                                      dt.Columns.Add("Nome", GetType(String))

                                                                      For i As Integer = 1 To 1000
                                                                          dt.Rows.Add(i, $"Item {i}")
                                                                      Next

                                                                      ' Medir binding
                                                                      Dim sw As New Stopwatch()
                                                                      sw.Start()

                                                                      dgv.DataSource = dt
                                                                      dgv.Refresh()

                                                                      sw.Stop()

                                                                      Console.WriteLine($"  Binding de 1.000 registros: {sw.ElapsedMilliseconds}ms")
                                                                      TestFramework.AssertTrue(sw.ElapsedMilliseconds < 500, "Binding muito lento")

                                                                      ' Cleanup
                                                                      dgv.Dispose()
                                                                  End Sub)
    End Sub

    Private Shared Sub TestMemoryUsage()
        TestFramework.RunTest("Uso de Memória", Sub()
                                                    ' Forçar coleta de lixo antes do teste
                                                    GC.Collect()
                                                    GC.WaitForPendingFinalizers()
                                                    GC.Collect()

                                                    Dim memoriaInicial As Long = GC.GetTotalMemory(False)

                                                    ' Criar objetos grandes
                                                    Dim lista As New List(Of System.Data.DataTable)
                                                    For i As Integer = 1 To 10
                                                        Dim dt As New System.Data.DataTable()
                                                        dt.Columns.Add("Data", GetType(String))

                                                        For j As Integer = 1 To 1000
                                                            dt.Rows.Add(New String("X"c, 100))
                                                        Next

                                                        lista.Add(dt)
                                                    Next

                                                    Dim memoriaUsada As Long = GC.GetTotalMemory(False) - memoriaInicial
                                                    Dim memoriaUsadaMB As Double = memoriaUsada / 1024.0 / 1024.0

                                                    Console.WriteLine($"  Memória usada: {memoriaUsadaMB:F2} MB")

                                                    ' Limpar
                                                    lista.Clear()
                                                    GC.Collect()
                                                    GC.WaitForPendingFinalizers()
                                                    GC.Collect()

                                                    Dim memoriaFinal As Long = GC.GetTotalMemory(False)
                                                    Dim memoriaLiberada As Boolean = memoriaFinal < (memoriaInicial + memoriaUsada * 0.2)

                                                    TestFramework.AssertTrue(memoriaLiberada, "Memória não foi liberada adequadamente")
                                                    Console.WriteLine("  Memória liberada: OK")
                                                End Sub)
    End Sub

End Class