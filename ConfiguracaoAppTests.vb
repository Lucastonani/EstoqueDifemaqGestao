Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Drawing

<TestClass>
Public Class ConfiguracaoAppTests

    <TestMethod>
    Public Sub ObterCorHeader_ReturnsExpectedColor()
        ' Arrange
        Dim expectedColor As Color = ColorTranslator.FromHtml(ConfiguracaoApp.COR_HEADER_GRID)

        ' Act
        Dim actualColor As Color = ConfiguracaoApp.ObterCorHeader()

        ' Assert
        Assert.AreEqual(expectedColor.ToArgb(), actualColor.ToArgb())
    End Sub

    <TestMethod>
    Public Sub ObterCorSelecao_ReturnsExpectedColor()
        ' Arrange
        Dim expectedColor As Color = ColorTranslator.FromHtml(ConfiguracaoApp.COR_SELECAO)

        ' Act
        Dim actualColor As Color = ConfiguracaoApp.ObterCorSelecao()

        ' Assert
        Assert.AreEqual(expectedColor.ToArgb(), actualColor.ToArgb())
    End Sub

    <TestMethod>
    Public Sub ObterCorAlternada_ReturnsExpectedColor()
        ' Arrange
        Dim expectedColor As Color = ColorTranslator.FromHtml(ConfiguracaoApp.COR_LINHA_ALTERNADA)

        ' Act
        Dim actualColor As Color = ConfiguracaoApp.ObterCorAlternada()

        ' Assert
        Assert.AreEqual(expectedColor.ToArgb(), actualColor.ToArgb())
    End Sub

End Class