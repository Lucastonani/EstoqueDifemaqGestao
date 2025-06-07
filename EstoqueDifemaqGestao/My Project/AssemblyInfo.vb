Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' As informações gerais sobre um assembly são controladas por
' conjunto de atributos. Altere estes valores de atributo para modificar as informações
' associada a um assembly.

' Revise os valores dos atributos do assembly

<Assembly: AssemblyTitle("EstoqueDifemaqGestao")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("")> 
<Assembly: AssemblyProduct("EstoqueDifemaqGestao")> 
<Assembly: AssemblyCopyright("Copyright ©  2025")> 
<Assembly: AssemblyTrademark("")> 

' Definir ComVisible como false torna invisíveis os tipos neste assembly
' para componentes COM.  Caso precise acessar um tipo neste assembly a partir de 
' COM, defina o atributo ComVisible como true nesse tipo.
<Assembly: ComVisible(False)>

'O GUID a seguir será destinado à ID de typelib se este projeto for exposto para COM
<Assembly: Guid("9fb56ac0-e41a-4be7-b47d-48c27a89dbb1")> 

' As informações da versão de um assembly consistem nos quatro valores a seguir:
'
'      Versão Principal
'      Versão Secundária 
'      Número da Versão
'      Revisão
'
' É possível especificar todos os valores ou usar como padrão os Números da Versão e da Revisão 
' utilizando o "*" como mostrado abaixo:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
