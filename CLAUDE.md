# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Test Commands

### Building the Project
```bash
# Build the main project (Visual Studio required)
msbuild EstoqueDifemaqGestao.sln /p:Configuration=Debug
msbuild EstoqueDifemaqGestao.sln /p:Configuration=Release

# Build only the main project
msbuild EstoqueDifemaqGestao/EstoqueDifemaqGestao.vbproj

# Build test project
msbuild MeuProjeto.Tests/MeuProjeto.Tests.vbproj
```

### Running Tests
```bash
# Run unit tests using MSTest
vstest.console.exe MeuProjeto.Tests/bin/Debug/MeuProjeto.Tests.dll

# Run tests with Visual Studio Test Explorer integration
dotnet test MeuProjeto.Tests/MeuProjeto.Tests.vbproj

# Run custom test framework (built into application)
# Tests are executed via TestFramework.vb and TestRunner.vb classes
```

### Debugging
```bash
# Debug through Visual Studio (Excel add-in)
# The project launches Excel with the add-in loaded automatically
# Debug executable: C:\Program Files\Microsoft Office\root\Office16\excel.exe
# Debug arguments: /x "Tratamento dos dados.xlsx"
```

## Architecture Overview

### Core System Design
This is a **VB.NET Excel Add-in** built as a VSTO (Visual Studio Tools for Office) project that provides inventory management capabilities integrated directly into Excel.

### Key Components

#### 1. Data Management Layer
- **PowerQueryManager.vb**: Core data engine with caching, retry logic, and Excel integration
- **PowerQueryManagerOtimizado.vb**: Async performance-optimized version with selective updates
- **PowerQueryManagerExtensions.vb**: Domain-specific business logic and singleton pattern implementation

#### 2. Configuration Management
- **ConfiguracaoApp.vb**: Centralized constants for table names, paths, timeouts, UI colors
- **DatabaseConfig.vb**: Database connection settings and migration path

#### 3. UI Layer
- **MainForm.vb**: Primary application interface
- **UcReposicaoEstoque.vb**: Stock replenishment user control
- **Sheet classes**: Excel worksheet integration (SheetEstoqueVisao, SheetProdutos, etc.)

#### 4. Infrastructure
- **LogErros.vb**: Thread-safe file-based logging with rotation
- **TestFramework.vb**: Custom testing infrastructure with comprehensive assertions
- **DataHelper.vb**: Data processing utilities and Excel interop helpers

### Excel Integration Pattern
The application follows a **Sheet-Centric Architecture**:
- Each business domain has dedicated Excel sheets (Produtos, Estoque, Compras, Vendas)
- Power Query handles data refresh and synchronization
- VSTO provides rich UI controls and business logic
- Host document: "Tratamento dos dados.xlsx"

### Performance Characteristics
- **Intelligent Caching**: 5-10 minute cache timeouts with validation
- **Selective Updates**: Optimized manager only updates essential connections
- **Excel State Management**: Disables alerts/screen updating during operations
- **Async Operations**: Non-blocking data refresh in optimized version

## Development Patterns

### Error Handling
Always use the established logging pattern:
```vb
Try
    ' Business logic here
    LogErros.RegistrarInfo("Operation started", "ClassName.MethodName")
    ' ... operation ...
    LogErros.RegistrarInfo("Operation completed", "ClassName.MethodName")
Catch ex As Exception
    LogErros.RegistrarErro(ex, "ClassName.MethodName")
    ' Handle appropriately (return False, throw, etc.)
End Try
```

### Cache Management
Follow the cache pattern established in PowerQueryManager:
- Use Dictionary(Of String, Object) for cache storage
- Implement timeout validation with DateTime comparison
- Invalidate cache after data modifications

### Excel Interop Best Practices
```vb
' Always configure Excel for performance during operations
Private Function ConfigurarExcel() As Dictionary(Of String, Object)
    Dim estadoAnterior As New Dictionary(Of String, Object)
    With Application
        estadoAnterior("DisplayAlerts") = .DisplayAlerts
        estadoAnterior("ScreenUpdating") = .ScreenUpdating
        estadoAnterior("EnableEvents") = .EnableEvents
        
        .DisplayAlerts = False
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    Return estadoAnterior
End Function
```

### Testing Strategy
Use the custom TestFramework for component testing:
```vb
TestFramework.RunTest("Test Description", Sub()
    ' Arrange
    ' Act  
    ' Assert using TestFramework.Assert* methods
End Sub)
```

## Key Constants & Configuration

### Table Names (from ConfiguracaoApp.vb)
- `TABELA_PRODUTOS`: "tblProdutos"
- `TABELA_ESTOQUE`: "tblEstoque" 
- `TABELA_COMPRAS`: "tblCompras"
- `TABELA_VENDAS`: "tblVendas"

### Performance Settings
- Cache timeout: 5 minutes (standard), 10 minutes (optimized)
- Retry attempts: 2 maximum
- Power Query timeout: 60 seconds (standard), 15 seconds (optimized)

### File Paths
- Images: "C:\ImagesEstoque\"
- Logs: "C:\Logs\GestaoEstoque\"

## Coding Standards

Follow the comprehensive coding standards defined in `Docs/CodeStandards/CODING_STANDARDS.md`:

### Naming Conventions
- **Classes**: PascalCase (`PowerQueryManager`, `UcReposicaoEstoque`)
- **Methods**: PascalCase with descriptive verbs (`AtualizarTodasConsultas`, `ObterTabela`)
- **Variables**: camelCase (`dadosCarregados`, `nomeTabela`)
- **Constants**: UPPER_CASE (`CACHE_TIMEOUT_MINUTES`)
- **Controls**: Type prefix + description (`dgvProdutos`, `btnAtualizar`)

### Documentation Requirements
All public methods must have XML documentation:
```vb
''' <summary>
''' Brief description of what the method does
''' </summary>
''' <param name="paramName">Parameter description</param>
''' <returns>Return value description</returns>
''' <exception cref="ExceptionType">When this exception occurs</exception>
Public Function MethodName(paramName As String) As Boolean
```

## Common Development Workflows

### Adding New Business Logic
1. Extend PowerQueryManagerExtensions.vb for data access
2. Add configuration constants to ConfiguracaoApp.vb
3. Implement UI in appropriate UserControl or Form
4. Add comprehensive error handling and logging
5. Create tests using TestFramework

### Modifying Excel Integration
1. Work with sheet-specific classes (Sheet*.vb files)
2. Use established Power Query patterns for data refresh
3. Maintain Excel state management best practices
4. Test with actual Excel workbook

### Performance Optimization
1. Choose appropriate manager: standard vs optimized
2. Implement caching for expensive operations
3. Use async patterns for long-running operations
4. Monitor via logging and performance tests

## Dependencies & Requirements

### Runtime Requirements
- Windows 10 or superior
- Microsoft Excel 2016+ 
- .NET Framework 4.8
- VSTO Runtime 4.0

### Development Requirements
- Visual Studio 2019+ with VSTO workload
- Microsoft Office Developer Tools
- Excel installed on development machine

### Key References
- Microsoft.Office.Interop.Excel
- Microsoft.Office.Tools.Excel  
- System.Windows.Forms
- MSTest.TestFramework (for unit tests)

## File Structure Notes

The project follows a **domain-driven folder structure**:
- `/Config`: Configuration management classes
- `/Core`: Business logic and data management
- `/Controls`: Custom user controls
- `/Forms`: Windows Forms
- `/Helpers`: Utility classes and extensions
- `/Services`: Testing and deployment services
- `/Tests`: Custom testing framework

Excel sheet files follow the pattern: `Sheet[Domain].vb` with corresponding Designer files.