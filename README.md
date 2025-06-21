# EstoqueDifemaqGestao

Sistema de gestão de estoque integrado com Excel para controle eficiente de produtos, movimentações e relatórios.

## 🚀 Funcionalidades

- **Gestão de Produtos**: Cadastro e controle de produtos com imagens
- **Controle de Estoque**: Monitoramento em tempo real dos níveis de estoque
- **Histórico de Movimentações**: Registro completo de compras e vendas
- **Relatórios Dinâmicos**: Dashboards e relatórios personalizáveis
- **Integração Excel**: Power Query para sincronização automática de dados

## 📋 Documentação

- [Coding Standards](Docs/CodeStandards/CODING_STANDARDS.md)
- [API Documentation](Docs/API.md)

## 🛠️ Tecnologias

- **Framework**: .NET Framework 4.7.2+
- **Linguagem**: Visual Basic .NET
- **Interface**: Windows Forms
- **Dados**: Microsoft Excel com Power Query
- **Controle de Versão**: Git

## 📦 Requisitos do Sistema

- Windows 10 ou superior
- Microsoft Excel 2016 ou superior
- .NET Framework 4.7.2 ou superior
- 4GB RAM mínimo (recomendado: 8GB)

## 🚀 Instalação

1. Clone o repositório:
   ```bash
   git clone [url-do-repositorio]
   ```

2. Abra o projeto no Visual Studio 2019 ou superior

3. Configure o caminho das imagens em `ConfiguracaoApp.vb`

4. Compile e execute o projeto

## 📖 Como Usar

1. **Primeira Execução**: Configure os caminhos dos arquivos Excel
2. **Carregamento de Dados**: Use o botão "Atualizar Dados" para sincronizar
3. **Gestão de Produtos**: Navegue pelos produtos usando os filtros
4. **Visualização**: Selecione um produto para ver detalhes e histórico

## 🏗️ Estrutura do Projeto

```
EstoqueDifemaqGestao/
├── Config/                 # Configurações do sistema
├── Controls/               # Controles personalizados
├── Core/                   # Classes principais (PowerQuery, etc.)
├── Docs/                   # Documentação
│   └── CodeStandards/      # Padrões de codificação
├── Forms/                  # Formulários da aplicação
├── Helpers/                # Classes auxiliares
└── Services/               # Serviços e testes
```

## 🤝 Contribuição

1. Siga os [padrões de codificação](Docs/CodeStandards/CODING_STANDARDS.md)
2. Execute os testes antes de commitar
3. Documente novas funcionalidades
4. Use commits descritivos

## 📄 Licença

Este projeto está sob licença [MIT](LICENSE).

## 📞 Suporte

Para dúvidas ou problemas, consulte a documentação ou entre em contato com a equipe de desenvolvimento.

---

*Última atualização: Junho 2025*
