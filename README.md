# Sistema de Auditoria de Fretes (LPU vs CT-e)

Ferramenta desenvolvida em Python para automatizar a conferência de faturas de transportadoras, identificando divergências de valores, pesos e rotas.

## 🚀 Funcionalidades Atuais
- Leitura de tabelas de frete (LPU) e relatórios de transportadoras (Excel/CSV).
- Identificação automática de rotas (Capital vs Interior).
- Cálculo reverso para descobrir se o peso cobrado foi o "peso cubado" ou "peso real".
- Geração de relatório final com destaque para divergências financeiras.
- Interface gráfica (GUI) construída com Tkinter.

## 🛠️ Tecnologias
- Python 3.12
- Pandas & Numpy (Análise de Dados)
- Tkinter (Interface Gráfica)
- Threading (Processamento assíncrono)

## 🚧 Próximos Passos (Roadmap)
Este projeto está em evolução constante. As próximas melhorias planejadas são:
- [ ] Refatoração: Separar a lógica de negócio da interface gráfica (MVC).
- [ ] Adicionar suporte a outros formatos de tabela.
- [ ] Criar testes unitários para a classe de cálculo.

## 📦 Como rodar
1. Instale as dependências: `pip install -r requirements.txt`
2. Execute o arquivo: `python main.py`