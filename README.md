# Sequência Violada - Automação de Controle de Dias Trabalhados Consecutivos

## 📌 Descrição

Este projeto tem como objetivo automatizar o processo de geração de relatórios de colaboradores que trabalharam em sequência de dias sem descanso, substituindo controles manuais em Excel por um processo automatizado em **VBA** e **SQL**, com integração à **API do TOTVS RM**.

A automação gera, para cada coligada:

- Uma aba **BASE** com os dados extraídos da API.
- Uma **Tabela Dinâmica** com o resumo por colaborador.
- Um **Gráfico** baseado nos dados da sequência.
- Um arquivo `.xlsx` salvo automaticamente em uma pasta nomeada com a coligada e período.

## ⚙️ Tecnologias e Ferramentas

- **VBA (Visual Basic for Applications)**: automação das etapas.
- **SQL (TOTVS / RM Reports)**: estruturação das consultas.
- **Excel**: geração dos relatórios.
- **API TOTVS RM**: fonte dos dados.

## 🧠 Lógica da Automação

1. Um módulo VBA (`Painel`) executa a extração para todas as coligadas configuradas.
2. A Sub `Extrair_API_Nova` realiza:
   - Requisição HTTP para a API.
   - Processamento do JSON retornado.
   - Preenchimento da aba "BASE".
   - Criação da tabela e atualização de Tabela Dinâmica e Gráfico.
3. As sequências de batidas são analisadas por consulta SQL para identificar violações de descanso semanal.
4. O sistema classifica a gravidade da sequência:
   - `OK`: até 6 dias consecutivos
   - `Alerta`: 7 a 14 dias
   - `Crítico`: 15 dias ou mais

## 📁 Estrutura dos Arquivos

SequenciaViolada/
├── VBA/
│ └── Extrair_API_Nova.bas
├── SQL/
│ └── ConsultaSequencia.sql
├── README.md
└── ExemploRelatorio/
└── ColigadaX_Sequencia_2025-05.xlsx

markdown
Copiar
Editar

## ✅ Resultados

- Redução de **3 a 5 horas** de trabalho manual por mês para cerca de **5 minutos**.
- Arquivos gerados automaticamente com dados confiáveis e formatados.

## 🚧 Futuras Melhorias

- Adicionar testes automatizados para as funções VBA.
- Exportar também para PDF.
- Interface de execução via formulário no Excel.
