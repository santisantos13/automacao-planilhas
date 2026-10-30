# Automação de Relatório Financeiro

Script em Python que lê uma planilha Excel de controle diário de ganhos e gastos como motorista de aplicativo (Uber/99Pop), limpa os dados e gera um relatório mensal consolidado.

## O problema

A planilha `Controle_Uber.xlsx` é preenchida manualmente todo dia, com uma aba por mês. Analisá-la exigia abrir cada aba e somar os valores manualmente. Além disso, a planilha tinha alguns problemas reais de dados:

- Cabeçalho fora da primeira linha
- Linhas de resumo (totais) misturadas junto com os dados diários
- Datas inconsistentes em algumas abas
- Dias sem quilometragem registrada, gerando erro de divisão por zero (`#DIV/0!`) no cálculo de R$/km

> Nota: os dados de Controle_Uber.xlsx são fictícios, criados para demonstração.

Este script automatiza a leitura, corrige esses problemas e gera um resumo confiável.

## O que o script faz

1. **Leitura** — lê automaticamente todas as abas (meses) do arquivo Excel
2. **Limpeza** — corrige o cabeçalho, remove linhas de resumo, descarta datas inválidas e dias sem dado preenchido
3. **Consolidação** — junta todos os meses em uma única tabela
4. **Relatório** — calcula, por mês: total de ganhos, total de gastos, líquido e R$/km médio (ponderado pelo total de km rodado, não uma média simples)
5. **Exportação** — salva o relatório final em um novo arquivo Excel

## Como rodar

1. Clone este repositório
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Execute o script:
   ```bash
   python main.py
   ```
4. O relatório será gerado em `Relatorio_Mensal.xlsx`

## Tecnologias usadas

- Python
- pandas
- openpyxl

## Possíveis melhorias futuras

- Separar totais por app (Uber vs 99Pop)
- Gerar gráficos de evolução mensal
- Agendar execução automática
