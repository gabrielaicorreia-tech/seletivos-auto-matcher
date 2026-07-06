# 📊 Seletivos Auto Matcher

Um sistema desenvolvido em Google Apps Script (GAS) para automatizar o processamento e a classificação de candidatos em processos seletivos de grande escala diretamente no Google Sheets.

Este projeto substitui processos manuais e planilhas lentas por um motor de classificação via código, desenhado para lidar com as regras complexas de editais, incluindo cálculos de notas, políticas de cotas e distribuição regionalizada.

## 🎯 Capacidades do Sistema

O script foi estruturado para atender à demanda de coordenadorias de recrutamento e seleção, aplicando automaticamente as seguintes regras de negócio:

- **Cálculo de Barema:** Processamento automatizado de pontuações, somando notas de provas, avaliação de titulações e tempo de experiência profissional.
- **Gestão de Cotas:** Classificação segmentada e cruzamento de listas para Ampla Concorrência (AC), Pessoas com Deficiência (PCD) e Cotas Raciais.
- **Critérios de Desempate:** Algoritmos para desempate baseados nas regras do edital (ex: prioridade para candidatos com maior idade).
- **Regionalização:** Estruturação das listas finais agrupadas e filtradas por Cargo concorrido, Município de atuação e DRE (Diretoria Regional de Educação).

## 🛠️ Vantagens Técnicas

- **Alta Performance com Dicionários:** Utiliza objetos em JavaScript na memória para mapear candidatos por chaves compostas (CPF tratado + Nome), processando milhares de linhas em segundos. Evita o colapso e a lentidão típicos de fórmulas como `PROCV/VLOOKUP`.
- **Tratamento e Higienização de Dados:** Padronização automática de dados de entrada (inserção de zeros à esquerda em CPFs desconfigurados e remoção de acentos/caracteres especiais).
- **Interface UI Customizada:** Injeta um menu próprio diretamente na interface do Google Sheets, permitindo que a equipe operacional execute rotinas complexas de classificação com apenas um clique, sem necessidade de interagir com o código.

## ⚙️ Estrutura e Modularidade

O código separa a lógica de processamento das configurações do edital, permitindo adaptações rápidas para novos processos seletivos.

```javascript
// Exemplo de configuração adaptável para diferentes editais:
var CONFIG = {
  nomeMenu: 'Automação Recrutamento', 
  acoes: [
    { nome: 'Processar Barema e Classificação', funcao: 'classificarCandidatos' },
    { nome: 'Gerar Lista de Convocação (Volante)', funcao: 'inserirOrdemDiretoNaPagina1' }
  ],
  colunasChave: ["CPF", "NOME"],
  colunasExtracao: ["TITULACAO", "EXPERIENCIA", "NOTAFINAL", "COTA", "IDADE", "CARGO", "MUNICIPIO", "DRE"]
};
