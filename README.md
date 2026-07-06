# 📊 Seletivos Auto Matcher (Google Sheets & Excel)

Um sistema de alta performance desenvolvido para automatizar o processamento e a classificação de candidatos em processos seletivos de grande escala. 

Este repositório contém o motor de classificação, garantindo flexibilidade para operar tanto em nuvem quanto em desktop:
- **Versão Nuvem:** Google Apps Script (JavaScript) para Google Sheets.
- **Versão Desktop:** VBA para Microsoft Excel.

Ele substitui processos manuais e planilhas lentas por processamento via código, atuando como um verdadeiro sistema de gestão de regras de negócio para editais complexos.

## 🎯 Capacidades do Sistema

O script foi estruturado para atender à demanda de coordenadorias de recrutamento e seleção, aplicando automaticamente as seguintes diretrizes legais:

- **Agrupamento Regionalizado:** Separação e ranqueamento de candidatos em lotes automáticos baseados na chave de lotação (DRE > Município > Cargo).
- **Leitura Dinâmica do Quadro de Vagas:** O motor cruza a lista de inscritos com a aba de vagas autorizadas, definindo automaticamente se a base de cálculo para as cotas será o número de vagas ofertadas ou o total de classificados (Cadastro Reserva).
- **Desempate em Cascata:** Algoritmo de ordenação rigoroso que processa múltiplos critérios de desempate sequencialmente (ex: *Nota Final*, seguida de *Maior Idade*, seguida de *Nota de Conhecimentos Específicos*).
- **Gestão de Cotas Inteligente:** Distribuição simultânea de classificações para Ampla Concorrência (AC), Pessoas com Deficiência (PCD) e Cotas Raciais, identificando excedentes de cota automaticamente.

## 🛠️ Vantagens Técnicas

- **Alta Performance (Dicionários na Memória):** Utiliza objetos de Dicionário em memória para processar agrupamentos e ordenações complexas em milhares de linhas em segundos, evitando o colapso típico de planilhas.
- **Estruturação Pronta para Publicação:** O sistema não apenas ranqueia, mas devolve a base de dados organizada e agrupada por polo, agilizando a exportação para o Diário Oficial.
- **Tratamento de Dados:** Padronização automática de chaves de busca e higienização de textos.

## ⚙️ Estrutura e Modularidade

O código isola a lógica matemática das regras do edital, permitindo adaptações rápidas alterando apenas o bloco de configuração:

```javascript
// Exemplo de configuração do Motor de Classificação:
var CONFIG_SELETIVO = {
  colunasAgrupamento: ['DRE', 'MUNICÍPIO', 'CARGO'], 
  colunaVagasNoQuadro: 'VAGAS', 
  
  colunaNota: 'NOTA FINAL',
  colunaIdade: 'IDADE', // 1º Critério de desempate
  colunaCE: 'CE',       // 2º Critério de desempate
  colunaCota: 'COTA',
  
  percentualPCD: 0.10, 
  percentualNegros: 0.20 
};
