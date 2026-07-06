# 📊 Seletivos Auto Matcher (Google Sheets & Excel)

Um sistema de alta performance desenvolvido para automatizar o processamento e a classificação de candidatos em processos seletivos de grande escala. 

Este repositório contém duas versões do motor de classificação, garantindo flexibilidade para operar tanto em nuvem quanto em desktop:
- **Versão Nuvem:** Google Apps Script (JavaScript) para Google Sheets.
- **Versão Desktop:** VBA para Microsoft Excel.

Ambos substituem processos manuais e planilhas lentas por processamento via código, desenhados para lidar com regras complexas de editais.

## 🎯 Capacidades do Sistema

O script foi estruturado para atender à demanda de coordenadorias de recrutamento e seleção, aplicando automaticamente as seguintes regras de negócio:

- **Cálculo de Barema:** Processamento automatizado de pontuações, somando notas de provas, avaliação de titulações e tempo de experiência profissional.
- **Gestão de Cotas:** Classificação segmentada e cruzamento de listas para Ampla Concorrência (AC), Pessoas com Deficiência (PCD) e Cotas Raciais.
- **Critérios de Desempate:** Algoritmos para desempate baseados nas regras do edital (ex: prioridade para candidatos com maior idade).
- **Regionalização:** Estruturação das listas finais agrupadas e filtradas por Cargo concorrido, Município de atuação e DRE (Diretoria Regional de Educação).

## 🛠️ Vantagens Técnicas

- **Alta Performance (Dicionários na Memória):** Ambas as versões (GAS e VBA) utilizam objetos de Dicionário na memória para mapear candidatos por chaves compostas (CPF tratado + Nome). Isso permite processar milhares de linhas em segundos, evitando o travamento típico de fórmulas nativas (`PROCV/VLOOKUP`).
- **Tratamento de Dados:** Padronização automática (inserção de zeros à esquerda em CPFs desconfigurados e remoção de caracteres especiais).
- **Interface Amigável:** Injeção de menus customizados nas planilhas, permitindo que a equipe operacional execute as rotinas com um clique.

## 📁 Estrutura do Repositório

- `Code.gs` -> Código-fonte em Google Apps Script para implementação no Google Sheets.
- `CruzarDados_Excel.vba` -> Código-fonte em Visual Basic for Applications para execução offline no Microsoft Excel.

---
*Desenvolvido para garantir eficiência, transparência e segurança de dados na gestão de processos seletivos.*
