/**
 * AUTO MATCH DATA - GOOGLE SHEETS
 * Desenvolvido por: Gabriela Correia
 * Descrição: Script para cruzamento de dados automático entre abas usando CPF e Nome.
 */

// ==========================================
// ⚙️ CONFIGURAÇÕES GERAIS (Altere aqui)
// ==========================================
var CONFIG = {
  nomeMenu: 'Automação RH', // Nome do menu que aparecerá no Sheets
  nomeBotao: 'Cruzar Dados Automático', // Nome da ação
  nomeAbaBase: 'base', // Aba onde estão os dados completos
  colunasParaPuxar: [
    // Liste aqui o nome exato das colunas que você quer trazer da base
    "DATA FIM DO CONTRATO", "CH", "CB", "CE", "NOTAREDACAO", 
    "NOTAFINAL", "CLASSGERAL", "CLASSAMPLA", "CLASSPCD", 
    "CLASSRACA", "TIPO", "SELETIVO"
  ]
};
// ==========================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(CONFIG.nomeMenu)
      .addItem(CONFIG.nomeBotao, 'inserirOrdemDiretoNaPagina1')
      .addToUi();
}

function inserirOrdemDiretoNaPagina1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaBase = ss.getSheetByName(CONFIG.nomeAbaBase);
  var abaAtual = ss.getActiveSheet(); // Pega a aba que o usuário estiver visualizando

  if (!abaBase) {
    SpreadsheetApp.getUi().alert("Erro: Não encontrei a aba chamada exatamente '" + CONFIG.nomeAbaBase + "'.");
    return;
  }

  var dadosBase = abaBase.getDataRange().getValues();
  var dadosAtual = abaAtual.getDataRange().getValues();

  if (dadosBase.length <= 1 || dadosAtual.length <= 1) {
    SpreadsheetApp.getUi().alert("Erro: Uma das abas está vazia ou contém apenas os cabeçalhos.");
    return;
  }

  var headersBase = dadosBase[0];
  var headersAtual = dadosAtual[0];

  function limparEPadronizarCPF(texto) {
    if (!texto) return "";
    var apenasNumeros = texto.toString().replace(/[^0-9]/g, "").trim();
    if (apenasNumeros === "") return "";
    while (apenasNumeros.length < 11) {
      apenasNumeros = "0" + apenasNumeros;
    }
    return apenasNumeros;
  }

  function limparTextoSecundario(texto) {
    if (!texto) return "";
    return texto.toString().normalize('NFD').replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9]/g, "").trim();
  }

  function encontrarColuna(headers, nomesPossiveis) {
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i].toString().toUpperCase().replace(/[^A-Z0-9]/g, "");
      for (var j = 0; j < nomesPossiveis.length; j++) {
        var p = nomesPossiveis[j].toUpperCase().replace(/[^A-Z0-9]/g, "");
        if (h === p || h.indexOf(p) !== -1) return i;
      }
    }
    return -1;
  }

  var baseIdxCPF = encontrarColuna(headersBase, ["CPF"]);
  var atualIdxCPF = encontrarColuna(headersAtual, ["CPF"]);
  var baseIdxNome = encontrarColuna(headersBase, ["NOME"]);
  var atualIdxNome = encontrarColuna(headersAtual, ["NOME"]);

  if (baseIdxCPF === -1 || atualIdxCPF === -1) {
    SpreadsheetApp.getUi().alert("Erro Crítico: Não localizei a coluna 'CPF' em uma das abas. Verifique a linha de cabeçalho.");
    return;
  }

  var indicesBasePuxar = [];
  var colunasNaoEncontradas = [];

  for (var i = 0; i < CONFIG.colunasParaPuxar.length; i++) {
    var idx = encontrarColuna(headersBase, [CONFIG.colunasParaPuxar[i]]);
    indicesBasePuxar.push(idx);
    if (idx === -1) {
      colunasNaoEncontradas.push(CONFIG.colunasParaPuxar[i]);
    }
  }

  var dicionarioBase = {};
  for (var r = 1; r < dadosBase.length; r++) {
    var row = dadosBase[r];
    var cpf = limparEPadronizarCPF(row[baseIdxCPF]);
    if (cpf === "") continue;

    var nome = baseIdxNome !== -1 ? limparTextoSecundario(row[baseIdxNome]) : "";
    
    var valoresExtraidos = [];
    for (var c = 0; c < indicesBasePuxar.length; c++) {
      var idx = indicesBasePuxar[c];
      valoresExtraidos.push(idx !== -1 ? row[idx] : "");
    }

    dicionarioBase[cpf] = valoresExtraidos;
    if (nome !== "") {
      dicionarioBase[cpf + "_" + nome] = valoresExtraidos;
    }
  }

  var novasLinhasPlanilha = [];
  var novoCabecalhoCompleto = headersAtual.slice();
  for (var i = 0; i < CONFIG.colunasParaPuxar.length; i++) {
    novoCabecalhoCompleto.push(CONFIG.colunasParaPuxar[i]);
  }
  novasLinhasPlanilha.push(novoCabecalhoCompleto);

  var contSucesso = 0;
  var contFalha = 0;

  for (var r = 1; r < dadosAtual.length; r++) {
    var row = dadosAtual[r];
    var novaLinhaModificada = row.slice();

    var cpf = limparEPadronizarCPF(row[atualIdxCPF]);
    var nome = atualIdxNome !== -1 ? limparTextoSecundario(row[atualIdxNome]) : "";

    var dadosPuxados = dicionarioBase[cpf + "_" + nome] || dicionarioBase[cpf];

    if (dadosPuxados) {
      novaLinhaModificada = novaLinhaModificada.concat(dadosPuxados);
      contSucesso++;
    } else {
      for (var i = 0; i < CONFIG.colunasParaPuxar.length; i++) {
        novaLinhaModificada.push("Não encontrado");
      }
      contFalha++;
    }
    novasLinhasPlanilha.push(novaLinhaModificada);
  }

  abaAtual.getRange(1, 1, novasLinhasPlanilha.length, novasLinhasPlanilha[0].length).setValues(novasLinhasPlanilha);
  
  var tamanhoColOriginal = headersAtual.length;
  abaAtual.getRange(1, tamanhoColOriginal + 1, 1, CONFIG.colunasParaPuxar.length).setBackground("#d9ead3");

  var mensagemResultado = "Processamento concluído!\n\n" + 
                          "✅ Localizados: " + contSucesso + "\n" +
                          "❌ Sem correspondência: " + contFalha;
  
  if (colunasNaoEncontradas.length > 0) {
    mensagemResultado += "\n\n⚠️ Atenção: As seguintes colunas não foram encontradas na base:\n" + colunasNaoEncontradas.join(", ");
  }

  SpreadsheetApp.getUi().alert(mensagemResultado);
}
