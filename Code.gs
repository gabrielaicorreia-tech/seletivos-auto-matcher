/**
 * SELETIVOS AUTO MATCHER & CLASSIFICATION SUITE
 * Desenvolvido por: Gabriela Correia
 * Descrição: Sistema para cruzamento de dados e classificação regionalizada com cotas.
 */

// ==========================================
// ⚙️ CONFIGURAÇÕES DA COORDENADORIA (CRS)
// ==========================================

var CONFIG_GERAL = {
  nomeMenu: 'Automação CRS'
};

var CONFIG_CRUZAMENTO = {
  nomeAbaBase: 'base',
  colunasParaPuxar: ["DATA FIM DO CONTRATO", "CH", "CB", "CE", "NOTAREDACAO", "NOTAFINAL", "CLASSGERAL", "CLASSAMPLA", "CLASSPCD", "CLASSRACA", "TIPO", "SELETIVO"]
};

var CONFIG_SELETIVO = {
  abaCandidatos: 'Candidatos', 
  abaQuadroVagas: 'Quadro de Vagas', 
  colunasAgrupamento: ['DRE', 'MUNICÍPIO', 'CARGO'], 
  colunaVagasNoQuadro: 'VAGAS', 
  colunaNota: 'NOTA FINAL',
  colunaIdade: 'IDADE',
  colunaCE: 'CE',
  colunaCota: 'COTA',
  percentualPCD: 0.10, 
  percentualNegros: 0.20 
};

// ==========================================
// MENU DE INTERFACE (UI)
// ==========================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(CONFIG_GERAL.nomeMenu)
      .addItem('1. Cruzar Dados da Base', 'inserirOrdemDiretoNaPagina1')
      .addSeparator()
      .addItem('2. Gerar Classificação e Cotas', 'processarClassificacaoCompleta')
      .addToUi();
}

// ==========================================
// MÓDULO 1: CRUZAMENTO DE DADOS (MATCH)
// ==========================================

function inserirOrdemDiretoNaPagina1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaBase = ss.getSheetByName(CONFIG_CRUZAMENTO.nomeAbaBase);
  var abaAtual = ss.getActiveSheet(); 

  if (!abaBase) {
    SpreadsheetApp.getUi().alert("Erro: Não encontrei a aba '" + CONFIG_CRUZAMENTO.nomeAbaBase + "'.");
    return;
  }

  var dadosBase = abaBase.getDataRange().getValues();
  var dadosAtual = abaAtual.getDataRange().getValues();

  if (dadosBase.length <= 1 || dadosAtual.length <= 1) return;

  var headersBase = dadosBase[0];
  var headersAtual = dadosAtual[0];

  function limparEPadronizarCPF(texto) {
    if (!texto) return "";
    var apenasNumeros = texto.toString().replace(/[^0-9]/g, "").trim();
    if (apenasNumeros === "") return "";
    while (apenasNumeros.length < 11) apenasNumeros = "0" + apenasNumeros;
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
    SpreadsheetApp.getUi().alert("Erro: Coluna 'CPF' não localizada.");
    return;
  }

  var indicesBasePuxar = [];
  for (var i = 0; i < CONFIG_CRUZAMENTO.colunasParaPuxar.length; i++) {
    indicesBasePuxar.push(encontrarColuna(headersBase, [CONFIG_CRUZAMENTO.colunasParaPuxar[i]]));
  }

  var dicionarioBase = {};
  for (var r = 1; r < dadosBase.length; r++) {
    var row = dadosBase[r];
    var cpf = limparEPadronizarCPF(row[baseIdxCPF]);
    if (cpf === "") continue;

    var nome = baseIdxNome !== -1 ? limparTextoSecundario(row[baseIdxNome]) : "";
    var valoresExtraidos = [];
    for (var c = 0; c < indicesBasePuxar.length; c++) {
      valoresExtraidos.push(indicesBasePuxar[c] !== -1 ? row[indicesBasePuxar[c]] : "");
    }

    dicionarioBase[cpf] = valoresExtraidos;
    if (nome !== "") dicionarioBase[cpf + "_" + nome] = valoresExtraidos;
  }

  var novasLinhasPlanilha = [];
  var novoCabecalhoCompleto = headersAtual.slice();
  for (var i = 0; i < CONFIG_CRUZAMENTO.colunasParaPuxar.length; i++) {
    novoCabecalhoCompleto.push(CONFIG_CRUZAMENTO.colunasParaPuxar[i]);
  }
  novasLinhasPlanilha.push(novoCabecalhoCompleto);

  var contSucesso = 0;
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
      for (var i = 0; i < CONFIG_CRUZAMENTO.colunasParaPuxar.length; i++) {
        novaLinhaModificada.push("Não encontrado");
      }
    }
    novasLinhasPlanilha.push(novaLinhaModificada);
  }

  abaAtual.getRange(1, 1, novasLinhasPlanilha.length, novasLinhasPlanilha[0].length).setValues(novasLinhasPlanilha);
  abaAtual.getRange(1, headersAtual.length + 1, 1, CONFIG_CRUZAMENTO.colunasParaPuxar.length).setBackground("#d9ead3");
  SpreadsheetApp.getUi().alert("Processamento concluído! \n✅ Localizados: " + contSucesso);
}

// ==========================================
// MÓDULO 2: CLASSIFICAÇÃO, DESEMPATE E COTAS
// ==========================================

function processarClassificacaoCompleta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCand = ss.getSheetByName(CONFIG_SELETIVO.abaCandidatos);
  var abaVagas = ss.getSheetByName(CONFIG_SELETIVO.abaQuadroVagas);
  
  if (!abaCand || !abaVagas) {
    SpreadsheetApp.getUi().alert("Erro: Verifique se as abas 'Candidatos' e 'Quadro de Vagas' existem.");
    return;
  }
  
  var dadosVagas = abaVagas.getDataRange().getValues();
  var cabVagas = dadosVagas[0];
  var idxVagasDRE = cabVagas.indexOf(CONFIG_SELETIVO.colunasAgrupamento[0]);
  var idxVagasMun = cabVagas.indexOf(CONFIG_SELETIVO.colunasAgrupamento[1]);
  var idxVagasCargo = cabVagas.indexOf(CONFIG_SELETIVO.colunasAgrupamento[2]);
  var idxQtdVagas = cabVagas.indexOf(CONFIG_SELETIVO.colunaVagasNoQuadro);
  
  var dicionarioVagas = {};
  
  function limparChave(texto) {
    if(!texto) return "";
    return texto.toString().normalize('NFD').replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
  }
  
  for (var v = 1; v < dadosVagas.length; v++) {
    var linhaV = dadosVagas[v];
    var chaveLocal = limparChave(linhaV[idxVagasDRE]) + "_" + limparChave(linhaV[idxVagasMun]) + "_" + limparChave(linhaV[idxVagasCargo]);
    dicionarioVagas[chaveLocal] = parseInt(linhaV[idxQtdVagas]) || 0;
  }
  
  var dadosCand = abaCand.getDataRange().getValues();
  var cabCand = dadosCand[0];
  var candidatos = dadosCand.slice(1);
  
  var idxCandDRE = cabCand.indexOf(CONFIG_SELETIVO.colunasAgrupamento[0]);
  var idxCandMun = cabCand.indexOf(CONFIG_SELETIVO.colunasAgrupamento[1]);
  var idxCandCargo = cabCand.indexOf(CONFIG_SELETIVO.colunasAgrupamento[2]);
  
  var idxNota = cabCand.indexOf(CONFIG_SELETIVO.colunaNota);
  var idxIdade = cabCand.indexOf(CONFIG_SELETIVO.colunaIdade);
  var idxCE = cabCand.indexOf(CONFIG_SELETIVO.colunaCE);
  var idxCota = cabCand.indexOf(CONFIG_SELETIVO.colunaCota);
  
  var grupos = {}; 
  
  for (var c = 0; c < candidatos.length; c++) {
    var linhaC = candidatos[c];
    var chaveCand = limparChave(linhaC[idxCandDRE]) + "_" + limparChave(linhaC[idxCandMun]) + "_" + limparChave(linhaC[idxCandCargo]);
                    
    if (!grupos[chaveCand]) grupos[chaveCand] = [];
    grupos[chaveCand].push(linhaC);
  }
  
  cabCand.push("CLASS GERAL", "CLASS AC", "CLASS NEGROS", "CLASS PCD", "BASE DE CÁLCULO");
  var candidatosClassificados = [];
  
  for (var chaveGrupo in grupos) {
    var lote = grupos[chaveGrupo];
    
    // ORDENAÇÃO COM DESEMPATE EM CASCATA
    lote.sort(function(a, b) {
      var notaA = parseFloat(a[idxNota]) || 0;
      var notaB = parseFloat(b[idxNota]) || 0;
      if (notaB !== notaA) return notaB - notaA; 
      
      var idadeA = parseInt(a[idxIdade]) || 0;
      var idadeB = parseInt(b[idxIdade]) || 0;
      if (idadeB !== idadeA) return idadeB - idadeA;
      
      var ceA = idxCE > -1 ? (parseFloat(a[idxCE]) || 0) : 0;
      var ceB = idxCE > -1 ? (parseFloat(b[idxCE]) || 0) : 0;
      return ceB - ceA;
    });
    
    var vagasOfertadas = dicionarioVagas[chaveGrupo] || 0;
    var inscritosNoLote = lote.length;
    var baseCalculo = vagasOfertadas > 0 ? vagasOfertadas : inscritosNoLote;
    
    var limitePCD = Math.ceil(baseCalculo * CONFIG_SELETIVO.percentualPCD);
    var limiteNegros = Math.ceil(baseCalculo * CONFIG_SELETIVO.percentualNegros);
    
    var posGeral = 1, posAC = 1, posPCD = 1, posNegros = 1;
    
    for (var i = 0; i < lote.length; i++) {
      var cotaAtual = lote[i][idxCota].toString().toUpperCase().trim();
      var linhaGeral = posGeral++;
      var linhaAC = "-", linhaPCD = "-", linhaNegros = "-";
      
      if (cotaAtual === "PCD" || cotaAtual === "PESSOA COM DEFICIÊNCIA") {
        linhaPCD = posPCD++;
      } else if (cotaAtual === "NEGROS" || cotaAtual === "COTA RACIAL") {
        linhaNegros = posNegros++;
      } else {
        linhaAC = posAC++;
      }
      
      var infoBase = vagasOfertadas > 0 ? (vagasOfertadas + " vagas") : (inscritosNoLote + " inscritos (CR)");
      lote[i].push(linhaGeral, linhaAC, linhaNegros, linhaPCD, infoBase);
      candidatosClassificados.push(lote[i]); 
    }
  }
  
  var dadosFinais = [cabCand].concat(candidatosClassificados);
  abaCand.clearContents();
  abaCand.getRange(1, 1, dadosFinais.length, dadosFinais[0].length).setValues(dadosFinais);
  abaCand.getRange(1, cabCand.length - 4, 1, 5).setBackground("#e2efda");
  
  SpreadsheetApp.getUi().alert("Classificação concluída!\nAgrupamento, desempate por idade e cotas aplicados com sucesso.");
}
