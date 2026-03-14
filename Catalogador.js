// ============================================================================
// BE - CATALOGADOR (Standalone Apps Script)
// ============================================================================
// Percorre TODOS os meses da estrutura Drive, encontra PDFs em "PARA CATALOGAR"
// e cataloga-os na spreadsheet do mês correspondente.
//
// Estrutura Drive esperada:
//   PASTA_GERAL_FATURAS/
//     YYYY/
//       Faturas_{CODIGO}_MM/YYYY/
//         #0 - Faturas_{CODIGO}_MM/YYYY   ← Spreadsheet do mês
//         #1 - Faturas e NCs normais/
//           PARA CATALOGAR/               ← PDFs por catalogar
//           (ficheiros já catalogados)
//         #2 - Faturas e NCs com reembolso/
//           PARA CATALOGAR/
//         #3 - Outros documentos/
//           PARA CATALOGAR/
//
// Spreadsheet tabs:
//   "Faturas e NCs normais"       ← para #1
//   "Faturas e NCs com reembolso" ← para #2
//   "Outros documentos"           ← para #3
//
// Corre a cada 4 horas via trigger.
// Suporta retoma automática (ScriptProperties) para não exceder 6 min.
// ============================================================================

// === CONFIGURAÇÃO POR EMPRESA (copiar para outras empresas) ===

// Pasta raiz das faturas de compras
const PASTA_GERAL_FATURAS = "17Onz--A6H-AdeMon0AvK3wRCvkNQhsq2";

// Código da empresa (usado nos nomes das pastas: Faturas_DL_MM/YYYY)
const CODIGO_EMPRESA = "DL";

// Spreadsheet de fornecedores (para lookup NIF → nome)
const FORNECEDORES_SHEET_ID = "1iUQQIGUJaTSDZn0MF9H3Kc7GREuWhGTycloovo9P3xc";

// NIFs da própria empresa (excluir da extracção de fornecedor)
const NIF_PROPRIO_1 = "516008803"; // Darkland
const NIF_PROPRIO_2 = "514654473"; // Darkpurple

// Email de notificação
const EMAIL_NOTIFICACAO = "financeiro@arrowplus.pt";

// Mapeamento sub-pasta → tab na spreadsheet
const PASTA_TAB_MAP = {
  "#1 - Faturas e NCs normais":       "Faturas e NCs normais",
  "#2 - Faturas e NCs com reembolso": "Faturas e NCs com reembolso",
  "#3 - Outros documentos":           "Outros documentos"
};

// Mapeamento NIF → Motivo/Rubrica (conhecidos)
const MOTIVO_POR_NIF = {
  "504615947": "Comunicações",                // MEO
  "503062081": "Comunicações",                // NOWO
  "503022136": "Contabilidade",               // AS Conta
  "500940231": "Seguros",                     // Generali/Tranquilidade
  "503504564": "Utilidades (água/luz/outros)", // EDP
  "500697370": "Utilidades (água/luz/outros)"  // Petrogal
};


// ============================================================================
// FUNÇÃO PRINCIPAL (activada por trigger, 4 em 4 horas)
// ============================================================================

function catalogarTudo() {
  var TIME_BUDGET_MS = 5 * 60 * 1000 - 15000; // 5 min - margem
  var start = Date.now();

  var props = PropertiesService.getScriptProperties();
  var YEAR_KEY = "CAT_YEAR_IDX";
  var MONTH_KEY = "CAT_MONTH_IDX";
  var SUB_KEY = "CAT_SUB_IDX";
  var FILE_KEY = "CAT_FILE_POS";

  // Totais para resumo
  var totalCatalogados = 0;
  var totalErros = 0;
  var resumo = "";

  // Descobrir anos disponíveis
  var pastaRaiz = DriveApp.getFolderById(PASTA_GERAL_FATURAS);
  var anos = _listarAnos(pastaRaiz);

  var yearIdx = Number(props.getProperty(YEAR_KEY) || 0);
  if (yearIdx >= anos.length) yearIdx = 0;

  for (; yearIdx < anos.length; yearIdx++) {
    var anoFolder = anos[yearIdx];
    var year = anoFolder.getName();
    Logger.log("═══ ANO: " + year + " ═══");

    // Listar pastas de mês
    var meses = _listarMeses(anoFolder, year);
    var monthIdx = Number(props.getProperty(MONTH_KEY) || 0);
    if (monthIdx >= meses.length) monthIdx = 0;

    for (; monthIdx < meses.length; monthIdx++) {
      var mesInfo = meses[monthIdx];
      Logger.log("── Mês: " + mesInfo.label + " ──");

      // Encontrar spreadsheet do mês
      var ss = _encontrarSpreadsheetDoMes(mesInfo.folder, mesInfo.label);
      if (!ss) {
        Logger.log("⚠️ Spreadsheet não encontrada para " + mesInfo.label);
        continue;
      }

      // Percorrer sub-pastas (#1, #2, #3)
      var subPastas = Object.keys(PASTA_TAB_MAP);
      var subIdx = Number(props.getProperty(SUB_KEY) || 0);
      if (subIdx >= subPastas.length) subIdx = 0;

      for (; subIdx < subPastas.length; subIdx++) {
        var nomeSub = subPastas[subIdx];
        var nomeTab = PASTA_TAB_MAP[nomeSub];

        var subFolder = _getSubFolder(mesInfo.folder, nomeSub);
        if (!subFolder) continue;

        var paraCatalogar = _getSubFolder(subFolder, "PARA CATALOGAR");
        if (!paraCatalogar) continue;

        // Listar PDFs
        var pdfs = _listarPDFs(paraCatalogar);
        if (pdfs.length === 0) continue;

        Logger.log("  " + nomeSub + " → " + pdfs.length + " PDFs em PARA CATALOGAR");

        // Abrir tab da spreadsheet
        var sheet = ss.getSheetByName(nomeTab);
        if (!sheet) {
          Logger.log("  ⚠️ Tab '" + nomeTab + "' não encontrada na spreadsheet");
          continue;
        }

        var filePos = Number(props.getProperty(FILE_KEY) || 0);
        if (filePos >= pdfs.length) filePos = 0;

        for (var f = filePos; f < pdfs.length; f++) {
          // Verificar time budget
          if ((Date.now() - start) > TIME_BUDGET_MS) {
            props.setProperty(YEAR_KEY, String(yearIdx));
            props.setProperty(MONTH_KEY, String(monthIdx));
            props.setProperty(SUB_KEY, String(subIdx));
            props.setProperty(FILE_KEY, String(f));
            Logger.log("⏱️ Time budget — retoma guardada.");
            _cleanupTempFiles();
            _enviarResumo(totalCatalogados, totalErros, resumo, true);
            return;
          }

          var pdf = pdfs[f];
          var resultado = _catalogarUmPDF(pdf, sheet, subFolder, paraCatalogar, mesInfo.month, mesInfo.year);

          if (resultado.sucesso) {
            totalCatalogados++;
            resumo += "\n ✅ " + resultado.mensagem;
          } else {
            totalErros++;
            resumo += "\n ❌ " + resultado.mensagem;
          }
        }

        // Terminou esta sub-pasta
        props.deleteProperty(FILE_KEY);
      }

      // Terminou sub-pastas #1/#2/#3
      props.deleteProperty(SUB_KEY);

      // ── Catalogar Recibos (#4) e Comprovativos (#5) ──
      var tabsFaturas = [
        ss.getSheetByName("Faturas e NCs normais"),
        ss.getSheetByName("Faturas e NCs com reembolso"),
        ss.getSheetByName("Outros documentos")
      ].filter(function(s) { return !!s; });

      // #4 - Recibos
      var pastaRecibos = _getSubFolder(mesInfo.folder, "#4 - Recibos");
      if (pastaRecibos) {
        var pcRecibos = _getSubFolder(pastaRecibos, "PARA CATALOGAR");
        if (pcRecibos) {
          var pdfsRecibos = _listarPDFs(pcRecibos);
          if (pdfsRecibos.length > 0) {
            Logger.log("  #4 - Recibos → " + pdfsRecibos.length + " PDFs em PARA CATALOGAR");
            for (var r = 0; r < pdfsRecibos.length; r++) {
              if ((Date.now() - start) > TIME_BUDGET_MS) {
                Logger.log("⏱️ Time budget (recibos)");
                _cleanupTempFiles();
                _enviarResumo(totalCatalogados, totalErros, resumo, true);
                return;
              }
              var resRecibo = _catalogarRecibo(pdfsRecibos[r], tabsFaturas, pastaRecibos, pcRecibos);
              if (resRecibo.sucesso) { totalCatalogados++; resumo += "\n ✅ " + resRecibo.mensagem; }
              else { totalErros++; resumo += "\n ❌ " + resRecibo.mensagem; }
            }
          }
        }
      }

      // #5 - Comprovativos de pagamento
      var pastaComprovativos = _getSubFolder(mesInfo.folder, "#5 - Comprovativos de pagamento");
      if (pastaComprovativos) {
        var pcComprovativos = _getSubFolder(pastaComprovativos, "PARA CATALOGAR");
        if (pcComprovativos) {
          var pdfsComprovativos = _listarPDFs(pcComprovativos);
          if (pdfsComprovativos.length > 0) {
            Logger.log("  #5 - Comprovativos → " + pdfsComprovativos.length + " PDFs em PARA CATALOGAR");
            for (var c = 0; c < pdfsComprovativos.length; c++) {
              if ((Date.now() - start) > TIME_BUDGET_MS) {
                Logger.log("⏱️ Time budget (comprovativos)");
                _cleanupTempFiles();
                _enviarResumo(totalCatalogados, totalErros, resumo, true);
                return;
              }
              var resComp = _catalogarComprovativo(pdfsComprovativos[c], tabsFaturas, pastaComprovativos, pcComprovativos);
              if (resComp.sucesso) { totalCatalogados++; resumo += "\n ✅ " + resComp.mensagem; }
              else { totalErros++; resumo += "\n ❌ " + resComp.mensagem; }
            }
          }
        }
      }
    }

    // Terminou este ano
    props.deleteProperty(MONTH_KEY);
  }

  // Terminou tudo — limpar estado
  props.deleteProperty(YEAR_KEY);
  props.deleteProperty(MONTH_KEY);
  props.deleteProperty(SUB_KEY);
  props.deleteProperty(FILE_KEY);

  Logger.log("✅ Catalogação completa: " + totalCatalogados + " catalogados, " + totalErros + " erros.");
  _cleanupTempFiles();
  _enviarResumo(totalCatalogados, totalErros, resumo, false);
}

/** Limpa o estado de retoma (recomeça do zero) */
function catalogadorReset() {
  var props = PropertiesService.getScriptProperties();
  props.getKeys().forEach(function(k) {
    if (k.indexOf("CAT_") === 0) props.deleteProperty(k);
  });
  Logger.log("🔄 Estado de retoma limpo.");
}


// ============================================================================
// LÓGICA DE CATALOGAÇÃO DE UM PDF
// ============================================================================

function _catalogarUmPDF(file, sheet, pastaDestino, paraCatalogar, mesPlanilha, anoPlanilha) {
  var fileName = file.getName();
  Logger.log("    📄 " + fileName);

  // 1. OCR
  var textoPDF = "";
  try {
    textoPDF = convertPDFToText(file.getId(), ['pt', 'en', null]) || "";
  } catch (e) {
    return { sucesso: false, mensagem: fileName + ": Falha OCR (" + String(e).substring(0, 80) + ")" };
  }
  if (!textoPDF.trim()) {
    return { sucesso: false, mensagem: fileName + ": PDF sem texto após OCR" };
  }

  // 2. Consenso de data (6 fontes — igual ao DISTRIBUIDOR)
  var consenso = _consensoData(fileName, textoPDF);
  var data = consenso.data || "ERRO AO ANALISAR";

  // 3. Extracção via IA (tudo numa chamada — valores financeiros, tipo, fornecedor, NIF)
  var valoresIA = _extrairTudoViaIA(textoPDF);

  // 4. Extracções regex como fallback
  var atcud = extractATCUD(textoPDF);
  var nif = extractNIF(textoPDF);
  var fornecedor = extractFornecedor(textoPDF);
  var tipo = extractTipoDocumento(textoPDF);

  // Mesclar: IA tem prioridade quando regex falha
  if (valoresIA) {
    if (!atcud && valoresIA.atcud) atcud = valoresIA.atcud;
    if (!tipo && valoresIA.tipo) tipo = valoresIA.tipo;
    if (!nif && valoresIA.nif) nif = valoresIA.nif;
    if (!fornecedor && valoresIA.fornecedor) fornecedor = valoresIA.fornecedor;
  }

  // 4. Verificar ATCUD duplicado (com lógica digital vs scan/papel)
  if (atcud) {
    var atcudNorm = String(atcud).replace(/\s/g, "");
    var dupInfo = _findATCUDDuplicado(sheet, atcudNorm);
    if (dupInfo) {
      var novoEhDigital = _isPDFDigital(file.getId());
      var existenteEhDigital = _isPDFDigital(dupInfo.fileId);
      Logger.log("    ⚠️ ATCUD duplicado: " + atcud + " | Existente: " + (existenteEhDigital ? "digital" : "scan") + " | Novo: " + (novoEhDigital ? "digital" : "scan"));

      if (existenteEhDigital && !novoEhDigital) {
        // Existente é digital, novo é scan → scan fica como Nº-P.pdf, sem Excel
        var numExistente = dupInfo.numero;
        var nomePapel = numExistente + "-P.pdf";
        file.setName(nomePapel);
        file.moveTo(pastaDestino);
        Logger.log("    📄 Papel arquivado como " + nomePapel + " (digital já existe)");
        return { sucesso: true, mensagem: nomePapel + " (papel, ATCUD duplicado com digital " + numExistente + ")" };

      } else if (!existenteEhDigital && novoEhDigital) {
        // Existente é scan, novo é digital → digital assume o nº, scan renomeia para Nº-P
        var numExistente = dupInfo.numero;

        // Renomear scan existente para Nº-P.pdf
        var ficheiroExistente = DriveApp.getFileById(dupInfo.fileId);
        ficheiroExistente.setName(numExistente + "-P.pdf");
        Logger.log("    📄 Scan existente renomeado para " + numExistente + "-P.pdf");

        // Digital assume o número
        file.setName(numExistente + ".pdf");
        file.moveTo(pastaDestino);

        // Actualizar link no Excel para apontar para o digital
        if (dupInfo.row > 0 && dupInfo.colLink > 0) {
          sheet.getRange(dupInfo.row, dupInfo.colLink).setFormula('=HYPERLINK("' + file.getUrl() + '";"LINK")');
          Logger.log("    🔗 Link no Excel actualizado para o digital");
        }

        return { sucesso: true, mensagem: numExistente + ".pdf (digital substitui scan, papel mantido como " + numExistente + "-P.pdf)" };

      } else {
        // Mesmo tipo (ambos digital ou ambos scan) → duplicado real, ignorar
        Logger.log("    ⚠️ Duplicado real (mesmo tipo): " + atcud);
        return { sucesso: false, mensagem: fileName + ": ATCUD duplicado (" + atcud + ")" };
      }
    }
  }

  // 5. Resolver fornecedor via BD (fluxo completo com fallbacks)
  var nomeFornecedorBD = null;
  Logger.log("    Fornecedor OCR: '" + fornecedor + "' | NIF: '" + nif + "'");

  // 5a. Se temos NIF → procurar na BD
  if (nif && FORNECEDORES_SHEET_ID) {
    nomeFornecedorBD = _findFornecedorByNIF(nif);
    Logger.log("    BD por NIF: '" + nomeFornecedorBD + "'");
  }

  // 5b. Se NIF não encontrado → tentar por email/telefone
  if (!nif && FORNECEDORES_SHEET_ID) {
    var emails = extractEmail(textoPDF);
    var telefones = extractTelefone(textoPDF);
    Logger.log("    NIF não encontrado, a tentar email=" + emails + " tel=" + telefones);
    nif = _findNIFByEmailTelefone(emails, telefones);
    if (nif) {
      nomeFornecedorBD = _findFornecedorByNIF(nif);
      Logger.log("    BD por email/tel: NIF=" + nif + " Nome='" + nomeFornecedorBD + "'");
    }
  }

  if (nomeFornecedorBD) fornecedor = nomeFornecedorBD;

  // 6. Mapear colunas
  var LINHA_CABECALHO = 2;
  var cols = {
    num:       encontraColunaNoCabecalho(sheet, "Nº", LINHA_CABECALHO),
    data:      encontraColunaNoCabecalho(sheet, "Data documento", LINHA_CABECALHO),
    tipo:      encontraColunaNoCabecalho(sheet, "Tipo de documento", LINHA_CABECALHO),
    atcud:     encontraColunaNoCabecalho(sheet, "ATCUD / Nº Documento", LINHA_CABECALHO),
    fornec:    encontraColunaNoCabecalho(sheet, "Fornecedor", LINHA_CABECALHO),
    nif:       encontraColunaNoCabecalho(sheet, "NIF/NIPC fornecedor", LINHA_CABECALHO),
    bt:        encontraColunaNoCabecalho(sheet, "Base tributável", LINHA_CABECALHO),
    iva:       encontraColunaNoCabecalho(sheet, "IVA", LINHA_CABECALHO),
    retencoes: encontraColunaNoCabecalho(sheet, "Retenções", LINHA_CABECALHO),
    outros:    encontraColunaNoCabecalho(sheet, "Outros", LINHA_CABECALHO),
    total:     encontraColunaNoCabecalho(sheet, "Valor total", LINHA_CABECALHO),
    motivo:    encontraColunaNoCabecalho(sheet, "Motivo", LINHA_CABECALHO),
    link:      encontraColunaNoCabecalho(sheet, "Link documento", LINHA_CABECALHO),
    colab:     encontraColunaNoCabecalho(sheet, "Colaborador que insere", LINHA_CABECALHO)
  };

  // 7. Número sequencial
  var ultimaLinha = sheet.getLastRow();
  var novoNum = 1;
  if (ultimaLinha > LINHA_CABECALHO && cols.num > 0) {
    var lastNum = sheet.getRange(ultimaLinha, cols.num).getValue();
    novoNum = (parseInt(lastNum) || 0) + 1;
  }

  // 8. Escrever na spreadsheet
  var novaLinha = ultimaLinha + 1;
  var ultimaColuna = sheet.getLastColumn();

  if (cols.num > 0) sheet.getRange(novaLinha, cols.num).setValue(novoNum);

  if (cols.data > 0) {
    sheet.getRange(novaLinha, cols.data).setValue(data);
    // Colorir: vermelho se inválida, laranja se mês diferente
    var corData = "black";
    if (data === "ERRO AO ANALISAR") {
      corData = "red";
    } else {
      var partes = data.split("/");
      if (partes.length === 3) {
        var mesFatura = partes[1];
        var anoFatura = partes[2].length === 4 ? partes[2] : partes[0];
        if (mesFatura !== mesPlanilha || anoFatura !== anoPlanilha) corData = "orange";
      } else {
        corData = "red";
      }
    }
    sheet.getRange(novaLinha, cols.data).setFontColor(corData);
  }

  if (cols.tipo > 0) sheet.getRange(novaLinha, cols.tipo).setValue(tipo || "");
  if (cols.atcud > 0) sheet.getRange(novaLinha, cols.atcud).setValue(atcud || "");

  if (cols.fornec > 0) {
    sheet.getRange(novaLinha, cols.fornec).setValue(fornecedor || "");
    if (!nomeFornecedorBD) {
      sheet.getRange(novaLinha, cols.fornec).setFontColor("red").setNote("NÃO ESTÁ REGISTADO");
    } else {
      sheet.getRange(novaLinha, cols.fornec).setFontColor("black");
    }
  }

  if (cols.nif > 0) sheet.getRange(novaLinha, cols.nif).setValue(nif || "");
  if (cols.motivo > 0 && nif) sheet.getRange(novaLinha, cols.motivo).setValue(MOTIVO_POR_NIF[nif] || "");
  if (cols.colab > 0) sheet.getRange(novaLinha, cols.colab).setValue("SOFTWARE");

  // Valores financeiros
  var bt = 0, iva = 0, ret = 0, out = 0;
  if (valoresIA) {
    bt = valoresIA.bt || 0;
    iva = valoresIA.iva || 0;
    ret = valoresIA.retencoes || 0;
    out = valoresIA.outros || 0;
  }

  var corValores = (bt > 0) ? "black" : "red";
  if (cols.bt > 0) sheet.getRange(novaLinha, cols.bt).setValue(bt).setFontColor(corValores);
  if (cols.iva > 0) sheet.getRange(novaLinha, cols.iva).setValue(iva).setFontColor(corValores);
  if (cols.retencoes > 0) sheet.getRange(novaLinha, cols.retencoes).setValue(ret).setFontColor(corValores);
  if (cols.outros > 0) sheet.getRange(novaLinha, cols.outros).setValue(out).setFontColor(corValores);

  // Valor total = fórmula
  if (cols.total > 0 && cols.bt > 0) {
    var formula = "=R[0]C[" + (cols.bt - cols.total) + "]+R[0]C[" + (cols.iva - cols.total) + "]-R[0]C[" + (cols.retencoes - cols.total) + "]+R[0]C[" + (cols.outros - cols.total) + "]";
    sheet.getRange(novaLinha, cols.total).setFormulaR1C1(formula).setBackground("#d9d9d9");
  }

  // Link para o ficheiro (antes de mover — URL não muda)
  if (cols.link > 0) {
    sheet.getRange(novaLinha, cols.link).setFormula('=HYPERLINK("' + file.getUrl() + '";"LINK")');
  }

  // Bordas
  if (ultimaColuna > 0) {
    sheet.getRange(novaLinha, 1, 1, ultimaColuna).setBorder(true, true, true, true, true, true);
  }

  // Colorir NCs
  if (tipo === "Nota de crédito" && cols.data > 0 && cols.total > 0) {
    sheet.getRange(novaLinha, cols.data, 1, cols.total - cols.data).setBackground("#FFEBEB");
  }

  // Flush para garantir que o Sheet é gravado antes de mover
  SpreadsheetApp.flush();

  // 9. Renomear e mover (o ID não muda, links mantêm-se válidos)
  var novoNome = novoNum + ".pdf";
  file.setName(novoNome);
  try {
    file.moveTo(pastaDestino);
  } catch (e) {
    Logger.log("    ⚠️ moveTo falhou, a tentar makeCopy: " + String(e).substring(0, 80));
    var copia = file.makeCopy(novoNome, pastaDestino);
    file.setTrashed(true);
    // Actualizar link para apontar para a cópia
    if (cols.link > 0) {
      sheet.getRange(novaLinha, cols.link).setFormula('=HYPERLINK("' + copia.getUrl() + '";"LINK")');
      SpreadsheetApp.flush();
    }
  }

  Logger.log("    ✅ Catalogado como " + novoNome + " | " + (fornecedor || "?") + " | " + data);
  return { sucesso: true, mensagem: novoNome + " (" + (fornecedor || "?") + ") → " + data };
}


// ============================================================================
// EXTRACÇÃO VIA IA — CONSENSO DE VALORES (4 modelos)
// ============================================================================

function _extrairTudoViaIA(textoPDF) {
  var prompt = "Analisa o seguinte texto de uma fatura portuguesa e extrai os valores numéricos de:\n" +
    "- Base tributável (valor sem IVA, também chamado 'incidência', 'base', 'subtotal sem IVA')\n" +
    "- IVA (imposto sobre valor acrescentado)\n" +
    "- Retenções (retenções na fonte de IRS/IRC, se existirem)\n" +
    "- Outros (outros impostos como Imposto de Selo, taxas, se existirem)\n" +
    "- ATCUD (código único do documento, formato tipicamente XXXXXXXX-NNNNN)\n" +
    "- Tipo de documento (um de: Fatura, Fatura simplificada, Fatura-recibo, 2ª via fatura, Nota de crédito, Recibo, Recibo de renda)\n" +
    "- NIF do fornecedor (número de identificação fiscal do emissor)\n" +
    "- Nome do fornecedor (nome/razão social do emissor)\n\n" +
    "Regras:\n" +
    "- Valor total = Base tributável + IVA - Retenções + Outros\n" +
    "- EXCEPÇÃO para faturas de crédito/leasing que discriminem Capital e Juros: a base tributável são APENAS os Juros e encargos (o Capital/amortização NÃO conta). Em comissões, seguros e outras faturas bancárias sem discriminação Capital/Juros, usa o valor total como base tributável normalmente.\n\n" +
    "Responde APENAS com um JSON válido:\n" +
    '{"bt": 0.00, "iva": 0.00, "retencoes": 0.00, "outros": 0.00, "atcud": "", "tipo": "", "nif": "", "fornecedor": ""}\n' +
    "Se não encontrares um valor, usa 0. Campos de texto desconhecidos = \"\".\n" +
    "Usa números com ponto decimal (ex: 123.45).\n\n" +
    "Texto da fatura:\n" + textoPDF.substring(0, 4000);

  var apis = [
    { fn: function() { return chamarMistral(prompt); }, nome: "Mistral" },
    { fn: function() { return chamarGroq(prompt); }, nome: "Groq" },
    { fn: function() { return chamarGemini(prompt, "gemini-2.0-flash"); }, nome: "Gemini 2.0" },
    { fn: function() { return chamarGemini(prompt, "gemini-3.1-flash-lite-preview"); }, nome: "Gemini Lite" }
  ];

  // Recolher respostas de todos os modelos
  var respostas = [];
  for (var i = 0; i < apis.length; i++) {
    try {
      var resposta = apis[i].fn();
      resposta = resposta.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
      var jsonMatch = resposta.match(/\{[\s\S]*\}/);
      if (!jsonMatch) { Logger.log("    💰 " + apis[i].nome + ": sem JSON"); continue; }

      var obj = JSON.parse(jsonMatch[0]);
      var parsed = {
        bt: Math.round((parseFloat(obj.bt) || 0) * 100) / 100,
        iva: Math.round((parseFloat(obj.iva) || 0) * 100) / 100,
        retencoes: Math.round((parseFloat(obj.retencoes) || 0) * 100) / 100,
        outros: Math.round((parseFloat(obj.outros) || 0) * 100) / 100,
        atcud: (obj.atcud && obj.atcud.trim()) || "",
        tipo: (obj.tipo && obj.tipo.trim()) || "",
        nif: (obj.nif && String(obj.nif).trim()) || "",
        fornecedor: (obj.fornecedor && obj.fornecedor.trim()) || ""
      };
      Logger.log("    💰 " + apis[i].nome + ": BT=" + parsed.bt + " IVA=" + parsed.iva + " Ret=" + parsed.retencoes + " Out=" + parsed.outros);
      respostas.push(parsed);
    } catch (e) {
      Logger.log("    💰 " + apis[i].nome + " ERRO: " + String(e).substring(0, 80));
    }
  }

  if (respostas.length === 0) return null;

  // Consenso: para cada campo numérico, escolher o valor mais votado
  function consensoNumero(campo) {
    var votos = {};
    for (var r = 0; r < respostas.length; r++) {
      var val = respostas[r][campo];
      var key = String(val);
      if (!votos[key]) votos[key] = 0;
      votos[key]++;
    }
    var melhor = null, melhorCount = 0;
    for (var k in votos) {
      if (votos[k] > melhorCount) { melhorCount = votos[k]; melhor = k; }
    }
    return { valor: parseFloat(melhor) || 0, votos: melhorCount };
  }

  // Consenso: para campos de texto, escolher o mais votado (não vazio)
  function consensoTexto(campo) {
    var votos = {};
    for (var r = 0; r < respostas.length; r++) {
      var val = respostas[r][campo];
      if (!val) continue;
      if (!votos[val]) votos[val] = 0;
      votos[val]++;
    }
    var melhor = "", melhorCount = 0;
    for (var k in votos) {
      if (votos[k] > melhorCount) { melhorCount = votos[k]; melhor = k; }
    }
    return melhor;
  }

  var btRes = consensoNumero("bt");
  var ivaRes = consensoNumero("iva");
  var retRes = consensoNumero("retencoes");
  var outRes = consensoNumero("outros");

  Logger.log("    💰 [CONSENSO] BT=" + btRes.valor + " (" + btRes.votos + "/" + respostas.length + ") IVA=" + ivaRes.valor + " (" + ivaRes.votos + "/" + respostas.length + ") Ret=" + retRes.valor + " Out=" + outRes.valor);

  return {
    bt: btRes.valor,
    iva: ivaRes.valor,
    retencoes: retRes.valor,
    outros: outRes.valor,
    atcud: consensoTexto("atcud"),
    tipo: consensoTexto("tipo"),
    nif: consensoTexto("nif"),
    fornecedor: consensoTexto("fornecedor")
  };
}


// ============================================================================
// HELPERS DE IA (Mistral, Groq, Gemini)
// ============================================================================

function chamarMistral(prompt) {
  var API_KEY = PropertiesService.getScriptProperties().getProperty("MISTRAL_API_KEY");
  if (!API_KEY) throw new Error("MISTRAL_API_KEY não configurada");
  var response = UrlFetchApp.fetch("https://api.mistral.ai/v1/chat/completions", {
    method: "post", contentType: "application/json",
    headers: { "Authorization": "Bearer " + API_KEY },
    payload: JSON.stringify({ model: "mistral-small-latest", messages: [{ role: "user", content: prompt }], temperature: 0 }),
    muteHttpExceptions: true
  });
  var json = JSON.parse(response.getContentText());
  if (json.error) throw new Error("Mistral: " + (json.error.message || JSON.stringify(json.error)));
  return json.choices[0].message.content;
}

function chamarGroq(prompt) {
  var API_KEY = PropertiesService.getScriptProperties().getProperty("GROQ_API_KEY");
  if (!API_KEY) throw new Error("GROQ_API_KEY não configurada");
  var response = UrlFetchApp.fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "post", contentType: "application/json",
    headers: { "Authorization": "Bearer " + API_KEY },
    payload: JSON.stringify({ model: "meta-llama/llama-4-scout-17b-16e-instruct", messages: [{ role: "user", content: prompt }], temperature: 0 }),
    muteHttpExceptions: true
  });
  var json = JSON.parse(response.getContentText());
  if (json.error) throw new Error("Groq: " + (json.error.message || JSON.stringify(json.error)));
  return json.choices[0].message.content;
}

function chamarGemini(prompt, modelo) {
  var API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!API_KEY) throw new Error("GEMINI_API_KEY não configurada");
  var modelId = modelo || "gemini-2.0-flash";
  var response = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/" + modelId + ":generateContent?key=" + API_KEY, {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0, maxOutputTokens: 500 } }),
    muteHttpExceptions: true
  });
  var json = JSON.parse(response.getContentText());
  if (json.error) throw new Error("Gemini: " + (json.error.message || JSON.stringify(json.error)));
  return json.candidates[0].content.parts[0].text;
}


// ============================================================================
// HELPERS DE NAVEGAÇÃO DRIVE
// ============================================================================

function _listarAnos(pastaRaiz) {
  var result = [];
  var it = pastaRaiz.getFolders();
  while (it.hasNext()) {
    var f = it.next();
    if (/^\d{4}$/.test(f.getName())) result.push(f);
  }
  result.sort(function(a, b) { return b.getName().localeCompare(a.getName()); }); // Decrescente: ano actual primeiro
  return result;
}

function _listarMeses(pastaAno, year) {
  var result = [];
  var re = new RegExp("^Faturas_" + CODIGO_EMPRESA + "_(\\d{2})/" + year + "$");
  var it = pastaAno.getFolders();
  while (it.hasNext()) {
    var f = it.next();
    var m = f.getName().match(re);
    if (m) result.push({ folder: f, label: m[0], month: m[1], year: year });
  }
  result.sort(function(a, b) { return b.month.localeCompare(a.month); }); // Decrescente: mês actual primeiro
  return result;
}

function _encontrarSpreadsheetDoMes(pastaMes, label) {
  // O spreadsheet chama-se "#0 - Faturas_{CODIGO}_{MM}/{YYYY}"
  var nomeEsperado = "#0 - " + label;
  var files = pastaMes.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (f.getMimeType() === MimeType.GOOGLE_SHEETS) {
      if (f.getName() === nomeEsperado || f.getName().indexOf("#0 -") === 0) {
        return SpreadsheetApp.openById(f.getId());
      }
    }
  }
  return null;
}

function _getSubFolder(parent, name) {
  var it = parent.getFolders();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName() === name) return f;
  }
  return null;
}

function _listarPDFs(folder) {
  var result = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getMimeType() === "application/pdf") result.push(f);
  }
  return result;
}


// ============================================================================
// HELPERS DE ATCUD E FORNECEDORES
// ============================================================================

/**
 * Detecta se um PDF é digital (texto nativo) ou digitalizado (scan/imagem).
 * Tenta extrair texto sem OCR — se tiver texto substancial, é digital.
 */
function _isPDFDigital(fileId) {
  var docId = null;
  try {
    var blob = DriveApp.getFileById(fileId).getBlob();
    var result = Drive.Files.insert({ title: "_TMP_DETECT_" + fileId, parents: [{ id: "root" }] }, blob, { ocr: false });
    if (!result || !result.id) return false;
    docId = result.id;
    _registerTempFile(docId);
    Utilities.sleep(300);
    var text = DocumentApp.openById(docId).getBody().getText().trim();
    DriveApp.getFileById(docId).setTrashed(true);
    return text.length > 20;
  } catch (e) {
    if (docId) { try { DriveApp.getFileById(docId).setTrashed(true); } catch (err) {} }
    return false;
  }
}

/**
 * Procura um ATCUD na spreadsheet e devolve info sobre a linha encontrada.
 * Retorna { row, numero, fileId, colLink } ou null se não encontrado.
 */
function _findATCUDDuplicado(sheet, atcudNorm) {
  var LINHA_CABECALHO = 2;
  var ultimaLinha = sheet.getLastRow();
  if (ultimaLinha <= LINHA_CABECALHO) return null;

  var colATCUD = encontraColunaNoCabecalho(sheet, "ATCUD / Nº Documento", LINHA_CABECALHO);
  if (colATCUD < 0) return null;

  var colNum = encontraColunaNoCabecalho(sheet, "Nº", LINHA_CABECALHO);
  var colLink = encontraColunaNoCabecalho(sheet, "Link documento", LINHA_CABECALHO);

  var atcuds = sheet.getRange(LINHA_CABECALHO + 1, colATCUD, ultimaLinha - LINHA_CABECALHO, 1)
    .getDisplayValues().flat();

  for (var i = 0; i < atcuds.length; i++) {
    if (String(atcuds[i]).replace(/\s/g, "") === atcudNorm) {
      var row = LINHA_CABECALHO + 1 + i;
      var numero = colNum > 0 ? String(sheet.getRange(row, colNum).getValue()) : "";

      // Extrair fileId do hyperlink existente
      var fileId = "";
      if (colLink > 0) {
        var formula = sheet.getRange(row, colLink).getFormula();
        var m = formula.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (m) fileId = m[1];
      }

      return { row: row, numero: numero, fileId: fileId, colLink: colLink };
    }
  }
  return null;
}

function _getATCUDsExistentes(sheet) {
  var LINHA_CABECALHO = 2;
  var ultimaLinha = sheet.getLastRow();
  if (ultimaLinha <= LINHA_CABECALHO) return [];

  var colATCUD = encontraColunaNoCabecalho(sheet, "ATCUD / Nº Documento", LINHA_CABECALHO);
  if (colATCUD < 0) return [];

  return sheet.getRange(LINHA_CABECALHO + 1, colATCUD, ultimaLinha - LINHA_CABECALHO, 1)
    .getDisplayValues().flat()
    .map(function(v) { return String(v).replace(/\s/g, ""); })
    .filter(Boolean);
}

function _findFornecedorByNIF(nif) {
  if (!FORNECEDORES_SHEET_ID) return null;
  try {
    var ss = SpreadsheetApp.openById(FORNECEDORES_SHEET_ID);
    var sheet = ss.getSheetByName("Fornecedores");
    if (!sheet) {
      // Tentar aba com nome parecido
      var all = ss.getSheets();
      for (var i = 0; i < all.length; i++) {
        if (all[i].getName().toLowerCase().indexOf("fornecedor") !== -1) { sheet = all[i]; break; }
      }
      if (!sheet) return null;
    }

    var nifCol = encontraColunaNoCabecalho(sheet, "NIF/NIPC fornecedor", 2);
    var nomeCol = encontraColunaNoCabecalho(sheet, "Fornecedor", 2);
    if (nifCol < 0 || nomeCol < 0) return null;

    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return null;

    var nifStr = String(nif).replace(/\D/g, '');
    var nifs = sheet.getRange(3, nifCol, lastRow - 2, 1).getDisplayValues();

    for (var i = 0; i < nifs.length; i++) {
      if (nifs[i][0].replace(/\D/g, '') === nifStr) {
        return sheet.getRange(i + 3, nomeCol).getValue();
      }
    }
  } catch (e) {
    Logger.log("Erro findFornecedorByNIF: " + e);
  }
  return null;
}

/**
 * Procura NIF na BD de fornecedores por email ou telefone (fallback quando NIF não encontrado no PDF).
 */
function _findNIFByEmailTelefone(emails, telefones) {
  if (!FORNECEDORES_SHEET_ID) return null;
  if ((!emails || !emails.length) && (!telefones || !telefones.length)) return null;

  try {
    var ss = SpreadsheetApp.openById(FORNECEDORES_SHEET_ID);
    var sheet = ss.getSheetByName("Fornecedores");
    if (!sheet) {
      var all = ss.getSheets();
      for (var i = 0; i < all.length; i++) {
        if (all[i].getName().toLowerCase().indexOf("fornecedor") !== -1) { sheet = all[i]; break; }
      }
      if (!sheet) return null;
    }

    var nifCol = encontraColunaNoCabecalho(sheet, "NIF/NIPC fornecedor", 2);
    var emailCol = encontraColunaNoCabecalho(sheet, "Email", 2);
    var telCol = encontraColunaNoCabecalho(sheet, "Telefone", 2);
    if (nifCol < 0) return null;

    var data = sheet.getDataRange().getDisplayValues();

    for (var r = 2; r < data.length; r++) {
      var row = data[r];
      var nifBD = row[nifCol - 1];

      // Match por telefone
      if (telCol > 0 && telefones && telefones.length) {
        var telBD = row[telCol - 1];
        if (telBD && telefones.indexOf(telBD) !== -1) {
          Logger.log("    BD match por telefone: " + telBD + " → NIF " + nifBD);
          return nifBD;
        }
      }

      // Match por email
      if (emailCol > 0 && emails && emails.length) {
        var emailBD = row[emailCol - 1];
        if (emailBD && emails.indexOf(emailBD) !== -1) {
          Logger.log("    BD match por email: " + emailBD + " → NIF " + nifBD);
          return nifBD;
        }
      }
    }
  } catch (e) {
    Logger.log("Erro findNIFByEmailTelefone: " + e);
  }
  return null;
}

function encontraColunaNoCabecalho(sheet, columnName, linhaDoCabecalho) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return -1;
  var headerRowValues = sheet.getRange(linhaDoCabecalho, 1, 1, lastColumn).getValues()[0];
  for (var i = 0; i < headerRowValues.length; i++) {
    if (headerRowValues[i] === columnName) return i + 1;
  }
  return -1;
}


// ============================================================================
// HELPERS DE EXTRACÇÃO REGEX (adaptados do BE - FATURAS original)
// ============================================================================

function extractATCUD(pdfText) {
  if (!pdfText) return null;
  var m = pdfText.match(/ATCUD:\s*([^\s]+)/);
  if (m) return m[1];
  m = pdfText.match(/ATCUD\s+([^\s]+)/);
  if (m) return m[1];
  return null;
}

/**
 * Valida NIF português algoritmicamente (check digit mod 11).
 */
function validaNIF(nif) {
  nif = String(nif).replace(/\s/g, '');
  if (nif.length !== 9) return false;

  // Prefixos válidos
  var first1 = nif.substring(0, 1);
  var first2 = nif.substring(0, 2);
  var validFirst1 = ['1', '2', '3', '5', '6', '8'];
  var validFirst2 = ['45', '70', '71', '72', '74', '75', '77', '79', '90', '91', '98', '99'];
  if (validFirst1.indexOf(first1) === -1 && validFirst2.indexOf(first2) === -1) return false;

  // Check digit
  var total = 0;
  for (var i = 0; i < 8; i++) total += Number(nif[i]) * (9 - i);
  var mod = total % 11;
  var checkDigit = mod < 2 ? 0 : 11 - mod;
  return checkDigit === Number(nif[8]);
}

/**
 * Extrai NIF do fornecedor (exclui NIFs da própria empresa, valida algoritmicamente).
 */
function extractNIF(text) {
  if (!text) return null;

  // Todos os números de 9 dígitos (com ou sem espaços)
  var regex = /\b(\d{9}|\d{3}\s\d{3}\s\d{3})\b/g;
  var matches = text.matchAll(regex);

  for (var match of matches) {
    var nif = match[0].replace(/\s/g, '');
    if (nif === NIF_PROPRIO_1 || nif === NIF_PROPRIO_2) continue;
    if (nif.charAt(0) === '8') continue; // empresário individual extinto
    if (validaNIF(nif)) return nif;
  }

  // Fallback: formato PTXXXXXXXXX
  var ptMatches = text.matchAll(/\bPT(\d{9})\b/g);
  for (var ptMatch of ptMatches) {
    var ptNIF = ptMatch[1];
    if (ptNIF === NIF_PROPRIO_1 || ptNIF === NIF_PROPRIO_2) continue;
    if (validaNIF(ptNIF)) return ptNIF;
  }

  return null;
}

/**
 * Extrai nome do fornecedor (Lda/S.A./CRL → fallback por labels → fallback por linha antes do NIF).
 */
function extractFornecedor(text) {
  if (!text) return null;

  // 1. Regex para nomes com sufixo empresarial
  var regex = /(.*?)(?:,\s*lda\.?|,\s*s\.?a\.?|,\s*c\.?r\.?l\.?| c\.?r\.?l\b)/ig;
  var matches = text.match(regex);
  if (matches) {
    for (var i = 0; i < matches.length; i++) {
      var m = matches[i];
      if (!m.toLowerCase().includes("darkland") && !m.toLowerCase().includes("darkpurple")) {
        return m.trim();
      }
    }
  }

  var lines = text.split(/\r?\n/).map(function(l) { return l.trim(); }).filter(function(l) { return l.length > 0; });
  var exclusoes = ["darkland", "darkpurple", NIF_PROPRIO_1, NIF_PROPRIO_2];

  function isLinhaExcluida(linha) {
    var lower = linha.toLowerCase();
    if (exclusoes.some(function(ex) { return lower.indexOf(ex.toLowerCase()) !== -1; })) return true;
    if (/^\d{4}[- ]?\d{3}\b/.test(linha)) return true;
    if (/^\d[\d\s.,\-\/]+$/.test(linha)) return true;
    if (/^(?:rua|av\.|avenida|travessa|largo|praça|estrada|urbanização|lote|bloco|piso|andar|r\/c|apartamento)\b/i.test(linha)) return true;
    return false;
  }

  // 2. Procurar labels conhecidos
  var labelRegex = /(?:emitente|raz[ãa]o\s+social|vendedor|designa[çc][ãa]o\s+social|nome\s*(?:do\s+)?fornecedor)\s*[:\-]?\s*(.+)/i;
  for (var j = 0; j < lines.length; j++) {
    var labelMatch = lines[j].match(labelRegex);
    if (labelMatch) {
      var nome = labelMatch[1].trim();
      if (nome.length >= 3 && !isLinhaExcluida(nome)) return nome;
    }
  }

  // 3. Linha antes do NIF do fornecedor
  var nifRegex = /\b(\d{9}|\d{3}\s\d{3}\s\d{3})\b/;
  for (var k = 0; k < lines.length; k++) {
    var nifMatch = lines[k].match(nifRegex);
    if (nifMatch) {
      var nifLimpo = nifMatch[0].replace(/\s/g, '');
      if (nifLimpo !== NIF_PROPRIO_1 && nifLimpo !== NIF_PROPRIO_2 && nifLimpo.charAt(0) !== '8') {
        for (var p = k - 1; p >= 0 && p >= k - 3; p--) {
          if (lines[p].length >= 3 && !isLinhaExcluida(lines[p])) return lines[p];
        }
        var textoAntes = lines[k].substring(0, lines[k].indexOf(nifMatch[0])).replace(/[:\-\s]+$/, '').trim();
        if (textoAntes.length >= 3 && !isLinhaExcluida(textoAntes)) return textoAntes;
        break;
      }
    }
  }

  return null;
}

function extractEmail(text) {
  var matches = text.match(/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g);
  return matches || [];
}

function extractTelefone(text) {
  var matches = text.match(/(?:\+351|00351)?2\d{8}/g);
  return matches || [];
}

function extractTipoDocumento(text) {
  if (!text) return null;
  text = text.replace(/[/_-]/g, ' ').replace(/\s{2,}/g, ' ');
  text = text.replace(/válido como recibo após/g, '');
  var tipos = ['fatura simplificada', 'factura simplificada', 'nota de crédito', 'fatura recibo', 'factura recibo', '2ª via', 'segunda via', 'fatura', 'factura', 'recibo de renda', 'recibo'];
  var outputs = ['Fatura simplificada', 'Fatura simplificada', 'Nota de crédito', 'Fatura-recibo', 'Fatura-recibo', '2ª via fatura', '2ª via fatura', 'Fatura', 'Fatura', 'Recibo de renda', 'Recibo'];
  var lower = text.toLowerCase();
  for (var i = 0; i < tipos.length; i++) {
    if (lower.includes(tipos[i])) return outputs[i];
  }
  return null;
}

function extractDataDocumento(pdfText) {
  if (!pdfText) return null;
  var m = pdfText.match(/\b(?:data(?:\s+de)?\s*(?:emiss[aã]o|doc(?:umento)?))\s*[:\-]?\s*(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/i);
  if (m) {
    var yyyy = m[3].length === 2 ? '20' + m[3] : m[3];
    return m[1].padStart(2, '0') + '/' + m[2].padStart(2, '0') + '/' + yyyy;
  }
  m = pdfText.match(/(\d{2})[\/\-](\d{2})[\/\-](\d{4})/);
  if (m) return m[1] + '/' + m[2] + '/' + m[3];
  m = pdfText.match(/(\d{4})[\/\-](\d{2})[\/\-](\d{2})/);
  if (m) return m[3] + '/' + m[2] + '/' + m[1];
  return null;
}


// ============================================================================
// CATALOGAÇÃO DE RECIBOS (#4)
// ============================================================================
// Recibos não criam linhas novas — procuram a fatura correspondente por valor
// e preenchem a coluna "Recibo" com link para o ficheiro.

function _catalogarRecibo(file, tabsFaturas, pastaDestino, paraCatalogar) {
  var fileName = file.getName();
  Logger.log("    🧾 Recibo: " + fileName);

  var textoPDF = "";
  try { textoPDF = convertPDFToText(file.getId(), ['pt', 'en', null]) || ""; }
  catch (e) { return { sucesso: false, mensagem: fileName + ": Falha OCR" }; }
  if (!textoPDF.trim()) return { sucesso: false, mensagem: fileName + ": PDF vazio" };

  var valores = extractAmountFromRecibos(textoPDF);
  if (!valores || valores.length === 0) {
    return { sucesso: false, mensagem: fileName + ": Nenhum valor extraído" };
  }

  var LINHA_CABECALHO = 2;

  for (var p = 0; p < tabsFaturas.length; p++) {
    var sheet = tabsFaturas[p];
    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha <= LINHA_CABECALHO) continue;

    var colValor = encontraColunaNoCabecalho(sheet, "Valor total", LINHA_CABECALHO);
    var colRecibo = encontraColunaNoCabecalho(sheet, "Recibo", LINHA_CABECALHO);
    var colNum = encontraColunaNoCabecalho(sheet, "Nº", LINHA_CABECALHO);
    if (colValor <= 0 || colRecibo <= 0 || colNum <= 0) continue;

    var valoresSheet = sheet.getRange(LINHA_CABECALHO + 1, colValor, ultimaLinha - LINHA_CABECALHO, 1).getValues();

    for (var i = 0; i < valoresSheet.length; i++) {
      var linhaAtual = LINHA_CABECALHO + 1 + i;
      var valorFatura = parseFloat(Number(valoresSheet[i][0]).toFixed(2));
      if (isNaN(valorFatura)) continue;

      var celRecibo = sheet.getRange(linhaAtual, colRecibo);
      if (!celRecibo.isBlank()) continue;

      for (var j = 0; j < valores.length; j++) {
        var valorRecibo = parseFloat(valores[j]);
        if (valorRecibo === 0) continue;

        if (Math.abs(valorFatura - valorRecibo) < 0.01) {
          var numeroDoc = sheet.getRange(linhaAtual, colNum).getValue();
          celRecibo.setFormula('=HYPERLINK("' + file.getUrl() + '";"LINK")');
          SpreadsheetApp.flush();

          var novoNome = "REC" + numeroDoc + ".pdf";
          file.setName(novoNome);
          file.moveTo(pastaDestino);

          Logger.log("    ✅ Recibo → fatura " + numeroDoc + " (valor " + valorRecibo + ")");
          return { sucesso: true, mensagem: novoNome + " → fatura " + numeroDoc };
        }
      }
    }
  }

  return { sucesso: false, mensagem: fileName + ": Nenhuma fatura correspondente" };
}

function extractAmountFromRecibos(content) {
  if (!content) return null;
  var amountRegex = /(?:\b|\D)(\d{1,3}(?:[.,]\d{3})*(?:[,\.]\d{2})?|\d{1,3}(?:[,\.]\d{2})?)(?:\s*euros?|\s*€)?(?:\b|\D)/gi;
  var match;
  var amounts = [];
  while ((match = amountRegex.exec(content)) !== null) {
    amounts.push(match[1].replace(',', '.'));
  }
  amounts = amounts.filter(function(v, i, a) { return a.indexOf(v) === i; }); // deduplica
  return amounts.length > 0 ? amounts : null;
}


// ============================================================================
// CATALOGAÇÃO DE COMPROVATIVOS (#5)
// ============================================================================
// Comprovativos não criam linhas novas — procuram a fatura correspondente
// por ATCUD (prioritário) ou por valor, e preenchem as colunas
// "Comprovativo de pagamento", "Data de pagamento" e "Método de pagamento".

function _catalogarComprovativo(file, tabsFaturas, pastaDestino, paraCatalogar) {
  var fileName = file.getName();
  Logger.log("    💳 Comprovativo: " + fileName);

  var textoPDF = "";
  try { textoPDF = convertPDFToText(file.getId(), ['pt', 'en', null]) || ""; }
  catch (e) { return { sucesso: false, mensagem: fileName + ": Falha OCR" }; }
  if (!textoPDF.trim()) return { sucesso: false, mensagem: fileName + ": PDF vazio" };

  var valorLiquido = parseFloat(extractAmountFromPayslip(textoPDF));
  var dataPagamento = extractDateFromPayslip(textoPDF);
  var atcudComprovativo = extractATCUDComprovativos(textoPDF);

  // Detectar banco
  var bancoDisplay = "Outro";
  if (/CREDITO\s*AGRICOLA/i.test(textoPDF)) bancoDisplay = "Crédito Agrícola";
  else if (/NET24/i.test(textoPDF)) bancoDisplay = "Montepio";
  else if (/MILLENNIUM/i.test(textoPDF) || /BCP/i.test(textoPDF)) bancoDisplay = "Millennium BCP";
  else if (/CAIXA\s*GERAL/i.test(textoPDF) || /CGD/i.test(textoPDF)) bancoDisplay = "CGD";

  Logger.log("    Valor=" + valorLiquido + " Data=" + dataPagamento + " ATCUD=" + atcudComprovativo + " Banco=" + bancoDisplay);

  var LINHA_CABECALHO = 2;
  var candidatosValor = [];

  // 1º: Match por ATCUD
  if (atcudComprovativo) {
    var atcudNorm = String(atcudComprovativo).replace(/\s/g, "");

    for (var p = 0; p < tabsFaturas.length; p++) {
      var sheet = tabsFaturas[p];
      var ultimaLinha = sheet.getLastRow();
      if (ultimaLinha <= LINHA_CABECALHO) continue;

      var colATCUD = encontraColunaNoCabecalho(sheet, "ATCUD / Nº Documento", LINHA_CABECALHO);
      var colMetodo = encontraColunaNoCabecalho(sheet, "Método de pagamento", LINHA_CABECALHO);
      var colDataPag = encontraColunaNoCabecalho(sheet, "Data de pagamento", LINHA_CABECALHO);
      var colComp = encontraColunaNoCabecalho(sheet, "Comprovativo de pagamento", LINHA_CABECALHO);
      var colNum = encontraColunaNoCabecalho(sheet, "Nº", LINHA_CABECALHO);
      if (colATCUD <= 0 || colComp <= 0 || colNum <= 0) continue;

      for (var i = LINHA_CABECALHO + 1; i <= ultimaLinha; i++) {
        var atcudLinha = String(sheet.getRange(i, colATCUD).getDisplayValue()).replace(/\s/g, "");
        if (!atcudLinha || atcudLinha !== atcudNorm) continue;
        if (!sheet.getRange(i, colComp).isBlank()) continue;

        var numDoc = sheet.getRange(i, colNum).getValue();
        if (colMetodo > 0) sheet.getRange(i, colMetodo).setValue(bancoDisplay);
        if (colDataPag > 0) sheet.getRange(i, colDataPag).setValue(dataPagamento).setNumberFormat('dd/MM/yyyy');
        sheet.getRange(i, colComp).setFormula('=HYPERLINK("' + file.getUrl() + '";"LINK")');
        SpreadsheetApp.flush();

        var novoNome = "COMP" + numDoc + ".pdf";
        file.setName(novoNome);
        file.moveTo(pastaDestino);

        Logger.log("    ✅ Comprovativo → fatura " + numDoc + " (ATCUD match)");
        return { sucesso: true, mensagem: novoNome + " → fatura " + numDoc + " (ATCUD)" };
      }
    }
  }

  // 2º: Match por valor (se ATCUD não encontrou)
  if (!isNaN(valorLiquido) && valorLiquido > 0) {
    for (var p2 = 0; p2 < tabsFaturas.length; p2++) {
      var sheet2 = tabsFaturas[p2];
      var ultimaLinha2 = sheet2.getLastRow();
      if (ultimaLinha2 <= LINHA_CABECALHO) continue;

      var colValor2 = encontraColunaNoCabecalho(sheet2, "Valor total", LINHA_CABECALHO);
      var colMetodo2 = encontraColunaNoCabecalho(sheet2, "Método de pagamento", LINHA_CABECALHO);
      var colDataPag2 = encontraColunaNoCabecalho(sheet2, "Data de pagamento", LINHA_CABECALHO);
      var colComp2 = encontraColunaNoCabecalho(sheet2, "Comprovativo de pagamento", LINHA_CABECALHO);
      var colNum2 = encontraColunaNoCabecalho(sheet2, "Nº", LINHA_CABECALHO);
      if (colValor2 <= 0 || colComp2 <= 0 || colNum2 <= 0) continue;

      var valoresSheet = sheet2.getRange(LINHA_CABECALHO + 1, colValor2, ultimaLinha2 - LINHA_CABECALHO, 1).getValues();
      for (var j = 0; j < valoresSheet.length; j++) {
        var linhaAtual = LINHA_CABECALHO + 1 + j;
        if (sheet2.getRange(linhaAtual, colComp2).isBlank() && Math.abs(valoresSheet[j][0] - valorLiquido) < 0.01) {
          candidatosValor.push({
            sheet: sheet2, row: linhaAtual,
            numDoc: sheet2.getRange(linhaAtual, colNum2).getValue(),
            colMetodo: colMetodo2, colDataPag: colDataPag2, colComp: colComp2
          });
        }
      }
    }

    if (candidatosValor.length === 1) {
      var cv = candidatosValor[0];
      if (cv.colMetodo > 0) cv.sheet.getRange(cv.row, cv.colMetodo).setValue(bancoDisplay);
      if (cv.colDataPag > 0) cv.sheet.getRange(cv.row, cv.colDataPag).setValue(dataPagamento).setNumberFormat('dd/MM/yyyy');
      cv.sheet.getRange(cv.row, cv.colComp).setFormula('=HYPERLINK("' + file.getUrl() + '";"LINK")');
      SpreadsheetApp.flush();

      var novoNome2 = "COMP" + cv.numDoc + ".pdf";
      file.setName(novoNome2);
      file.moveTo(pastaDestino);

      Logger.log("    ✅ Comprovativo → fatura " + cv.numDoc + " (valor match)");
      return { sucesso: true, mensagem: novoNome2 + " → fatura " + cv.numDoc + " (valor)" };

    } else if (candidatosValor.length > 1) {
      Logger.log("    ⚠️ Múltiplas faturas possíveis (" + candidatosValor.length + ") para valor " + valorLiquido);
      return { sucesso: false, mensagem: fileName + ": Múltiplas faturas para valor " + valorLiquido };
    }
  }

  return { sucesso: false, mensagem: fileName + ": Nenhuma fatura correspondente (ATCUD=" + (atcudComprovativo || "?") + " Valor=" + valorLiquido + ")" };
}

function extractAmountFromPayslip(content) {
  if (!content) return null;
  var amountRegex = /-?\d{1,3}(?:[ .]?\d{3})*(?:[,]\d{2})/g;
  var match;
  var amounts = [];
  while ((match = amountRegex.exec(content)) !== null) {
    var number = match[0];
    if (number.replace(/[^0-9]/g, '').length > 8) continue;
    var indexOfSaldo = content.toLowerCase().indexOf("saldo");
    if (indexOfSaldo !== -1 && indexOfSaldo < match.index) continue;
    var sanitized = number.replace(/\./g, '').replace(',', '.').replace('-', '').replace(' ', '');
    amounts.push(sanitized);
  }
  return amounts.length > 0 ? amounts[0] : null;
}

function extractDateFromPayslip(content) {
  if (!content) return null;
  var words = content.split(/\s+/);
  for (var i = 0; i < words.length; i++) {
    var word = words[i];
    if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(word)) {
      var p = word.split(/[-/]/);
      return p[2] + '/' + p[1] + '/' + p[0];
    }
    if (/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(word)) {
      var p2 = word.split(/[-/]/);
      return p2[0] + '/' + p2[1] + '/' + p2[2];
    }
  }
  var m = content.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (m) return ("0" + m[1]).slice(-2) + "/" + ("0" + m[2]).slice(-2) + "/" + m[3];
  return null;
}

function extractATCUDComprovativos(pdfText) {
  if (!pdfText) return null;
  var text = pdfText.replace(/\s+/g, ' ').toUpperCase();
  var regex = /\b([A-Z0-9]{8,}-\d{2,})\b/g;
  var match;
  var candidatos = [];
  while ((match = regex.exec(text)) !== null) {
    candidatos.push(match[1]);
  }
  return candidatos.length > 0 ? candidatos[0] : null;
}


// ============================================================================
// CLEANUP DE FICHEIROS TEMPORÁRIOS (OCR + detecção digital)
// ============================================================================

var __tempFileIds = [];

function _registerTempFile(fileId) {
  __tempFileIds.push(fileId);
}

/** Apaga todos os ficheiros temporários criados durante a execução */
function _cleanupTempFiles() {
  for (var i = 0; i < __tempFileIds.length; i++) {
    try { DriveApp.getFileById(__tempFileIds[i]).setTrashed(true); } catch (e) {}
  }
  // Fallback: procurar e apagar ficheiros _TMP_ órfãos no Drive
  try {
    var orphans = DriveApp.searchFiles("title contains '_TMP_' and trashed=false");
    while (orphans.hasNext()) {
      var f = orphans.next();
      if (f.getName().indexOf("_TMP_") === 0) {
        f.setTrashed(true);
        Logger.log("🧹 Temp órfão apagado: " + f.getName());
      }
    }
  } catch (e) {}
  __tempFileIds = [];
}


// ============================================================================
// ALGORITMO DE CONSENSO DE DATAS (6 fontes → data mais votada)
// ============================================================================

function _safeDate_(dd, mm, yyyy) {
  var y = Number(yyyy), m = Number(mm), d = Number(dd);
  if (!y || !m || !d || y < 2000) return null;
  var dt = new Date(y, m - 1, d);
  var today = new Date();
  if (isNaN(dt.getTime()) || dt > today) return null;
  if (dt.getFullYear() !== y || (dt.getMonth() + 1) !== m || dt.getDate() !== d) return null;
  return String(d).padStart(2, '0') + '/' + String(m).padStart(2, '0') + '/' + String(y);
}

function _extrairDataDoNomeFicheiro(fileName) {
  var m = fileName.match(/[_](\d{4})-(\d{2})[_]/);
  if (m) return { month: m[2], year: m[1] };
  return null;
}

function _normalizarDataIA(resultado) {
  if (!resultado) return null;
  var limpo = resultado.trim().replace(/```/g, "").replace(/["""]/g, "").trim();
  var m = limpo.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return _safeDate_(m[1], m[2], m[3]);
  return null;
}

function _extrairDataViaModelo(texto, nomeFuncao, funcaoChamar) {
  var prompt = "Extraia apenas a data de emissão do seguinte texto de um documento.\n" +
    "Retorne SOMENTE a data no formato DD/MM/AAAA, sem mais nada.\n" +
    'Se não encontrar, retorne "Não encontrada".\n\nTexto:\n' + texto;
  try {
    var resultado = funcaoChamar(prompt);
    var data = _normalizarDataIA(resultado);
    Logger.log("  [CONSENSO] " + nomeFuncao + ": " + (data || "sem data"));
    return data;
  } catch (e) {
    Logger.log("  [CONSENSO] " + nomeFuncao + " ERRO: " + String(e).substring(0, 80));
    return null;
  }
}

function _consensoData(fileName, textoPDF) {
  Logger.log("🗳️ [CONSENSO] A iniciar para: " + fileName);
  var textoParaIA = textoPDF.substring(0, 4000);
  var votos = {};
  var fontes = {};

  function registarVoto(data, fonte) {
    if (!data) return;
    if (!votos[data]) { votos[data] = 0; fontes[data] = []; }
    votos[data]++;
    fontes[data].push(fonte);
  }

  // FONTE 1: Regex OCR
  var dataRegex = extractDataDocumento(textoPDF);
  registarVoto(dataRegex, "Regex");

  // FONTE 2: IA Mistral
  var dataMistral = _extrairDataViaModelo(textoParaIA, "Mistral", chamarMistral);
  registarVoto(dataMistral, "Mistral");

  // FONTE 3: IA Groq
  var dataGroq = _extrairDataViaModelo(textoParaIA, "Groq", chamarGroq);
  registarVoto(dataGroq, "Groq");

  // FONTE 4: IA Gemini 2.5 Flash
  var dataGeminiFlash = _extrairDataViaModelo(textoParaIA, "Gemini 2.0 Flash", function(p) { return chamarGemini(p, "gemini-2.0-flash"); });
  registarVoto(dataGeminiFlash, "Gemini 2.0 Flash");

  // FONTE 5: IA Gemini 3.1 Flash Lite
  var dataGeminiLite = _extrairDataViaModelo(textoParaIA, "Gemini Lite", function(p) { return chamarGemini(p, "gemini-3.1-flash-lite-preview"); });
  registarVoto(dataGeminiLite, "Gemini Lite");

  // FONTE 6: Nome do ficheiro (voto parcial — só MM/YYYY)
  var nomeInfo = _extrairDataDoNomeFicheiro(fileName);
  if (nomeInfo) {
    var votouPorNome = false;
    for (var d in votos) {
      var parts = d.split("/");
      if (parts[1] === nomeInfo.month && parts[2] === nomeInfo.year) {
        registarVoto(d, "Nome");
        votouPorNome = true;
        break;
      }
    }
    if (!votouPorNome) Logger.log("  [CONSENSO] Nome: " + nomeInfo.month + "/" + nomeInfo.year + " (sem match de dia)");
  } else {
    Logger.log("  [CONSENSO] Nome: formato não reconhecido");
  }

  // APURAMENTO
  var melhorData = null;
  var melhorVotos = 0;
  for (var d in votos) {
    if (votos[d] > melhorVotos) { melhorVotos = votos[d]; melhorData = d; }
  }

  if (melhorData) {
    Logger.log("🗳️ [CONSENSO] RESULTADO: " + melhorData + " (" + melhorVotos + "/6 votos: " + fontes[melhorData].join(", ") + ")");
  } else {
    Logger.log("🗳️ [CONSENSO] RESULTADO: SEM DATA");
  }

  return { data: melhorData, votos: melhorVotos, fontes: melhorData ? fontes[melhorData] : [], todas: votos };
}


// ============================================================================
// OCR (conversão PDF → texto via Google Drive OCR)
// ============================================================================

function convertPDFToText(fileId, languages) {
  if (!fileId) throw new Error("convertPDFToText: fileId em falta.");
  if (!Array.isArray(languages)) languages = [languages || "pt"];

  var file = DriveApp.getFileById(fileId);
  var mime = file.getMimeType();

  if (mime === MimeType.GOOGLE_DOCS || mime === "application/vnd.google-apps.document") {
    return DocumentApp.openById(fileId).getBody().getText();
  }

  for (var i = 0; i < languages.length; i++) {
    var lang = languages[i];
    var maxTentativas = 3;
    var esperaMs = 2000;

    for (var tentativa = 1; tentativa <= maxTentativas; tentativa++) {
      var docId = null;
      try {
        var blob = file.getBlob();
        var resource = { title: "_TMP_OCR_" + file.getName() };
        var options = { ocr: true, ocrLanguage: lang || undefined };

        var ocrResult;
        try {
          ocrResult = Drive.Files.insert(resource, blob, options);
        } catch (e) {
          resource.mimeType = "application/pdf";
          ocrResult = Drive.Files.insert(resource, blob, options);
        }

        if (!ocrResult || !ocrResult.id) throw new Error("ID nulo no OCR.");
        docId = ocrResult.id;
        _registerTempFile(docId);
        Utilities.sleep(500);
        var doc = DocumentApp.openById(docId);
        var textContent = doc.getBody().getText();
        DriveApp.getFileById(docId).setTrashed(true);

        if (textContent && textContent.trim().length > 0) return textContent;
        break;

      } catch (e) {
        if (docId) { try { DriveApp.getFileById(docId).setTrashed(true); } catch (err) {} }
        var msg = (e && e.message) ? e.message : String(e);

        if (msg.includes("User rate limit exceeded") || msg.includes("403")) {
          if (tentativa < maxTentativas) {
            Utilities.sleep(esperaMs);
            esperaMs *= 2;
            continue;
          }
        }
        break;
      }
    }
  }
  return "";
}


// ============================================================================
// EMAIL DE RESUMO
// ============================================================================

function _enviarResumo(catalogados, erros, resumo, parcial) {
  if (catalogados === 0 && erros === 0) return;
  var assunto = (parcial ? "[Parcial] " : "") + "Catalogação (" + CODIGO_EMPRESA + "): " + catalogados + " ok, " + erros + " erros";
  var corpo = "Catalogados: " + catalogados + "\nErros: " + erros + "\n" + resumo;
  try {
    MailApp.sendEmail(EMAIL_NOTIFICACAO, assunto, corpo);
  } catch (e) {
    Logger.log("Erro ao enviar email: " + e);
  }
}
