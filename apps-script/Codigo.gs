/**
 * Gerador de Relatórios / Recibos - Carreta Digital
 * ------------------------------------------------------------------
 * Apps Script vinculado à planilha do Google Sheets.
 *
 * Fluxo:
 *  1. O usuário abre o menu "Gerador de Relatórios" e o painel lateral.
 *  2. Escolhe o mês de competência.
 *  3. Marca "Todos os colaboradores" OU seleciona individualmente.
 *  4. O script copia o modelo (Google Doc), substitui os campos
 *     <<...>> com os dados de cada colaborador e exporta um PDF por
 *     colaborador na pasta de destino do Drive.
 *
 * IMPORTANTE: preencha as 2 constantes em CONFIG abaixo antes de usar.
 */

var CONFIG = {
  // Nome da aba que contém os dados dos colaboradores.
  SHEET_NAME: 'Dados_Destino',

  // ID do Google DOC modelo (não é o .docx! converta para Google Docs).
  // Ex.: na URL https://docs.google.com/document/d/AQUI_O_ID/edit
  TEMPLATE_DOC_ID: 'COLE_AQUI_O_ID_DO_DOCUMENTO_MODELO',

  // ID da pasta do Drive onde os PDFs serão salvos.
  // Ex.: na URL https://drive.google.com/drive/folders/AQUI_O_ID
  OUTPUT_FOLDER_ID: 'COLE_AQUI_O_ID_DA_PASTA_DESTINO',

  // Token compartilhado entre o site (GitHub Pages) e este Web App.
  // Use uma string longa e aleatória. DEVE ser igual ao API_TOKEN do config.js.
  API_TOKEN: '2ZNfDJkg8hwE3tzAEXvZ7Q5k7LMCLTFa'
};

// Nomes EXATOS dos cabeçalhos das colunas na aba Dados_Destino.
var COLUNAS = {
  id:        'Id',
  cpf:       'CNPJ / CPF / IG Favorecido',
  cargo:     'Cargo',
  valor:     'Valor DL',
  descricao: 'DESCRIÇÃO ATIVIDADES',
  nome:      'Nome Completo',
  telefone:  'Telefone',
  email:     'Email',
  endereco:  'Endereço',
  cep:       'CEP',
  cidade:    'Cidade'
};

// ---------------------------------------------------------------------------
// MENU / ABERTURA DO PAINEL
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📄 Gerador de Relatórios')
    .addItem('Abrir painel', 'abrirPainel')
    .addToUi();
}

function abrirPainel() {
  var html = HtmlService.createHtmlOutputFromFile('Painel')
    .setTitle('Gerador de Relatórios')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---------------------------------------------------------------------------
// API WEB APP (consumida pela tela hospedada no GitHub Pages)
// ---------------------------------------------------------------------------
//
// Publique este script como Web App (Implantar > Nova implantação > Web app):
//   - "Executar como": Eu (dono da planilha)
//   - "Quem tem acesso": Qualquer pessoa
// Copie a URL .../exec para o config.js do site.
//
// Observação sobre CORS: o ContentService do Apps Script não permite definir
// cabeçalhos. Por isso a tela faz POST com Content-Type "text/plain" (requisição
// "simples", sem preflight); a resposta vem do redirecionamento para
// googleusercontent.com, que já inclui Access-Control-Allow-Origin: *.

function doGet(e)  { return rotear_(e); }
function doPost(e) { return rotear_(e); }

function rotear_(e) {
  try {
    var req = parseRequisicao_(e);

    if (!req.token || req.token !== CONFIG.API_TOKEN) {
      return jsonOut_({ ok: false, erro: 'Token inválido ou ausente.' });
    }

    var resposta;
    switch (req.action) {
      case 'listar':
        resposta = { ok: true, colaboradores: getColaboradores() };
        break;
      case 'gerar':
        resposta = gerarRelatorios({
          todos:      !!req.todos,
          rowIndexes: req.rowIndexes || [],
          mes:        req.mes,
          ano:        req.ano
        });
        break;
      default:
        resposta = { ok: false, erro: 'Ação desconhecida: ' + req.action };
    }
    return jsonOut_(resposta);
  } catch (err) {
    return jsonOut_({ ok: false, erro: String((err && err.message) || err) });
  }
}

/** Lê os parâmetros tanto de POST (JSON) quanto de GET (query string). */
function parseRequisicao_(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (x) { /* ignora */ }
  }
  var p = (e && e.parameter) || {};
  var obj = {};
  Object.keys(p).forEach(function (k) { obj[k] = p[k]; });
  if (obj.rowIndexes) { try { obj.rowIndexes = JSON.parse(obj.rowIndexes); } catch (x) {} }
  if (obj.todos != null) obj.todos = (obj.todos === 'true' || obj.todos === true);
  if (obj.mes != null) obj.mes = parseInt(obj.mes, 10);
  if (obj.ano != null) obj.ano = parseInt(obj.ano, 10);
  return obj;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// LEITURA DOS DADOS (chamado pelo painel/API)
// ---------------------------------------------------------------------------

/**
 * Lê a aba Dados_Destino e devolve a lista de colaboradores
 * para popular o painel.
 * @return {{id, nome, cargo}[]}
 */
function getColaboradores() {
  var dados = lerPlanilha_();
  return dados.linhas.map(function (linha, i) {
    return {
      rowIndex: i, // posição na lista (usada para seleção)
      id:    String(linha[COLUNAS.id] || ''),
      nome:  String(linha[COLUNAS.nome] || '(sem nome)'),
      cargo: String(linha[COLUNAS.cargo] || '')
    };
  });
}

/**
 * Lê a planilha e retorna { headers, linhas } onde cada linha é um
 * objeto { nomeColuna: valor }.
 */
function lerPlanilha_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error('Aba "' + CONFIG.SHEET_NAME + '" não encontrada na planilha.');
  }

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return { headers: values[0] || [], linhas: [] };
  }

  var headers = values[0].map(function (h) { return String(h).trim(); });
  var linhas = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    // Ignora linhas totalmente vazias.
    var temConteudo = row.some(function (c) { return String(c).trim() !== ''; });
    if (!temConteudo) continue;

    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      obj[headers[c]] = row[c];
    }
    linhas.push(obj);
  }
  return { headers: headers, linhas: linhas };
}

// ---------------------------------------------------------------------------
// GERAÇÃO DOS RELATÓRIOS (chamado pelo painel)
// ---------------------------------------------------------------------------

/**
 * @param {{ rowIndexes:number[], todos:boolean, mes:string, ano:number }} payload
 * @return {{ ok:boolean, gerados:{nome,url}[], erros:string[], pastaUrl:string }}
 */
function gerarRelatorios(payload) {
  validarConfig_();

  var dados = lerPlanilha_();
  var linhas = dados.linhas;

  // Seleciona quais linhas processar.
  var selecionadas;
  if (payload.todos) {
    selecionadas = linhas;
  } else {
    selecionadas = (payload.rowIndexes || []).map(function (i) { return linhas[i]; })
                                             .filter(function (l) { return !!l; });
  }

  if (selecionadas.length === 0) {
    throw new Error('Nenhum colaborador selecionado.');
  }

  var mesCompetencia = formatarMesCompetencia_(payload.mes, payload.ano); // ex.: "maio de 2026"
  var reciboPrefixo = formatarPrefixoRecibo_(payload.mes, payload.ano);   // ex.: "202605"
  var pasta = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  var template = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);

  var gerados = [];
  var erros = [];

  selecionadas.forEach(function (linha) {
    var nome = String(linha[COLUNAS.nome] || 'colaborador');
    try {
      var url = gerarUmRelatorio_(linha, mesCompetencia, reciboPrefixo, template, pasta);
      gerados.push({ nome: nome, url: url });
    } catch (e) {
      erros.push(nome + ': ' + e.message);
    }
  });

  return {
    ok: erros.length === 0,
    gerados: gerados,
    erros: erros,
    pastaUrl: pasta.getUrl()
  };
}

/**
 * Gera um único PDF a partir do modelo e devolve a URL do PDF.
 */
function gerarUmRelatorio_(linha, mesCompetencia, reciboPrefixo, template, pasta) {
  var nome = String(linha[COLUNAS.nome] || 'colaborador');
  var nRecibo = reciboPrefixo + '-' + String(linha[COLUNAS.id] || ''); // ex.: "202606-1"
  var nomeArquivo = 'Recibo ' + sanitizar_(nRecibo) + ' - ' + sanitizar_(nome);

  // 1. Copia o modelo como Google Doc temporário.
  var copia = template.makeCopy(nomeArquivo, pasta);
  var doc = DocumentApp.openById(copia.getId());

  // 2. Monta os valores de substituição.
  var valorNum = parseValor_(linha[COLUNAS.valor]);
  var subs = {
    'N_Recibo':                 nRecibo,
    'mes_competencia':          mesCompetencia,
    'ID':                       linha[COLUNAS.id],
    'Nome\\s+Completo':         linha[COLUNAS.nome],
    'Telefone':                 linha[COLUNAS.telefone],
    'Cargo':                    linha[COLUNAS.cargo],
    'Email':                    linha[COLUNAS.email],
    'Endereço':            linha[COLUNAS.endereco],
    'CEP':                      linha[COLUNAS.cep],
    'Cidade':                   linha[COLUNAS.cidade],
    'CPF':                      linha[COLUNAS.cpf],
    'Descrição\\s+das\\s+Atividades': linha[COLUNAS.descricao],
    'valor_extenso':            valorPorExtenso_(valorNum),
    // O modelo já tem "R$ " antes de <<valor>>, então aqui vai só o número.
    'valor':                    formatarNumero_(valorNum)
  };

  // 3. Substitui no corpo, cabeçalho e rodapé.
  var alvos = [doc.getBody()];
  if (doc.getHeader()) alvos.push(doc.getHeader());
  if (doc.getFooter()) alvos.push(doc.getFooter());

  alvos.forEach(function (alvo) {
    Object.keys(subs).forEach(function (token) {
      var valor = subs[token] == null ? '' : String(subs[token]);
      // Tolera espaços internos: <<  Cargo  >>
      var regex = '<<\\s*' + token + '\\s*>>';
      alvo.replaceText(regex, valor);
    });
  });

  doc.saveAndClose();

  // 4. Exporta para PDF e remove o Doc temporário.
  var pdfBlob = DriveApp.getFileById(copia.getId()).getAs('application/pdf');
  var pdf = pasta.createFile(pdfBlob).setName(nomeArquivo + '.pdf');
  copia.setTrashed(true);

  return pdf.getUrl();
}

// ---------------------------------------------------------------------------
// HELPERS
// ---------------------------------------------------------------------------

function validarConfig_() {
  if (!CONFIG.TEMPLATE_DOC_ID || CONFIG.TEMPLATE_DOC_ID.indexOf('COLE_AQUI') === 0) {
    throw new Error('Configure CONFIG.TEMPLATE_DOC_ID com o ID do Google Doc modelo.');
  }
  if (!CONFIG.OUTPUT_FOLDER_ID || CONFIG.OUTPUT_FOLDER_ID.indexOf('COLE_AQUI') === 0) {
    throw new Error('Configure CONFIG.OUTPUT_FOLDER_ID com o ID da pasta de destino.');
  }
}

var MESES_PT = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
                'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

function getMeses() {
  return MESES_PT;
}

/** Prefixo do recibo no formato AAAAMM. Ex.: maio/2026 -> "202605". */
function formatarPrefixoRecibo_(mes, ano) {
  var idx;
  if (typeof mes === 'number' || /^\d+$/.test(String(mes))) {
    idx = parseInt(mes, 10);
  } else {
    idx = MESES_PT.indexOf(String(mes).toLowerCase());
  }
  var mm = ('0' + (idx + 1)).slice(-2);
  return '' + ano + mm;
}

function formatarMesCompetencia_(mes, ano) {
  // mes vem como índice 0-11 (string ou número) ou nome.
  var nome;
  if (typeof mes === 'number' || /^\d+$/.test(String(mes))) {
    nome = MESES_PT[parseInt(mes, 10)] || '';
  } else {
    nome = String(mes).toLowerCase();
  }
  return nome + ' de ' + ano; // ex.: "maio de 2026"
}

function sanitizar_(texto) {
  return String(texto).replace(/[\\\/:*?"<>|]/g, '-').trim();
}

/** Converte o valor da planilha (número ou string "R$ 1.500,00") em número. */
function parseValor_(valor) {
  if (typeof valor === 'number') return valor;
  if (valor == null || valor === '') return 0;
  var s = String(valor).replace(/[^\d,.-]/g, '');
  // Formato brasileiro: 1.500,00 -> 1500.00
  if (s.indexOf(',') > -1) {
    s = s.replace(/\./g, '').replace(',', '.');
  }
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/** Formata número no padrão brasileiro sem símbolo: 1500 -> "1.500,00". */
function formatarNumero_(n) {
  var negativo = n < 0;
  n = Math.abs(Math.round(n * 100) / 100);
  var inteiro = Math.floor(n);
  var centavos = Math.round((n - inteiro) * 100);
  var intStr = inteiro.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  var centStr = (centavos < 10 ? '0' : '') + centavos;
  return (negativo ? '-' : '') + intStr + ',' + centStr;
}

/** Formata como moeda brasileira: 1500 -> "R$ 1.500,00". */
function formatarMoeda_(n) {
  return 'R$ ' + formatarNumero_(n);
}

// ---- Valor por extenso (Real brasileiro) ----------------------------------

var EXT_UNIDADES = ['zero', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis',
                    'sete', 'oito', 'nove', 'dez', 'onze', 'doze', 'treze',
                    'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito',
                    'dezenove'];
var EXT_DEZENAS = ['', '', 'vinte', 'trinta', 'quarenta', 'cinquenta',
                   'sessenta', 'setenta', 'oitenta', 'noventa'];
var EXT_CENTENAS = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos',
                    'quinhentos', 'seiscentos', 'setecentos', 'oitocentos',
                    'novecentos'];

/** Converte 0..999 em extenso. */
function extensoGrupo_(n) {
  if (n === 0) return '';
  if (n === 100) return 'cem';
  var partes = [];
  var c = Math.floor(n / 100);
  var resto = n % 100;
  if (c > 0) partes.push(EXT_CENTENAS[c]);
  if (resto > 0) {
    if (resto < 20) {
      partes.push(EXT_UNIDADES[resto]);
    } else {
      var d = Math.floor(resto / 10);
      var u = resto % 10;
      partes.push(EXT_DEZENAS[d] + (u > 0 ? ' e ' + EXT_UNIDADES[u] : ''));
    }
  }
  return partes.join(' e ');
}

/** Converte inteiro em extenso (até bilhões). */
function extensoInteiro_(n) {
  if (n === 0) return 'zero';
  var grupos = [];
  var escalas = [
    { s: ['', ''] },                         // unidades
    { s: [' mil', ' mil'] },                 // milhares
    { s: [' milhão', ' milhões'] },          // milhões
    { s: [' bilhão', ' bilhões'] }           // bilhões
  ];

  var partes = [];
  var i = 0;
  while (n > 0 && i < escalas.length) {
    var grupo = n % 1000;
    if (grupo > 0) {
      var texto = extensoGrupo_(grupo);
      var sufixo = escalas[i].s[grupo === 1 ? 0 : 1];
      // "um mil" -> "mil"
      if (i === 1 && grupo === 1) texto = '';
      partes.unshift((texto + sufixo).trim());
    }
    n = Math.floor(n / 1000);
    i++;
  }
  return partes.join(' e ').trim();
}

/** Valor monetário por extenso. Ex.: 1500.5 -> "mil e quinhentos reais e cinquenta centavos". */
function valorPorExtenso_(valor) {
  valor = Math.round(valor * 100) / 100;
  var reais = Math.floor(valor);
  var centavos = Math.round((valor - reais) * 100);

  var partes = [];
  if (reais > 0) {
    partes.push(extensoInteiro_(reais) + (reais === 1 ? ' real' : ' reais'));
  } else if (centavos === 0) {
    partes.push('zero reais');
  }
  if (centavos > 0) {
    partes.push(extensoInteiro_(centavos) + (centavos === 1 ? ' centavo' : ' centavos'));
  }
  return partes.join(' e ');
}
