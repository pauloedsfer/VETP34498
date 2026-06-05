/**
 * CONTROLADOS — FÓRMULA ANIMAL
 * Lógica client-side: lê os dois XLS do Farma Fácil,
 * cruza os dados e gera o Excel de controlados para download.
 * Usa SheetJS (XLSX) que já está carregado no HTML.
 */

// ── Estado ───────────────────────────────────────────────────────────────
let dadosMov = null;   // dados extraídos do MOVIMENTO.XLS
let dadosCE  = null;   // dados extraídos do CLIENTE_END.XLS
let xlsxBlob = null;   // arquivo Excel gerado

// ── Referências DOM ───────────────────────────────────────────────────────
const zoneMov    = document.getElementById('zone-mov');
const zoneCE     = document.getElementById('zone-ce');
const fileMov    = document.getElementById('file-mov');
const fileCE     = document.getElementById('file-ce');
const fnameMov   = document.getElementById('fname-mov');
const fnameCE    = document.getElementById('fname-ce');
const btnGerar   = document.getElementById('btn-gerar');
const btnDownload= document.getElementById('btn-download');
const progWrap   = document.getElementById('progress-wrap');
const progBar    = document.getElementById('progress-bar');
const progText   = document.getElementById('progress-text');
const logBox     = document.getElementById('log-box');
const resultCard = document.getElementById('result-card');
const statsGrid  = document.getElementById('stats-grid');

// ── Helpers visuais ───────────────────────────────────────────────────────
function setProgress(pct, txt) {
  progWrap.classList.add('visible');
  progBar.style.width = pct + '%';
  progText.textContent = txt;
}

function log(msg, tipo = '') {
  logBox.classList.add('visible');
  const line = document.createElement('div');
  if (tipo) line.className = 'log-' + tipo;
  line.textContent = msg;
  logBox.appendChild(line);
  logBox.scrollTop = logBox.scrollHeight;
}

function checkReady() {
  btnGerar.disabled = !(dadosMov && dadosCE);
}

// ── Upload handlers ───────────────────────────────────────────────────────
function setupDrop(zone, input, fnameEl, tipo) {
  input.addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;
    readXLS(file, tipo, fnameEl, zone);
  });

  zone.addEventListener('dragover', e => {
    e.preventDefault();
    zone.classList.add('drag-over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (!file) return;
    readXLS(file, tipo, fnameEl, zone);
  });
}

function readXLS(file, tipo, fnameEl, zone) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', codepage: 1252 });
      const sh = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(sh, { header: 1, defval: '' });

      if (tipo === 'mov') {
        dadosMov = raw;
        fnameEl.textContent = '✓ ' + file.name;
        zone.classList.add('ready');
        log(`MOVIMENTO carregado: ${raw.length} linhas`, 'ok');
      } else {
        dadosCE = raw;
        fnameEl.textContent = '✓ ' + file.name;
        zone.classList.add('ready');
        log(`CLIENTE_END carregado: ${raw.length} linhas`, 'ok');
      }
      checkReady();
    } catch(err) {
      log('Erro ao ler arquivo: ' + err.message, 'err');
    }
  };
  reader.readAsArrayBuffer(file);
}

setupDrop(zoneMov, fileMov, fnameMov, 'mov');
setupDrop(zoneCE,  fileCE,  fnameCE,  'ce');

// ── Extração do MOVIMENTO ─────────────────────────────────────────────────
function extrairMovimento(raw) {
  const registros = [];
  let substancia = '', lista = '';

  function cell(row, c) {
    if (!row) return '';
    const v = row[c];
    return v !== undefined && v !== null ? String(v).trim() : '';
  }

  for (let r = 0; r < raw.length; r++) {
    const row = raw[r];

    // linha de cabeçalho de substância (col 6 contém 'Produto:')
    if (String(row[6] || '').includes('Produto:')) {
      substancia = cell(row, 8);
      lista      = cell(row, 3);
      continue;
    }

    // linha de dispensação (col 4 == 'O.M.')
    if (cell(row, 4) === 'O.M.') {
      const dataStr = cell(row, 0);
      let dt = null;
      try {
        const p = dataStr.split('/');
        if (p.length === 3) {
          const y = parseInt(p[2]) < 100 ? 2000 + parseInt(p[2]) : parseInt(p[2]);
          dt = new Date(y, parseInt(p[1]) - 1, parseInt(p[0]));
        }
      } catch(e) {}

      let qtdG = null;
      try { qtdG = parseFloat(String(row[17]).replace(',','.')); } catch(e) {}

      const crmvRaw = cell(row, 20);
      const crmvNr  = crmvRaw.replace(/CRMV\s+\w+:\s*/i, '').trim();

      function limpaNr(v) {
        const s = String(v).trim();
        const n = parseFloat(s);
        return !isNaN(n) && s === String(n) ? String(Math.round(n)) : s;
      }

      registros.push({
        substancia, lista,
        data: dt, dataStr,
        tutor:     cell(row, 7),
        nrOm:      limpaNr(row[11]),
        nrDoc:     limpaNr(row[12]),
        calculo:   cell(row, 15),
        qtdG, crmvRaw, crmvNr,
        nrReceita: limpaNr(row[25]),
      });
    }
  }
  return registros;
}

// ── Extração do CLIENTE_END ───────────────────────────────────────────────
function extrairClienteEnd(raw) {
  const dados = {};

  function cell(r, c) {
    if (r < 0 || r >= raw.length) return '';
    const v = raw[r][c];
    return v !== undefined && v !== null ? String(v).trim() : '';
  }

  for (let r = 0; r < raw.length; r++) {
    const status = cell(r, 12);
    if (status !== 'Ativa' && status !== 'Cancelada') continue;

    const nrRaw = cell(r, 36);
    let nr = nrRaw;
    const nrF = parseFloat(nrRaw);
    if (!isNaN(nrF)) nr = String(Math.round(nrF));

    const c1 = cell(r, 37);
    const c2 = cell(r + 1, 37);
    const cliente = (c1 + ' ' + c2).trim();

    const e1 = cell(r, 52);
    const e2 = cell(r + 1, 52);
    const endereco = (e1 + ' ' + e2).trim().replace(/(\d+)\.0\b/g, '$1');

    let prescritor = '', crmvNr = '', qtdeTexto = '', formula = '', doseMg = '';

    for (let off = 3; off < 10; off++) {
      if (r + off >= raw.length) break;
      if (cell(r + off, 0) === 'Prescritor:') {
        const pr = r + off;
        const p1 = cell(pr, 11);
        const p2 = cell(pr + 1, 11);
        prescritor = (p1 + ' ' + (cell(pr + 1, 0) === '' ? p2 : '')).trim();

        const cv = cell(pr, 43);
        const cvF = parseFloat(cv);
        crmvNr = !isNaN(cvF) ? String(Math.round(cvF)) : cv;

        qtdeTexto = cell(pr, 59);

        const fr = pr + 2;
        if (fr < raw.length) {
          formula = cell(fr, 43);
          const dr = cell(fr, 64);
          const drF = parseFloat(dr);
          doseMg = !isNaN(drF) ? drF : dr;
        }
        break;
      }
    }

    dados[nr] = { status, cliente, endereco, prescritor, crmvNr, qtdeTexto, formula, doseMg };
  }
  return dados;
}

// ── Cruzamento ────────────────────────────────────────────────────────────
function cruzar(movs, ced) {
  return movs.map(m => {
    const ce = ced[m.nrOm] || {};
    return {
      ...m,
      clienteFull: ce.cliente    || m.tutor,
      endereco:    ce.endereco   || '',
      prescritor:  ce.prescritor || '',
      crmvNrCE:    ce.crmvNr     || m.crmvNr,
      qtdeTexto:   ce.qtdeTexto  || '',
      doseMg:      ce.doseMg     || '',
      status:      ce.status     || 'Ativa',
    };
  });
}

// ── Geração do Excel ──────────────────────────────────────────────────────
function gerarExcel(dados, estInicialPadrao, nomeEstab) {
  const wb = XLSX.utils.book_new();

  // substâncias únicas mantendo ordem
  const substs = [...new Set(dados.map(d => d.substancia))];
  const paleta = ['4472C4','70AD47','ED7D31','FFC000','5B9BD5','A9D18E'];
  const scoreColor = {};
  substs.forEach((s, i) => scoreColor[s] = paleta[i % paleta.length]);

  const datas  = dados.filter(d => d.data).map(d => d.data);
  const periodo = datas.length
    ? `Período: ${fmtData(new Date(Math.min(...datas)))} a ${fmtData(new Date(Math.max(...datas)))}`
    : '';

  // helper: adiciona linha a uma worksheet (array de arrays)
  function makeWS(rows) {
    return XLSX.utils.aoa_to_sheet(rows);
  }

  // ── RESUMO ────────────────────────────────────────────────────────────
  const resumoRows = [
    [`RESUMO DE MOVIMENTAÇÃO — CONTROLADOS VETERINÁRIOS — ${nomeEstab}`],
    [periodo],
    [],
    ['Substância','Lista','Dispensações','Total (g)','Canceladas','Ativas'],
  ];
  for (const s of substs) {
    const ds  = dados.filter(d => d.substancia === s);
    const at  = ds.filter(d => d.status === 'Ativa').length;
    const ca  = ds.filter(d => d.status === 'Cancelada').length;
    const tg  = ds.filter(d => d.status === 'Ativa' && d.qtdG)
                  .reduce((acc, d) => acc + d.qtdG, 0);
    const li  = ds[0]?.lista || '';
    resumoRows.push([s, li, ds.length, Math.round(tg * 10000) / 10000, ca, at]);
  }
  const wsResumo = makeWS(resumoRows);
  wsResumo['!cols'] = [{wch:30},{wch:8},{wch:14},{wch:12},{wch:12},{wch:10}];
  XLSX.utils.book_append_sheet(wb, wsResumo, 'RESUMO');

  // ── CONTROLE ──────────────────────────────────────────────────────────
  const ctrlRows = [
    [`BASE DE DADOS — CONTROLADOS ${nomeEstab.toUpperCase()}`],
    [periodo],
    [],
    ['Nº OM','Nº DOC','Data','Tutor/Cliente','Endereço','CRMV nº',
     'Veterinário','Substância','Lista','Fórmula (Cálculo)',
     'Dose (mg)','Qtde Texto','Qtd (g)','Nº Receita','Status'],
    ...dados.map(d => [
      d.nrOm, d.nrDoc,
      d.data ? fmtData(d.data) : d.dataStr,
      d.clienteFull, d.endereco,
      d.crmvNrCE || d.crmvNr, d.prescritor,
      d.substancia, d.lista, d.calculo,
      d.doseMg, d.qtdeTexto, d.qtdG,
      d.nrReceita, d.status,
    ])
  ];
  const wsCtrl = makeWS(ctrlRows);
  wsCtrl['!cols'] = [
    {wch:9},{wch:9},{wch:12},{wch:30},{wch:40},{wch:10},
    {wch:24},{wch:22},{wch:6},{wch:18},{wch:10},{wch:16},{wch:9},{wch:14},{wch:10}
  ];
  XLSX.utils.book_append_sheet(wb, wsCtrl, 'CONTROLE');

  // ── CORPO por substância ──────────────────────────────────────────────
  for (const s of substs) {
    const ds   = dados.filter(d => d.substancia === s && d.status === 'Ativa');
    const li   = ds[0]?.lista || '';
    const safe = s.replace(/[^\w]/g, '_').substring(0, 25);

    const rows = [
      ['LIVRO DE REGISTRO DE ESTOQUE DE SUBSTÂNCIAS SUJEITAS A CONTROLE ESPECIAL DE USO VETERINÁRIO'],
      [`SUBSTÂNCIA (DCB): ${s}   |   Lista: ${li}   |   ${nomeEstab}`],
      [],
      ['DIA','MÊS','ANO','EST. INICIAL (g)','ENTRADA (g)','SAÍDA (g)',
       'PERDAS (g)','EST. FINAL (g)','REG / NR DOC','OUTRAS INFORMAÇÕES'],
      // linha 5: estoque inicial — preencher
      ['','','ESTOQUE INICIAL →', estInicialPadrao, '', '', '', estInicialPadrao,
       '','▶ Preencha o estoque inicial na coluna D desta linha'],
    ];

    let estoqueCalc = estInicialPadrao;
    for (const d of ds) {
      const dt = d.data;
      const saida = d.qtdG || 0;
      const novoEst = Math.round((estoqueCalc - saida) * 10000) / 10000;
      const outras = `Receita: ${d.nrReceita} | ${d.crmvRaw} | ${d.prescritor} | ${d.calculo}`;
      rows.push([
        dt ? dt.getDate()     : '',
        dt ? dt.getMonth() + 1: '',
        dt ? dt.getFullYear() : '',
        estoqueCalc,    // est. inicial
        '',             // entrada
        saida,          // saída
        '',             // perdas
        novoEst,        // est. final
        `${d.nrOm} / ${d.nrDoc}`,
        outras,
      ]);
      estoqueCalc = novoEst;
    }

    const ws2 = makeWS(rows);
    ws2['!cols'] = [
      {wch:6},{wch:6},{wch:6},{wch:14},{wch:11},
      {wch:11},{wch:9},{wch:14},{wch:20},{wch:55}
    ];
    XLSX.utils.book_append_sheet(wb, ws2, `CORPO_${safe}`);
  }

  // ── FICHAS_IMPRIMIR ───────────────────────────────────────────────────
  const fichasRows = [
    ['FICHAS INDIVIDUAIS DE DISPENSAÇÃO — IMPRIMIR E COLAR NO LIVRO FÍSICO'],
    ['Nº OM','Nº DOC','Data','Tutor','Endereço','Veterinário',
     'CRMV','Substância','Fórmula','Qtd (g)','Nº Receita','Status'],
    ...dados.map(d => [
      d.nrOm, d.nrDoc,
      d.data ? fmtData(d.data) : d.dataStr,
      d.clienteFull, d.endereco, d.prescritor,
      d.crmvNrCE || d.crmvNr,
      d.substancia, d.calculo, d.qtdG, d.nrReceita, d.status,
    ])
  ];
  const wsFichas = makeWS(fichasRows);
  wsFichas['!cols'] = [
    {wch:9},{wch:9},{wch:12},{wch:28},{wch:38},
    {wch:22},{wch:10},{wch:20},{wch:16},{wch:9},{wch:14},{wch:10}
  ];
  XLSX.utils.book_append_sheet(wb, wsFichas, 'FICHAS_IMPRIMIR');

  // Gerar o arquivo
  const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Blob([wbOut], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

function fmtData(d) {
  if (!d) return '';
  const dd = String(d.getDate()).padStart(2,'0');
  const mm = String(d.getMonth()+1).padStart(2,'0');
  return `${dd}/${mm}/${d.getFullYear()}`;
}

// ── Botão Gerar ───────────────────────────────────────────────────────────
btnGerar.addEventListener('click', async () => {
  logBox.innerHTML = '';
  resultCard.classList.remove('visible');
  xlsxBlob = null;

  const estInicial  = parseFloat(document.getElementById('est-inicial').value) || 0;
  const nomeEstab   = document.getElementById('estabelecimento').value.trim() || 'Fórmula Animal';

  // Spinner no botão
  btnGerar.disabled = true;
  btnGerar.innerHTML = `<div class="spinner"></div> Processando...`;

  // Pequeno delay para o browser pintar
  await new Promise(r => setTimeout(r, 50));

  try {
    setProgress(10, 'Lendo MOVIMENTO.XLS...');
    log('Iniciando extração do Movimento...', 'ok');
    const movs = extrairMovimento(dadosMov);
    log(`  → ${movs.length} dispensações encontradas`);

    setProgress(35, 'Lendo CLIENTE_END.XLS...');
    log('Iniciando extração do Receituário...', 'ok');
    const ced = extrairClienteEnd(dadosCE);
    log(`  → ${Object.keys(ced).length} registros de receituário encontrados`);

    setProgress(55, 'Cruzando dados...');
    log('Cruzando Movimento × Receituário pelo Nº OM...', 'ok');
    const dados = cruzar(movs, ced);

    // verificar cruzamentos sem match
    const semMatch = dados.filter(d => !d.endereco && !d.prescritor).length;
    if (semMatch > 0)
      log(`  ⚠ ${semMatch} registros sem correspondência no Receituário — verifique os períodos`, 'warn');

    const substs  = [...new Set(dados.map(d => d.substancia))];
    const ativas  = dados.filter(d => d.status === 'Ativa').length;
    const totalG  = dados
      .filter(d => d.status === 'Ativa' && d.qtdG)
      .reduce((a, d) => a + d.qtdG, 0);

    setProgress(75, 'Gerando Excel...');
    log(`Gerando planilha: ${substs.join(', ')}...`, 'ok');
    xlsxBlob = gerarExcel(dados, estInicial, nomeEstab);

    setProgress(100, 'Concluído!');
    log('Planilha gerada com sucesso!', 'ok');

    // Mostrar estatísticas
    statsGrid.innerHTML = '';
    const stats = [
      { num: dados.length,               lbl: 'Dispensações' },
      { num: ativas,                     lbl: 'Ativas' },
      { num: substs.length,              lbl: 'Substâncias' },
      { num: totalG.toFixed(4) + ' g',   lbl: 'Total saída' },
    ];
    stats.forEach(s => {
      statsGrid.innerHTML += `
        <div class="stat-box">
          <div class="stat-num">${s.num}</div>
          <div class="stat-lbl">${s.lbl}</div>
        </div>`;
    });

    resultCard.classList.add('visible');

  } catch(err) {
    log('ERRO: ' + err.message, 'err');
    console.error(err);
    setProgress(0, 'Erro no processamento.');
  } finally {
    btnGerar.disabled = false;
    btnGerar.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round">
        <path d="M12 2L2 7l10 5 10-5-10-5z"/>
        <path d="M2 17l10 5 10-5"/>
        <path d="M2 12l10 5 10-5"/>
      </svg>
      Gerar Planilha de Controlados`;
    checkReady();
  }
});

// ── Botão Download ────────────────────────────────────────────────────────
btnDownload.addEventListener('click', () => {
  if (!xlsxBlob) return;
  const url = URL.createObjectURL(xlsxBlob);
  const a   = document.createElement('a');
  a.href    = url;
  a.download = 'Controlados_FormulaAnimal.xlsx';
  a.click();
  URL.revokeObjectURL(url);
});
