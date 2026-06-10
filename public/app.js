/**
 * CONTROLADOS v2.0 — FÓRMULA ANIMAL
 * ─────────────────────────────────────────────────────
 * Novidades v2:
 *  • Estoque inicial individual por substância
 *  • Histórico persistente (localStorage)
 *  • Estoque final → inicial automático no próximo período
 *  • Backup/restauração em JSON
 */

// ── Substâncias controladas cadastradas ──────────────────────────────────
const SUBSTANCIAS = [
  { nome: 'Gabapentina',  lista: 'C1', dcb: '04369' },
  { nome: 'Fluoxetina',   lista: 'C1', dcb: '03094' },
  { nome: 'Amitriptilina',lista: 'C1', dcb: '00423' },
  { nome: 'Selegilina',   lista: 'C1', dcb: '07929' },
  { nome: 'Tramadol',     lista: 'A2', dcb: '08806' },
  { nome: 'Codeína',      lista: 'A2', dcb: '01706' },
  { nome: 'Ribavirina',   lista: 'C1', dcb: '07168' },
];

// ── Chave do localStorage ─────────────────────────────────────────────────
const LS_KEY = 'controlados_fa_v2';

// ── Estado global ──────────────────────────────────────────────────────────
let dadosMov  = null;
let dadosCE   = null;
let xlsxBlob  = null;

// ── Helpers de localStorage ───────────────────────────────────────────────
function loadHistorico() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch(e) { return []; }
}

function saveHistorico(hist) {
  localStorage.setItem(LS_KEY, JSON.stringify(hist));
}

/** Retorna o último estoque final registrado para uma substância */
function ultimoEstoqueFinal(nomeSubst) {
  const hist = loadHistorico();
  if (!hist.length) return 0;
  // percorre do mais recente para o mais antigo
  for (let i = hist.length - 1; i >= 0; i--) {
    const reg = hist[i];
    if (reg.estoquesFinal && reg.estoquesFinal[nomeSubst] !== undefined) {
      return reg.estoquesFinal[nomeSubst];
    }
  }
  return 0;
}

// ── Montar grid de estoques iniciais ─────────────────────────────────────
function montarEstGrid() {
  const grid = document.getElementById('est-grid');
  grid.innerHTML = '';
  SUBSTANCIAS.forEach(s => {
    const ultimo = ultimoEstoqueFinal(s.nome);
    const div = document.createElement('div');
    div.className = 'est-field';
    div.innerHTML = `
      <div class="subst-name">${s.nome}</div>
      <label>Lista ${s.lista} · DCB ${s.dcb}</label>
      <input type="number" id="est-${s.nome}" value="${ultimo}" step="0.0001" min="0" />
      <div class="last-val">${ultimo > 0 ? '↑ do período anterior: ' + ultimo.toFixed(4) + ' g' : 'Sem histórico anterior'}</div>
    `;
    grid.appendChild(div);
  });
}

function getEstoqueInicial() {
  const est = {};
  SUBSTANCIAS.forEach(s => {
    const input = document.getElementById('est-' + s.nome);
    est[s.nome] = input ? (parseFloat(input.value) || 0) : 0;
  });
  return est;
}

// ── Tabs ──────────────────────────────────────────────────────────────────
function switchTab(id) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + id).classList.add('active');
  event.target.classList.add('active');
  if (id === 'historico') renderHistorico();
}

// ── Upload handlers ───────────────────────────────────────────────────────
function setupDrop(zoneId, inputId, fnameId, tipo) {
  const zone  = document.getElementById(zoneId);
  const input = document.getElementById(inputId);
  const fnEl  = document.getElementById(fnameId);

  input.addEventListener('change', e => {
    if (e.target.files[0]) readXLS(e.target.files[0], tipo, fnEl, zone);
  });
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) readXLS(e.dataTransfer.files[0], tipo, fnEl, zone);
  });
}

function readXLS(file, tipo, fnEl, zone) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb  = XLSX.read(new Uint8Array(e.target.result), { type: 'array', codepage: 1252 });
      const sh  = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(sh, { header: 1, defval: '' });
      if (tipo === 'mov') {
        dadosMov = raw;
        fnEl.textContent = '✓ ' + file.name;
        zone.classList.add('ready');
        log(`MOVIMENTO carregado — ${raw.length} linhas`, 'ok');
      } else {
        dadosCE = raw;
        fnEl.textContent = '✓ ' + file.name;
        zone.classList.add('ready');
        log(`CLIENTE_END carregado — ${raw.length} linhas`, 'ok');
      }
      checkReady();
    } catch(err) { log('Erro ao ler arquivo: ' + err.message, 'err'); }
  };
  reader.readAsArrayBuffer(file);
}

// ── Extração MOVIMENTO ────────────────────────────────────────────────────
function extrairMovimento(raw) {
  const recs = [];
  let subst = '', lista = '';
  function cell(row, c) {
    if (!row) return '';
    const v = row[c];
    return (v !== undefined && v !== null) ? String(v).trim() : '';
  }
  function limpaNr(v) {
    const n = parseFloat(String(v));
    return (!isNaN(n) && String(v).trim() === String(n)) ? String(Math.round(n)) : String(v).trim();
  }
  for (let r = 0; r < raw.length; r++) {
    const row = raw[r];
    if (String(row[6] || '').includes('Produto:')) {
      subst = cell(row, 8);
      lista = cell(row, 3);
      continue;
    }
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
      try { qtdG = parseFloat(String(row[17]).replace(',', '.')); } catch(e) {}
      const crmvRaw = cell(row, 20);
      const crmvNr  = crmvRaw.replace(/CRMV\s+\w+:\s*/i, '').trim();
      recs.push({
        substancia: subst, lista, data: dt, dataStr,
        tutor: cell(row, 7),
        nrOm: limpaNr(row[11]), nrDoc: limpaNr(row[12]),
        calculo: cell(row, 15), qtdG, crmvRaw, crmvNr,
        nrReceita: limpaNr(row[25]),
      });
    }
  }
  return recs;
}

// ── Extração CLIENTE_END ──────────────────────────────────────────────────
function extrairClienteEnd(raw) {
  const dados = {};
  function cell(r, c) {
    if (r < 0 || r >= raw.length) return '';
    const v = raw[r][c];
    return (v !== undefined && v !== null) ? String(v).trim() : '';
  }
  for (let r = 0; r < raw.length; r++) {
    const status = cell(r, 12);
    if (status !== 'Ativa' && status !== 'Cancelada') continue;
    const nrRaw = cell(r, 36);
    let nr = nrRaw;
    const nrF = parseFloat(nrRaw);
    if (!isNaN(nrF)) nr = String(Math.round(nrF));
    const cliente  = (cell(r, 37) + ' ' + cell(r+1, 37)).trim();
    const end_l1   = cell(r, 52);
    const end_l2   = cell(r+1, 52);
    const endereco = (end_l1 + ' ' + end_l2).trim().replace(/(\d+)\.0\b/g, '$1');
    let prescritor = '', crmvNr = '', qtdeTexto = '', formula = '', doseMg = '';
    for (let off = 3; off < 10; off++) {
      if (r + off >= raw.length) break;
      if (cell(r + off, 0) === 'Prescritor:') {
        const pr = r + off;
        const p1 = cell(pr, 11);
        const p2 = (cell(pr+1, 0) === '') ? cell(pr+1, 11) : '';
        prescritor = (p1 + ' ' + p2).trim();
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
    return { ...m,
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

// ── Geração Excel ─────────────────────────────────────────────────────────
function gerarExcel(dados, estInicial, nomeEstab, periodoLabel) {
  const wb = XLSX.utils.book_new();
  const substs = [...new Set(dados.map(d => d.substancia))];

  const datas   = dados.filter(d => d.data).map(d => d.data);
  const periodo = datas.length
    ? `${fmtData(new Date(Math.min(...datas)))} a ${fmtData(new Date(Math.max(...datas)))}`
    : '';
  const titulo  = periodoLabel || periodo;

  const estoquesFinal = {};

  // ── RESUMO ──
  const resumoRows = [
    [`RELATÓRIO DE MOVIMENTAÇÃO — CONTROLADOS VETERINÁRIOS — ${nomeEstab}`],
    [`Período: ${titulo}`],
    [],
    ['Substância','Lista','DCB','Est. Inicial (g)','Dispensações','Total Saída (g)','Est. Final (g)'],
  ];
  for (const s of SUBSTANCIAS) {
    // encontra dados desta substância pelo nome ou pelo DCB
    const ds = dados.filter(d =>
      d.substancia.toUpperCase().includes(s.nome.toUpperCase()) ||
      d.substancia.includes(s.dcb)
    );
    const ativas = ds.filter(d => d.status === 'Ativa');
    const totalSaida = arred(ativas.reduce((a, d) => a + (d.qtdG || 0), 0));
    const estIni  = estInicial[s.nome] || 0;
    const estFin  = arred(estIni - totalSaida);
    estoquesFinal[s.nome] = estFin;
    resumoRows.push([
      s.nome, s.lista, s.dcb,
      estIni, ativas.length, totalSaida, estFin
    ]);
  }
  const wsRes = XLSX.utils.aoa_to_sheet(resumoRows);
  wsRes['!cols'] = [{wch:20},{wch:7},{wch:8},{wch:15},{wch:14},{wch:16},{wch:14}];
  XLSX.utils.book_append_sheet(wb, wsRes, 'RESUMO');

  // ── CONTROLE (banco geral) ──
  const ctrlRows = [
    [`BASE DE DADOS — ${nomeEstab.toUpperCase()} — ${titulo}`],
    [],
    ['Nº OM','Nº DOC','Data','Tutor/Cliente','Endereço','CRMV nº',
     'Veterinário','Substância','Lista','Fórmula','Dose (mg)',
     'Qtde Texto','Qtd (g)','Nº Receita','Status'],
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
  const wsCtrl = XLSX.utils.aoa_to_sheet(ctrlRows);
  wsCtrl['!cols'] = [{wch:9},{wch:9},{wch:12},{wch:30},{wch:40},{wch:10},{wch:24},
                     {wch:20},{wch:6},{wch:18},{wch:10},{wch:16},{wch:9},{wch:14},{wch:10}];
  XLSX.utils.book_append_sheet(wb, wsCtrl, 'CONTROLE');

  // ── CORPO por substância ──
  for (const s of SUBSTANCIAS) {
    const ds = dados.filter(d =>
      (d.substancia.toUpperCase().includes(s.nome.toUpperCase()) ||
       d.substancia.includes(s.dcb)) && d.status === 'Ativa'
    );

    const safe = s.nome.replace(/[^\w]/g, '_');
    const estIni = estInicial[s.nome] || 0;

    const rows = [
      ['LIVRO DE REGISTRO DE ESTOQUE DE SUBSTÂNCIAS SUJEITAS A CONTROLE ESPECIAL DE USO VETERINÁRIO'],
      [`SUBSTÂNCIA (DCB): ${s.nome}   |   Lista: ${s.lista}   |   ${nomeEstab}`],
      [`Período: ${titulo}`],
      [],
      ['DIA','MÊS','ANO','EST. INICIAL (g)','ENTRADA (g)','SAÍDA (g)',
       'PERDAS (g)','EST. FINAL (g)','REG / NR DOC','OUTRAS INFORMAÇÕES'],
      // linha de estoque inicial
      ['', '', 'ESTOQUE INICIAL', estIni, '', '', '', estIni,
       '', `Estoque inicial do período — ${titulo}`],
    ];

    let saldo = estIni;
    for (const d of ds) {
      const dt    = d.data;
      const saida = d.qtdG || 0;
      const entrada = 0;
      const perda   = 0;
      const novoSaldo = arred(saldo + entrada - saida - perda);
      const outras = [
        `Receita: ${d.nrReceita}`,
        d.crmvRaw,
        d.prescritor,
        d.calculo,
      ].filter(Boolean).join(' | ');

      rows.push([
        dt ? dt.getDate()      : '',
        dt ? dt.getMonth() + 1 : '',
        dt ? dt.getFullYear()  : '',
        arred(saldo),
        entrada || '',
        saida   || '',
        perda   || '',
        novoSaldo,
        `${d.nrOm} / ${d.nrDoc}`,
        outras,
      ]);
      saldo = novoSaldo;
    }

    // linha de estoque final
    rows.push(['', '', 'ESTOQUE FINAL', '', '', '', '', arred(saldo),
               '', `Estoque final do período — transferir para próximo período`]);

    const ws2 = XLSX.utils.aoa_to_sheet(rows);
    ws2['!cols'] = [{wch:5},{wch:5},{wch:14},{wch:14},{wch:11},{wch:11},
                    {wch:9},{wch:14},{wch:20},{wch:55}];
    XLSX.utils.book_append_sheet(wb, ws2, `CORPO_${safe}`);
  }

  // ── FICHAS ──
  const fichasRows = [
    [`FICHAS DE DISPENSAÇÃO — ${nomeEstab} — ${titulo}`],
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
  const wsFichas = XLSX.utils.aoa_to_sheet(fichasRows);
  wsFichas['!cols'] = [{wch:9},{wch:9},{wch:12},{wch:28},{wch:38},
                       {wch:22},{wch:10},{wch:20},{wch:16},{wch:9},{wch:14},{wch:10}];
  XLSX.utils.book_append_sheet(wb, wsFichas, 'FICHAS_IMPRIMIR');

  const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return {
    blob: new Blob([wbOut], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    estoquesFinal,
  };
}

// ── Salvar no histórico ────────────────────────────────────────────────────
function salvarNoHistorico(dados, estInicial, estoquesFinal, periodoLabel, nomeEstab) {
  const hist = loadHistorico();
  const datas = dados.filter(d => d.data).map(d => d.data);
  const registro = {
    id:           Date.now(),
    geradoEm:     new Date().toISOString(),
    periodoLabel: periodoLabel,
    estabelecimento: nomeEstab,
    dataInicio:   datas.length ? new Date(Math.min(...datas)).toISOString() : null,
    dataFim:      datas.length ? new Date(Math.max(...datas)).toISOString() : null,
    totalRegistros: dados.length,
    substanciasAtivas: [...new Set(dados.filter(d => d.status === 'Ativa').map(d => d.substancia))],
    estoquesInicial: estInicial,
    estoquesFinal:   estoquesFinal,
  };
  hist.push(registro);
  saveHistorico(hist);
}

// ── Renderizar histórico ──────────────────────────────────────────────────
function renderHistorico() {
  const hist     = loadHistorico();
  const lista    = document.getElementById('hist-lista');
  if (!hist.length) {
    lista.innerHTML = `<div class="hist-empty">Nenhum registro gerado ainda.<br>Gere sua primeira planilha para começar o histórico.</div>`;
    return;
  }
  lista.innerHTML = '';
  // do mais recente para o mais antigo
  [...hist].reverse().forEach(reg => {
    const dt = new Date(reg.geradoEm);
    const dtStr = dt.toLocaleDateString('pt-BR') + ' ' + dt.toLocaleTimeString('pt-BR', {hour:'2-digit',minute:'2-digit'});
    const div = document.createElement('div');
    div.className = 'hist-entry';
    // pills de substâncias com estoque final
    const pillsHtml = SUBSTANCIAS.map(s => {
      const fin = reg.estoquesFinal?.[s.nome];
      const ini = reg.estoquesInicial?.[s.nome];
      if (fin === undefined && ini === undefined) return '';
      return `<span class="hist-pill">${s.nome} <span>${fin !== undefined ? fin.toFixed(4) + 'g' : '—'}</span></span>`;
    }).join('');

    div.innerHTML = `
      <div class="hist-header">
        <div class="hist-periodo">${reg.periodoLabel || 'Período ' + (reg.dataInicio ? fmtData(new Date(reg.dataInicio)) : '?')}</div>
        <div class="hist-date">${dtStr}</div>
      </div>
      <div style="font-family:var(--mono);font-size:.72rem;color:var(--muted);margin-bottom:8px">
        ${reg.totalRegistros} dispensações · ${reg.estabelecimento}
      </div>
      <div style="font-family:var(--mono);font-size:.68rem;color:var(--muted);margin-bottom:8px">
        Estoques finais:
      </div>
      <div class="hist-substs">${pillsHtml}</div>
      <div class="hist-actions">
        <button class="btn-secondary" style="font-size:.75rem;padding:7px 14px"
          onclick="usarComoInicial(${reg.id})">
          ↑ Usar como estoque inicial
        </button>
        <button class="btn-danger" onclick="excluirRegistro(${reg.id})">Excluir</button>
      </div>`;
    lista.appendChild(div);
  });
}

function usarComoInicial(id) {
  const hist = loadHistorico();
  const reg  = hist.find(r => r.id === id);
  if (!reg || !reg.estoquesFinal) return;
  SUBSTANCIAS.forEach(s => {
    const input = document.getElementById('est-' + s.nome);
    if (input && reg.estoquesFinal[s.nome] !== undefined) {
      input.value = reg.estoquesFinal[s.nome];
    }
  });
  switchTab('gerar');
  // Simula click na aba Gerar
  document.querySelectorAll('.tab')[0].classList.add('active');
  document.querySelectorAll('.tab')[1].classList.remove('active');
  log('Estoques iniciais carregados do registro selecionado.', 'ok');
}

function excluirRegistro(id) {
  if (!confirm('Excluir este registro do histórico?')) return;
  const hist = loadHistorico().filter(r => r.id !== id);
  saveHistorico(hist);
  renderHistorico();
}

// ── Backup / Restauração ──────────────────────────────────────────────────
function exportarBackup() {
  const hist = loadHistorico();
  const json = JSON.stringify({ versao: 2, exportadoEm: new Date().toISOString(), historico: hist }, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = `backup_controlados_${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function importarBackup(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = JSON.parse(e.target.result);
      const hist = data.historico || data; // compatibilidade
      if (!Array.isArray(hist)) throw new Error('Formato inválido');
      if (!confirm(`Importar ${hist.length} registros? O histórico atual será substituído.`)) return;
      saveHistorico(hist);
      renderHistorico();
      montarEstGrid();
      alert('Backup importado com sucesso!');
    } catch(err) {
      alert('Erro ao importar: ' + err.message);
    }
  };
  reader.readAsText(file);
  input.value = '';
}

// ── Utils ─────────────────────────────────────────────────────────────────
function fmtData(d) {
  if (!d) return '';
  return String(d.getDate()).padStart(2,'0') + '/' +
         String(d.getMonth()+1).padStart(2,'0') + '/' +
         d.getFullYear();
}

function arred(n) { return Math.round((n || 0) * 10000) / 10000; }

function setProgress(pct, txt) {
  document.getElementById('progress-wrap').classList.add('visible');
  document.getElementById('progress-bar').style.width = pct + '%';
  document.getElementById('progress-text').textContent = txt;
}

function log(msg, tipo) {
  const box = document.getElementById('log-box');
  box.classList.add('visible');
  const line = document.createElement('div');
  if (tipo) line.className = 'log-' + tipo;
  line.textContent = msg;
  box.appendChild(line);
  box.scrollTop = box.scrollHeight;
}

function checkReady() {
  document.getElementById('btn-gerar').disabled = !(dadosMov && dadosCE);
}

// ── Botão Gerar ───────────────────────────────────────────────────────────
document.getElementById('btn-gerar').addEventListener('click', async () => {
  document.getElementById('log-box').innerHTML = '';
  document.getElementById('result-card').classList.remove('visible');
  xlsxBlob = null;

  const nomeEstab    = document.getElementById('estabelecimento').value.trim() || 'Fórmula Animal';
  const periodoLabel = document.getElementById('periodo-label').value.trim();
  const estInicial   = getEstoqueInicial();

  const btn = document.getElementById('btn-gerar');
  btn.disabled = true;
  btn.innerHTML = `<div class="spinner"></div> Processando...`;

  await new Promise(r => setTimeout(r, 50));

  try {
    setProgress(10, 'Lendo MOVIMENTO.XLS...');
    const movs = extrairMovimento(dadosMov);
    log(`Movimento: ${movs.length} dispensações encontradas`, 'ok');

    setProgress(35, 'Lendo CLIENTE_END.XLS...');
    const ced = extrairClienteEnd(dadosCE);
    log(`Receituário: ${Object.keys(ced).length} registros encontrados`, 'ok');

    setProgress(55, 'Cruzando dados...');
    const dados = cruzar(movs, ced);

    const semMatch = dados.filter(d => !d.prescritor).length;
    if (semMatch > 0) log(`⚠ ${semMatch} registros sem correspondência no Receituário`, 'warn');

    setProgress(75, 'Gerando Excel...');
    const { blob, estoquesFinal } = gerarExcel(dados, estInicial, nomeEstab, periodoLabel);
    xlsxBlob = blob;

    setProgress(90, 'Salvando histórico...');
    salvarNoHistorico(dados, estInicial, estoquesFinal, periodoLabel, nomeEstab);
    montarEstGrid(); // atualiza os campos com os novos valores finais

    setProgress(100, 'Concluído!');
    log('Planilha gerada e histórico atualizado!', 'ok');

    // Estatísticas
    const ativas   = dados.filter(d => d.status === 'Ativa');
    const substs   = [...new Set(dados.map(d => d.substancia))];
    const totalG   = ativas.reduce((a, d) => a + (d.qtdG || 0), 0);
    const statsGrid = document.getElementById('stats-grid');
    statsGrid.innerHTML = '';
    [
      { num: dados.length,            lbl: 'Dispensações' },
      { num: ativas.length,           lbl: 'Ativas' },
      { num: substs.length,           lbl: 'Substâncias' },
      { num: arred(totalG) + ' g',    lbl: 'Total saída' },
    ].forEach(s => {
      statsGrid.innerHTML += `<div class="stat-box"><div class="stat-num">${s.num}</div><div class="stat-lbl">${s.lbl}</div></div>`;
    });

    // Mostrar estoques finais calculados
    log('── Estoques finais calculados ──', 'ok');
    SUBSTANCIAS.forEach(s => {
      const fin = estoquesFinal[s.nome];
      if (fin !== undefined) log(`  ${s.nome}: ${fin.toFixed(4)} g`, 'ok');
    });

    document.getElementById('result-card').classList.add('visible');

  } catch(err) {
    log('ERRO: ' + err.message, 'err');
    console.error(err);
    setProgress(0, 'Erro no processamento.');
  } finally {
    btn.disabled = false;
    btn.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round">
        <path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/>
      </svg>
      Gerar Planilha de Controlados`;
    checkReady();
  }
});

// ── Botão Download ────────────────────────────────────────────────────────
document.getElementById('btn-download').addEventListener('click', () => {
  if (!xlsxBlob) return;
  const url = URL.createObjectURL(xlsxBlob);
  const a   = document.createElement('a');
  a.href    = url;
  a.download = 'Controlados_FormulaAnimal.xlsx';
  a.click();
  URL.revokeObjectURL(url);
});

// ── Init ──────────────────────────────────────────────────────────────────
setupDrop('zone-mov', 'file-mov', 'fname-mov', 'mov');
setupDrop('zone-ce',  'file-ce',  'fname-ce',  'ce');
montarEstGrid();
