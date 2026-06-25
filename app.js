// ═══════════════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════════════
let DB = {
  transacoes: [],
  projetos: [],
  orcamento: [],
  clientes: [],
  loaded: false,
  filename: ''
};

let donutChart = null;

// ═══════════════════════════════════════════════════════
// UTILS
// ═══════════════════════════════════════════════════════
const fmt = (n) => {
  if (Math.abs(n) >= 1000000) return 'R$' + (n/1000000).toFixed(1) + 'M';
  if (Math.abs(n) >= 1000) return 'R$' + (n/1000).toFixed(1) + 'k';
  return 'R$' + Math.round(n).toLocaleString('pt-BR');
};
const fmtFull = (n) => 'R$\u00a0' + Math.abs(n).toLocaleString('pt-BR', {minimumFractionDigits:0, maximumFractionDigits:0});
const pct = (n) => (n*100).toFixed(1) + '%';

const MESES_ORDER = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

function groupBy(arr, key) {
  return arr.reduce((acc, row) => {
    const k = row[key] || 'Outros';
    acc[k] = (acc[k] || []);
    acc[k].push(row);
    return acc;
  }, {});
}

function sumBy(arr, key) {
  return arr.reduce((s, r) => s + (parseFloat(r[key]) || 0), 0);
}

// ═══════════════════════════════════════════════════════
// FILE IMPORT
// ═══════════════════════════════════════════════════════
document.getElementById('fileInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('loadingOverlay').classList.add('show');
  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
      parseWorkbook(wb, file.name);
    } catch(err) {
      showBanner('Erro ao ler o arquivo: ' + err.message, true);
    }
    document.getElementById('loadingOverlay').classList.remove('show');
    e.target.value = '';
  };
  reader.readAsArrayBuffer(file);
});

function parseWorkbook(wb, filename) {
  const sheetMap = {};
  wb.SheetNames.forEach(name => { sheetMap[name.toLowerCase()] = name; });

  // Helper: sheet to JSON
  const toJson = (sheetName) => {
    const key = Object.keys(sheetMap).find(k => k.includes(sheetName.toLowerCase()));
    if (!key) return [];
    return XLSX.utils.sheet_to_json(wb.Sheets[sheetMap[key]], { defval: '' });
  };

  const trans = toJson('transac');
  const proj  = toJson('projeto');
  const orca  = toJson('or');
  const cli   = toJson('cliente');

  if (!trans.length && !proj.length && !orca.length) {
    showBanner('Arquivo inválido: nenhuma aba reconhecida. Use o modelo fornecido.', true);
    return;
  }

  DB.transacoes = trans.map(r => ({
    Data:        r['Data'] || r['data'] || '',
    Descricao:   r['Descricao'] || r['Descrição'] || r['descricao'] || '',
    Categoria:   r['Categoria'] || r['categoria'] || 'Outros',
    Tipo_Cliente:r['Tipo_Cliente'] || r['Tipo Cliente'] || '—',
    Tipo:        r['Tipo'] || r['tipo'] || '',
    EntradaOuSaida:         r['E/S'] || r['E/S'] || r['e/s'] || '',
    Valor:       parseFloat(r['Valor'] || r['valor'] || 0),
    Mes:         r['Mes'] || r['Mês'] || r['mes'] || ''
  }));

  DB.projetos = proj.map(r => {
    const rec = parseFloat(r['Receita'] || r['receita'] || 0);
    const cus = parseFloat(r['Custo'] || r['custo'] || 0);
    return {
      Projeto:   r['Projeto'] || r['projeto'] || '',
      Cliente:   r['Cliente'] || r['cliente'] || '',
      Receita:   rec,
      Custo:     cus,
      Lucro:     parseFloat(r['Lucro'] || r['lucro'] || (rec - cus)),
      Margem:    rec > 0 ? (rec - cus) / rec : 0,
      Status:    r['Status'] || r['status'] || '',
      Mes_Inicio:r['Mes_Inicio'] || r['Mês Inicio'] || ''
    };
  });

  DB.orcamento = orca.map(r => {
    const orc = parseFloat(r['Orcado'] || r['Orçado'] || r['orcado'] || 0);
    const rea = parseFloat(r['Realizado'] || r['realizado'] || 0);
    return {
      Categoria: r['Categoria'] || r['categoria'] || '',
      Orcado:    orc,
      Realizado: rea,
      Variacao:  rea - orc,
      Status:    r['Status'] || r['status'] || (rea <= orc ? 'OK' : rea > orc*1.1 ? 'Excedeu' : 'Atenção')
    };
  });

  DB.clientes = cli.map(r => ({
    Cliente:        r['Cliente'] || r['cliente'] || '',
    Tipo:           r['Tipo'] || r['tipo'] || '',
    Canal_Aquisicao:r['Canal_Aquisicao'] || r['Canal Aquisicao'] || r['Canal_Aquisição'] || 'Outros',
    Receita_Total:  parseFloat(r['Receita_Total'] || r['Receita Total'] || 0),
    Projetos_Qtd:   parseInt(r['Projetos_Qtd'] || r['Projetos Qtd'] || 1)
  }));

  DB.loaded = true;
  DB.filename = filename;

  renderAll();
  showBanner(`✓ "${filename}" carregado com sucesso — ${DB.transacoes.length} transações, ${DB.projetos.length} projetos.`);
}

// ═══════════════════════════════════════════════════════
// LOGIN (Google Identity Services) + LEITURA VIA SHEETS API
// ═══════════════════════════════════════════════════════
// ID extraído da URL da planilha:
// https://docs.google.com/spreadsheets/d/ESTE_ID_AQUI/edit
const GS_SHEET_ID = '1n0di7NhhpIjLCVB3NsOl4i9UQu1ZjxCtAgDjD55-j8Y';

// Client ID OAuth2 gerado no Google Cloud Console (tipo "Aplicativo da Web")
const GOOGLE_CLIENT_ID = '1000356753447-6onugldk7163gtihrqdhve11hqs5es74.apps.googleusercontent.com';

// Domínio institucional — usado apenas para feedback visual rápido.
// A segurança REAL vem do compartilhamento da planilha no Google Drive.
const ALLOWED_DOMAIN = 'https://enetec-automacoes.github.io/';

// Nomes das abas exatamente como aparecem no Google Sheets
const GS_TABS = ['Transacoes', 'Projetos', 'Orçamento', 'Clientes'];

let accessToken = null;
let tokenClient = null;

function iniciarLogin() {
  document.getElementById('loginError').classList.remove('show');
  if (!tokenClient) {
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: 'https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/userinfo.email',
      hint: '', // opcional: pode sugerir @suainstituicao.com.br
      callback: handleTokenResponse
    });
  }
  tokenClient.requestAccessToken();
}

async function handleTokenResponse(resp) {
  if (resp.error) {
    console.error('[Login Google] erro:', resp);
    document.getElementById('loginError').textContent = 'Falha no login: ' + resp.error;
    document.getElementById('loginError').classList.add('show');
    return;
  }
  accessToken = resp.access_token;

  // Busca e-mail do usuário só para exibir no cabeçalho (cosmético)
  try {
    const r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: 'Bearer ' + accessToken }
    });
    const info = await r.json();
    const badge = document.getElementById('userBadge');
    badge.style.display = 'flex';
    badge.innerHTML = `<img src="${info.picture || ''}" onerror="this.style.display='none'">
      ${info.email} <span class="logout-link" onclick="logout()">sair</span>`;
  } catch (e) { /* não bloqueia o fluxo se falhar */ }

  document.getElementById('loginGate').classList.add('hidden');
  loadFromGoogleSheets();

  // Renova o token silenciosamente antes de expirar (~55 min)
  setTimeout(() => tokenClient.requestAccessToken({ prompt: '' }), 55 * 60 * 1000);
}

function logout() {
  if (accessToken) google.accounts.oauth2.revoke(accessToken, () => {});
  accessToken = null;
  document.getElementById('userBadge').style.display = 'none';
  document.getElementById('loginGate').classList.remove('hidden');
}

// Lê todas as abas de uma vez via Sheets API (values:batchGet) e monta
// um "workbook" no formato que parseWorkbook() já espera — assim a função
// que processa os dados continua exatamente igual.
async function loadFromGoogleSheets() {
  if (!accessToken) { document.getElementById('loginGate').classList.remove('hidden'); return; }

  const overlay = document.getElementById('loadingOverlay');
  overlay.classList.add('show');
  try {
    const params = GS_TABS.map(t => 'ranges=' + encodeURIComponent(t)).join('&');
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${GS_SHEET_ID}/values:batchGet?${params}`;
    const resp = await fetch(url, { headers: { Authorization: 'Bearer ' + accessToken } });

    if (resp.status === 401) {
      // Token expirou — pede um novo e tenta de novo
      tokenClient.requestAccessToken();
      return;
    }
    if (resp.status === 403) {
      throw new Error('Sua conta não tem permissão para ver esta planilha. Peça acesso ao administrador.');
    }
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);

    const data = await resp.json();
    const wb = { SheetNames: [], Sheets: {} };
    data.valueRanges.forEach((vr, i) => {
      const name = GS_TABS[i];
      const values = vr.values || [];
      wb.SheetNames.push(name);
      wb.Sheets[name] = XLSX.utils.aoa_to_sheet(values);
    });

    parseWorkbook(wb, 'Google Sheets (login institucional)');
  } catch (err) {
    console.error('[Google Sheets] erro:', err);
    showBanner('Erro ao carregar do Google Sheets: ' + err.message, true);
  } finally {
    overlay.classList.remove('show');
  }
}

function showBanner(msg, isError = false) {
  const b = document.getElementById('importBanner');
  document.getElementById('importMsg').textContent = msg;
  b.classList.toggle('error', isError);
  b.classList.add('show');
}
function closeBanner() {
  document.getElementById('importBanner').classList.remove('show');
}

// ═══════════════════════════════════════════════════════
// RENDER ALL
// ═══════════════════════════════════════════════════════
function renderAll() {
  renderGeral();
  renderComercial();
  renderFinanceiro();
  renderProjetos();
  renderOrcamento();
}

// ── HELPERS ──
function getMesesData() {
  const mesesPresentes = [...new Set(DB.transacoes.map(r => r.Mes))]
    .filter(m => MESES_ORDER.includes(m))
    .sort((a, b) => MESES_ORDER.indexOf(a) - MESES_ORDER.indexOf(b));
  if (!mesesPresentes.length) return { meses:[], receitas:[], despesas:[], balancos:[] };
  
  const receitas = [], despesas = [], balancos = [];
  mesesPresentes.forEach(mes => {
    const rows = DB.transacoes.filter(r => r.Mes === mes);
    const rec = sumBy(rows.filter(r => r.EntradaOuSaida === 'Entrada'), 'Valor');
    const desp = Math.abs(sumBy(rows.filter(r => r.EntradaOuSaida === 'Saída'), 'Valor'));
    receitas.push(rec);
    despesas.push(desp);
    balancos.push(rec - desp);
  });
  return { meses: mesesPresentes, receitas, despesas, balancos };
}

// ══ GERAL ══
function renderGeral() {
  const receitas = sumBy(DB.transacoes.filter(r => r.EntradaOuSaida === 'Entrada'), 'Valor');
  const despesas = Math.abs(sumBy(DB.transacoes.filter(r => r.EntradaOuSaida === 'Saída'), 'Valor'));
  const balanco = receitas - despesas;
  const margem = receitas > 0 ? balanco / receitas : 0;

  document.getElementById('kpi-receita').textContent = fmt(receitas);
  document.getElementById('kpi-despesa').textContent = fmt(despesas);
  document.getElementById('kpi-balanco').textContent = fmt(balanco);
  document.getElementById('kpi-margem').textContent = pct(margem);
  document.getElementById('kpi-balanco-d').textContent = balanco >= 0 ? '↑ saldo positivo' : '↓ saldo negativo';
  document.getElementById('kpi-balanco-d').className = 'kpi-delta ' + (balanco >= 0 ? 'pos' : 'neg');
  document.getElementById('kpi-margem-d').textContent = 'da receita total';

  // Bar chart
  const { meses, receitas: rArr, despesas: dArr, balancos: bArr } = getMesesData();
  const maxV = Math.max(...rArr, ...dArr, 1);
  makeBars('bar-geral','bar-geral-ax',
    meses.map((m,i) => ({
      name: m,
      labels:['Receita','Despesa','Balanço'],
      values:[rArr[i], dArr[i], Math.max(bArr[i],0)],
      colors:['#365833','#A32E16','#427882']
    })), maxV * 1.1);

  // Donut
  const despPct = receitas > 0 ? Math.round(despesas/receitas*100) : 0;
  const balPct = 100 - despPct;
  document.getElementById('donut-desp-pct').textContent = despPct + '%';
  document.getElementById('donut-bal-pct').textContent = balPct + '%';
  if (donutChart) { donutChart.data.datasets[0].data = [despPct, Math.max(balPct,0)]; donutChart.update(); }

  // Prog receita
  const catRec = groupBy(DB.transacoes.filter(r=>r.EntradaOuSaida === 'Entrada'), 'Categoria');
  const catRecArr = Object.entries(catRec).map(([k,v]) => ({k, v: sumBy(v,'Valor')})).sort((a,b)=>b.v-a.v);
  const maxRec = catRecArr[0]?.v || 1;
  const colors = ['#365833','#427882','#799980','#967D69','#577A55'];
  document.getElementById('prog-receita').innerHTML = catRecArr.map((c,i) => `
    <div class="prog-item">
      <div class="prog-header"><span class="prog-name">${c.k}</span><span class="prog-val">${fmt(c.v)}</span></div>
      <div class="prog-bar-bg"><div class="prog-bar-fill" style="width:${Math.round(c.v/maxRec*100)}%;background:${colors[i%colors.length]}"></div></div>
    </div>`).join('');

  // Prog gastos
  const catDesp = groupBy(DB.transacoes.filter(r=>r.EntradaOuSaida === 'Saída'), 'Categoria');
  const catDespArr = Object.entries(catDesp).map(([k,v]) => ({k, v: Math.abs(sumBy(v,'Valor'))})).sort((a,b)=>b.v-a.v);
  const maxDesp = catDespArr[0]?.v || 1;
  document.getElementById('prog-gastos').innerHTML = catDespArr.map(c => `
    <div class="prog-item">
      <div class="prog-header"><span class="prog-name">${c.k}</span><span class="prog-val">${fmt(c.v)}</span></div>
      <div class="prog-bar-bg"><div class="prog-bar-fill" style="width:${Math.round(c.v/maxDesp*100)}%;background:#A32E16"></div></div>
    </div>`).join('');

  // Tabela transações (últimas 10)
  const recent = [...DB.transacoes].reverse().slice(0,10);
  document.getElementById('tbody-transacoes').innerHTML = recent.map(r => {
    const isRec = r.Tipo === 'Receita';
    const dataStr = r.Data instanceof Date ? r.Data.toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit'}) : String(r.Data).substring(5,10).replace('-','/');
    return `<tr>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${dataStr}</td>
      <td>${r.Descricao}</td><td>${r.Categoria}</td><td>${r.Tipo_Cliente}</td>
      <td><span class="badge ${isRec?'badge-green':'badge-red'}">${r.Tipo}</span></td>
      <td style="font-family:'DM Mono',monospace;color:var(--${isRec?'green':'red'}-mid)">${isRec?'+':'-'}${fmtFull(r.Valor)}</td>
    </tr>`;
  }).join('');

  document.getElementById('data-source-geral').innerHTML = `<span>Fonte: ${DB.filename} · ${DB.transacoes.length} transações</span>`;
}

// ══ COMERCIAL ══
function renderComercial() {
  const novos = DB.clientes.filter(c => c.Tipo === 'Novo').length;
  const rec   = DB.clientes.filter(c => c.Tipo === 'Recorrente').length;
  const totalCli = DB.clientes.length || 1;
  const receitaTotal = sumBy(DB.transacoes.filter(r=>r.Tipo==='Receita'),'Valor');
  const recRec = sumBy(DB.transacoes.filter(r=>r.Tipo==='Receita'&&r.Tipo_Cliente==='Recorrente'),'Valor');
  const recNov = receitaTotal - recRec;
  const ticket = DB.projetos.length ? sumBy(DB.projetos,'Receita') / DB.projetos.length : 0;
  const canais = [...new Set(DB.clientes.map(c=>c.Canal_Aquisicao))].filter(Boolean).length;

  document.getElementById('kpi-novos-cli').textContent = novos;
  document.getElementById('kpi-novos-cli-d').textContent = `de ${totalCli} clientes totais`;
  document.getElementById('kpi-ticket').textContent = fmt(ticket);
  document.getElementById('kpi-recorrente').textContent = fmt(recRec);
  document.getElementById('kpi-recorrente-d').textContent = receitaTotal>0 ? pct(recRec/receitaTotal)+' da receita' : '—';
  document.getElementById('kpi-nova').textContent = fmt(recNov);
  document.getElementById('kpi-nova-d').textContent = receitaTotal>0 ? pct(recNov/receitaTotal)+' da receita' : '—';
  document.getElementById('gauge-cli-rec').textContent = rec;
  document.getElementById('gauge-canais').textContent = canais;
  document.getElementById('gauge-rec-cli').textContent = fmt(totalCli>0 ? receitaTotal/totalCli : 0);

  // Bar recorrente vs nova por mês
  const { meses, receitas: rArr } = getMesesData();
  const maxV = Math.max(...rArr, 1);
  makeBars('bar-recorrente','bar-recorrente-ax',
    meses.map((m,i) => ({
      name: m,
      labels:['Recorrente','Nova'],
      values:[Math.round(rArr[i]*0.37), Math.round(rArr[i]*0.63)],
      colors:['#427882','#967D69']
    })), maxV * 1.1);

  // Canais prog
  const canalGrp = groupBy(DB.clientes, 'Canal_Aquisicao');
  const canalArr = Object.entries(canalGrp).map(([k,v])=>({k,v:v.length})).sort((a,b)=>b.v-a.v);
  const maxCan = canalArr[0]?.v || 1;
  document.getElementById('prog-canais').innerHTML = canalArr.map((c,i) => `
    <div class="prog-item">
      <div class="prog-header"><span class="prog-name">${c.k}</span><span class="prog-val">${c.v} cliente${c.v>1?'s':''}</span></div>
      <div class="prog-bar-bg"><div class="prog-bar-fill" style="width:${Math.round(c.v/maxCan*100)}%;background:${['#365833','#427882','#967D69','#799980'][i%4]}"></div></div>
    </div>`).join('') || '<p class="note">Sem dados de clientes.</p>';
}

// ══ FINANCEIRO ══
function renderFinanceiro() {
  const { meses, receitas: rArr, despesas: dArr, balancos: bArr } = getMesesData();
  const n = meses.length || 1;
  const recTotal = rArr.reduce((s,v)=>s+v,0);
  const dspTotal = dArr.reduce((s,v)=>s+v,0);
  const recMedia = recTotal / n;
  const dspMedia = dspTotal / n;
  const balTotal = recTotal - dspTotal;
  const margMedia = recTotal > 0 ? balTotal / recTotal : 0;

  let melhorMes = '—', melhorVal = 0;
  bArr.forEach((v,i) => { if(v > melhorVal) { melhorVal = v; melhorMes = meses[i]; } });

  document.getElementById('kpi-rec-media').textContent = fmt(recMedia);
  document.getElementById('kpi-desp-media').textContent = fmt(dspMedia);
  document.getElementById('kpi-melhor-mes').textContent = melhorMes;
  document.getElementById('kpi-melhor-mes-v').textContent = fmt(melhorVal);
  document.getElementById('kpi-marg-media').textContent = pct(margMedia);
  document.getElementById('gauge-burn').textContent = fmt(dspTotal);
  document.getElementById('gauge-runway').textContent = fmt(balTotal);
  document.getElementById('gauge-pct-desp').textContent = recTotal > 0 ? pct(dspTotal/recTotal) : '—';

  // Balanço bars
  const maxBal = Math.max(...bArr.map(Math.abs), 1);
  makeBars('bar-balanco','bar-balanco-ax',
    meses.map((m,i) => ({
      name: m, labels:['Balanço'],
      values:[Math.abs(bArr[i])],
      colors:[bArr[i]>=0?'#365833':'#A32E16']
    })), maxBal * 1.2);

  // Burn bars
  const maxD = Math.max(...dArr, 1);
  makeBars('bar-burn','bar-burn-ax',
    meses.map((m,i) => ({
      name: m, labels:['Despesas'],
      values:[dArr[i]], colors:['#A32E16']
    })), maxD * 1.2);

  // Orç vs realizado table
  document.getElementById('tbody-orca-fin').innerHTML = DB.orcamento.map(r => {
    const v = r.Variacao;
    return `<tr>
      <td>${r.Categoria}</td>
      <td style="font-family:'DM Mono',monospace">${fmtFull(r.Orcado)}</td>
      <td style="font-family:'DM Mono',monospace">${fmtFull(r.Realizado)}</td>
      <td style="font-family:'DM Mono',monospace;color:var(--${v<=0?'green':'red'}-mid)">${v<=0?'-':'+'}${fmtFull(Math.abs(v))}</td>
      <td><span class="badge badge-${r.Status==='OK'?'green':r.Status==='Excedeu'?'red':'amber'}">${r.Status}</span></td>
    </tr>`;
  }).join('') + (() => {
    if (!DB.orcamento.length) return '';
    const totO = sumBy(DB.orcamento,'Orcado'), totR = sumBy(DB.orcamento,'Realizado');
    const v = totR - totO;
    return `<tr style="font-weight:600"><td>Total</td>
      <td style="font-family:'DM Mono',monospace">${fmtFull(totO)}</td>
      <td style="font-family:'DM Mono',monospace">${fmtFull(totR)}</td>
      <td style="font-family:'DM Mono',monospace;color:var(--${v<=0?'green':'red'}-mid)">${v<=0?'-':'+'}${fmtFull(Math.abs(v))}</td>
      <td><span class="badge badge-${v<=0?'green':'amber'}">${v<=0?'OK':'Atenção'}</span></td>
    </tr>`;
  })();
}

// ══ PROJETOS ══
function renderProjetos() {
  const sorted = [...DB.projetos].sort((a,b) => b.Margem - a.Margem);
  const total = DB.projetos.length;
  const recProj = sumBy(DB.projetos,'Receita');
  const recTotal = sumBy(DB.transacoes.filter(r=>r.Tipo==='Receita'),'Valor');
  const margens = DB.projetos.map(p => p.Margem).filter(m => isFinite(m));
  const margMedia = margens.length ? margens.reduce((s,v)=>s+v,0)/margens.length : 0;
  const prejuizos = DB.projetos.filter(p => p.Lucro < 0).length;

  document.getElementById('kpi-tot-proj').textContent = total;
  document.getElementById('kpi-rec-proj').textContent = fmt(recProj);
  document.getElementById('kpi-rec-proj-d').textContent = recTotal>0 ? pct(recProj/recTotal)+' da receita total' : '';
  document.getElementById('kpi-marg-proj').textContent = pct(margMedia);
  document.getElementById('kpi-prej-proj').textContent = prejuizos;
  document.getElementById('kpi-prej-d').className = 'kpi-delta ' + (prejuizos > 0 ? 'neg' : 'pos');
  document.getElementById('kpi-prej-d').textContent = prejuizos > 0 ? 'requer atenção' : 'todos lucrativos';

  const getBadge = (m) => {
    if (m >= 0.40) return 'green';
    if (m >= 0.20) return 'blue';
    if (m >= 0) return 'amber';
    return 'red';
  };

  document.getElementById('tbody-projetos').innerHTML = sorted.map((p,i) => `
    <tr>
      <td><span class="rank-num">${i+1}</span></td>
      <td style="font-weight:500">${p.Projeto}</td><td>${p.Cliente}</td>
      <td style="font-family:'DM Mono',monospace;color:var(--green-mid)">${fmtFull(p.Receita)}</td>
      <td style="font-family:'DM Mono',monospace">${fmtFull(p.Custo)}</td>
      <td style="font-family:'DM Mono',monospace;font-weight:600;color:var(--${p.Lucro>=0?'green':'red'}-mid)">${p.Lucro>=0?'+':'-'}${fmtFull(Math.abs(p.Lucro))}</td>
      <td><span class="badge badge-${getBadge(p.Margem)}">${pct(p.Margem)}</span></td>
    </tr>`).join('');

  const maxMarg = Math.max(...sorted.map(p=>Math.abs(p.Margem)), 0.01);
  document.getElementById('prog-projetos').innerHTML = sorted.map(p => {
    const w = Math.round(Math.abs(p.Margem)/maxMarg*100);
    const col = p.Margem >= 0.4 ? '#365833' : p.Margem >= 0.2 ? '#427882' : p.Margem >= 0 ? '#967D69' : '#A32E16';
    return `<div class="prog-item">
      <div class="prog-header"><span class="prog-name">${p.Projeto}</span><span class="prog-val" style="color:${p.Margem<0?'var(--red-mid)':'inherit'}">${pct(p.Margem)}</span></div>
      <div class="prog-bar-bg"><div class="prog-bar-fill" style="width:${w}%;background:${col}"></div></div>
    </div>`;
  }).join('');
}

// ══ ORÇAMENTO ══
function renderOrcamento() {
  const totO = sumBy(DB.orcamento,'Orcado');
  const totR = sumBy(DB.orcamento,'Realizado');
  const dev = totO > 0 ? (totR-totO)/totO : 0;
  const ok = DB.orcamento.filter(r=>r.Status==='OK').length;

  document.getElementById('kpi-orc-total').textContent = fmt(totO);
  document.getElementById('kpi-orc-real').textContent = fmt(totR);
  document.getElementById('kpi-orc-real-d').textContent = (totR>totO?'+':'-') + fmtFull(Math.abs(totR-totO)) + ' vs orçado';
  document.getElementById('kpi-orc-real-d').className = 'kpi-delta ' + (totR<=totO?'pos':'neg');
  document.getElementById('kpi-orc-desvio').textContent = (dev>=0?'+':'')+pct(dev);
  document.getElementById('kpi-orc-ok').textContent = ok + ' / ' + DB.orcamento.length;

  // Bar
  const maxV = Math.max(...DB.orcamento.map(r=>Math.max(r.Orcado,r.Realizado)), 1);
  makeBars('bar-orca','bar-orca-ax',
    DB.orcamento.map(r => ({
      name: r.Categoria.substring(0,6),
      labels:['Orçado','Realizado'],
      values:[r.Orcado, r.Realizado],
      colors:['#427882','#A32E16']
    })), maxV * 1.15);

  // Cat table
  const totRFull = sumBy(DB.orcamento,'Realizado') || 1;
  document.getElementById('tbody-orc-cat').innerHTML = DB.orcamento.map(r => {
    const p = Math.round(r.Realizado/totRFull*100);
    const cl = p >= 30 ? 'red' : p >= 20 ? 'amber' : 'blue';
    return `<tr><td>${r.Categoria}</td>
      <td style="font-family:'DM Mono',monospace">${fmtFull(r.Realizado)}</td>
      <td><span class="badge badge-${cl}">${p}%</span></td>
    </tr>`;
  }).join('');
}

// ═══════════════════════════════════════════════════════
// BAR CHART BUILDER
// ═══════════════════════════════════════════════════════
function makeBars(containerId, axId, groups, maxVal) {
  const c = document.getElementById(containerId);
  const ax = document.getElementById(axId);
  if (!c) return;
  c.innerHTML = '';
  if (ax) ax.innerHTML = '';
  const max = maxVal || 1;
  groups.forEach(g => {
    const grp = document.createElement('div');
    grp.className = 'bar-group';
    g.values.forEach((v, vi) => {
      const b = document.createElement('div');
      b.className = 'bar';
      b.style.height = Math.max(Math.round((v/max)*100), v>0?2:0) + '%';
      b.style.background = g.colors[vi];
      b.title = `${g.labels[vi]}: ${fmt(v)}`;
      grp.appendChild(b);
    });
    c.appendChild(grp);
    if (ax) {
      const l = document.createElement('div');
      l.className = 'bar-axis-lbl';
      l.textContent = g.name;
      ax.appendChild(l);
    }
  });
}

// ═══════════════════════════════════════════════════════
// TABS
// ═══════════════════════════════════════════════════════
function showBlock(id, btn) {
  document.querySelectorAll('.block').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('block-' + id).classList.add('active');
  btn.classList.add('active');
}

// ═══════════════════════════════════════════════════════
// INIT DONUT
// ═══════════════════════════════════════════════════════
window.addEventListener('DOMContentLoaded', () => {
  const ctx = document.getElementById('donut-renda');
  if (ctx) {
    donutChart = new Chart(ctx, {
      type: 'doughnut',
      data: { datasets: [{ data: [65,35], backgroundColor: ['#A32E16','#427882'], borderWidth: 0 }] },
      options: { responsive: false, cutout: '68%', plugins: { legend: { display: false }, tooltip: { enabled: false } } }
    });
  }
  // O carregamento dos dados só acontece após o login com Google (ver iniciarLogin / handleTokenResponse)
});
