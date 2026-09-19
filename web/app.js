(() => {
  'use strict';

  const DATA = window.DASHBOARD_DATA || { registros: [], geradoEm: '' };
  const FARM = String(DATA.fazenda || DATA.codigoFazenda || '').trim();
  const logoStyle = document.createElement('style');
  logoStyle.textContent = '.brand-logo{display:block;width:260px;max-height:74px;object-fit:contain;object-position:left center;padding:3px}.brand{min-width:286px}.controls{display:flex!important;gap:12px;align-items:center;flex-wrap:wrap;background:#fff;border:1px solid #e4e9ed;border-radius:12px;padding:10px 14px;margin:10px 0;box-shadow:0 2px 7px rgba(27,41,54,.05)}.controls select,.controls input,.controls button{height:36px;border:1px solid #d8e0e4;border-radius:7px;background:#fff;padding:0 10px;font:14px "Segoe UI",Arial,sans-serif}.controls input{min-width:220px}.controls button{font-weight:700;cursor:pointer}.calendar svg{display:block;width:34px;height:34px;stroke:#fff;fill:none;stroke-width:2.4}.donut .divider{stroke:#fff;stroke-width:3.2}.callout.done{right:-58px!important}.callout.pending{left:-62px!important}.hero-date{white-space:pre-line}.capture-mode .controls{display:none!important}.capture-mode .page{max-width:none;width:1400px;padding:0 10px 28px}@media(max-width:900px){.brand-logo{width:230px}.controls input{min-width:160px}}';
  document.head.appendChild(logoStyle);
  const unitFilterStyle = document.createElement('style');
  unitFilterStyle.textContent = '#unitSelect{display:none!important}';
  document.head.appendChild(unitFilterStyle);
  const downloadStyle = document.createElement('style');
  downloadStyle.textContent = '.controls a{height:36px;box-sizing:border-box;border:1px solid #d8e0e4;border-radius:7px;background:#fff;padding:0 10px;font:700 14px "Segoe UI",Arial,sans-serif;display:inline-flex;align-items:center;cursor:pointer;text-decoration:none;color:#101424}';
  document.head.appendChild(downloadStyle);
  document.querySelector('.brand-logo').src = window.LOGO_EXPORT_DATA_URI || './logoboaesperanca-branca.png';
  const qs = new URLSearchParams(location.search);
  const captureMode = window.DASHBOARD_CAPTURE === true || qs.get('capture') === '1';
  if (captureMode) document.body.classList.add('capture-mode');

  const $ = (s) => document.querySelector(s);
  const cardsEl = $('#cards');
  const unitEl = $('#unitSelect');
  const searchEl = $('#searchInput');
  const sortEl = $('#sortSelect');
  const summaryEl = $('#summary');
  const heroSubEl = $('#heroSub');
  const heroDateEl = $('#heroDate');

  const norm = (v) => String(v ?? '').trim();
  const key = (v) => norm(v).toLocaleUpperCase('pt-BR');
  const num = (v) => Number(v) || 0;
  const clamp = (v) => Math.max(0, Math.min(1, Number(v) || 0));
  const fmtHa = (v) => new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 1 }).format(num(v)) + ' ha';
  const fmtPct = (v) => Math.round(clamp(v) * 100) + '%';
  const subtitle = () => `${FARM ? `FAZENDA: ${FARM} • ` : ''}Atualizado em ${DATA.geradoEm || '—'}`;

  function getUnits() {
    return [...new Set(DATA.registros.map(r => norm(r.unidade)).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR'));
  }

  function initUnits() {
    unitEl.innerHTML = '';
    ['TODAS', ...getUnits()].forEach(u => {
      const op = document.createElement('option');
      op.value = u; op.textContent = u; unitEl.appendChild(op);
    });
    const requested = norm(window.DASHBOARD_SELECTED_UNIT || qs.get('unidade')).toUpperCase();
    if (requested && [...unitEl.options].some(o => o.value.toUpperCase() === requested)) unitEl.value = [...unitEl.options].find(o => o.value.toUpperCase() === requested).value;
  }

  function aggregate(unit) {
    const map = new Map();
    DATA.registros.forEach((r, idx) => {
      if (unit !== 'TODAS' && key(r.unidade) !== key(unit)) return;
      const op = norm(r.operacao);
      if (!op) return;
      const opKey = key(op);
      if (!map.has(opKey)) map.set(opKey, { operacao: op, planejado:0, realizado:0, pendente:0, concluidos:0, pendTalhoes:0, order:idx });
      const x = map.get(opKey);
      x.planejado += num(r.planejado); x.realizado += num(r.realizado); x.pendente += num(r.pendente);
      x.concluidos += num(r.talhoesConcluidos); x.pendTalhoes += num(r.talhoesPendentes);
    });
    return [...map.values()].map(x => ({...x, execucao: x.planejado > 0 ? clamp(x.realizado/x.planejado) : 0, pendPct: x.planejado > 0 ? clamp(x.pendente/x.planejado) : 1 }));
  }

  function status(p) {
    if (p >= .9999) return ['CONCLUÍDO','final'];
    // Mantém a mesma regra do painel do Excel: a partir de 80%, a operação
    // entra em finalização; a partir de 100%, fica concluída.
    if (p >= .8) return ['FINALIZANDO','final'];
    if (p >= .5) return ['EM ANDAMENTO',''];
    return ['ATENÇÃO',''];
  }

  function donutHtml(p, pending) {
    const r = 60, c = 2 * Math.PI * r, dash = c * clamp(p);
    const angle = (Math.PI * 2 * clamp(p)) - (Math.PI / 2);
    const x = (80 + r * Math.cos(angle)).toFixed(2), y = (80 + r * Math.sin(angle)).toFixed(2);
    return `<div class="donut"><svg viewBox="0 0 160 160" aria-hidden="true"><circle class="base" cx="80" cy="80" r="${r}"></circle><circle class="done" cx="80" cy="80" r="${r}" stroke-dasharray="${dash} ${c-dash}"></circle><line class="divider" x1="80" y1="80" x2="80" y2="20"></line><line class="divider" x1="80" y1="80" x2="${x}" y2="${y}"></line></svg><span class="callout pending">${fmtPct(pending)}</span><span class="callout done">${fmtPct(p)}</span><div class="donut-center"><strong>${fmtPct(p)}</strong><span>EXECUTADO</span></div></div>`;
  }

  function cardHtml(x) {
    const [st, cls] = status(x.execucao);
    return `<article class="card">
      <div class="card-head"><div><div class="kicker">OPERAÇÃO</div><h2 class="op-name">${escapeHtml(x.operacao)}</h2></div><div class="pill ${cls}">${st}</div></div>
      <div class="main">
        <div class="donut-wrap">${donutHtml(x.execucao, x.pendPct)}</div>
        <div class="metrics">
          <div class="metric"><span class="dot blue"></span><div><small>Realizado</small><strong>${fmtHa(x.realizado)}</strong></div><span class="pct">${fmtPct(x.execucao)}</span></div>
          <div class="metric"><span class="dot red"></span><div><small>Pendente</small><strong>${fmtHa(x.pendente)}</strong></div><span class="pct">${fmtPct(x.pendPct)}</span></div>
          <div class="metric"><span class="dot gray"></span><div><small>Planejado</small><strong>${fmtHa(x.planejado)}</strong></div><span class="pct"></span></div>
        </div>
      </div>
      <div class="bar"><span style="width:${Math.round(x.execucao*100)}%"></span></div>
      <div class="footer"><span><b>${Math.round(x.concluidos)} de ${Math.round(x.concluidos + x.pendTalhoes)}</b> talhões concluídos</span><span>${Math.round(x.pendTalhoes)} pendentes</span></div>
    </article>`;
  }

  function escapeHtml(s){ return String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

  function currentRows() {
    let rows = aggregate(unitEl.value || 'TODAS');
    const q = norm(searchEl.value).toLocaleLowerCase('pt-BR');
    if (q) rows = rows.filter(x => x.operacao.toLocaleLowerCase('pt-BR').includes(q));
    switch (sortEl.value) {
      case 'progressAsc': rows.sort((a,b)=>a.execucao-b.execucao); break;
      case 'progressDesc': rows.sort((a,b)=>b.execucao-a.execucao); break;
      case 'pendingDesc': rows.sort((a,b)=>b.pendente-a.pendente); break;
      case 'name': rows.sort((a,b)=>a.operacao.localeCompare(b.operacao,'pt-BR')); break;
      default: rows.sort((a,b)=>a.order-b.order);
    }
    return rows;
  }

  function render() {
    const rows = currentRows();
    const unit = unitEl.value || 'TODAS';
    const p = rows.reduce((s,x)=>s+x.planejado,0), r = rows.reduce((s,x)=>s+x.realizado,0), pe = rows.reduce((s,x)=>s+x.pendente,0);
    const pg = p > 0 ? r/p : 0;
    heroSubEl.textContent = subtitle();
    heroDateEl.innerHTML = escapeHtml(DATA.geradoEm || '—').replace(' ', '<br>');
    summaryEl.innerHTML = `<span><strong>${rows.length}</strong> operações</span><span>Execução geral: <strong>${fmtPct(pg)}</strong></span><span>Planejado: <strong>${fmtHa(p)}</strong></span><span>Realizado: <strong>${fmtHa(r)}</strong></span><span>Pendente: <strong>${fmtHa(pe)}</strong></span>`;
    cardsEl.innerHTML = rows.length ? rows.map(cardHtml).join('') : '<div class="empty">Nenhuma operação encontrada para o filtro selecionado.</div>';
    document.title = `Operações - ${FARM || unit}`;
  }

  function serializeSvg(rows, unit) {
    const W=1400, pad=24, gap=26, cols=2, cardW=(W-pad*2-gap)/2, cardH=390, headerH=112, rowGap=26;
    const nRows=Math.max(1,Math.ceil(rows.length/cols)), H=headerH+26+nRows*cardH+(nRows-1)*rowGap+34;
    const esc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    const fHa = v => new Intl.NumberFormat('pt-BR',{maximumFractionDigits:1}).format(num(v))+' ha';
    const fPct = v => Math.round(clamp(v)*100)+'%';
    const font="Segoe UI,Arial,sans-serif";
    const chunks=[];
    chunks.push(`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><rect width="100%" height="100%" fill="#ffffff"/><rect x="0" y="0" width="${W}" height="${headerH}" fill="#5d554f"/><text x="34" y="46" font-family="${font}" font-size="31" font-weight="700" fill="#fff">ACOMPANHAMENTO DE OPERAÇÕES</text><text x="34" y="78" font-family="${font}" font-size="14" fill="#eee">UNIDADE: ${esc(unit)} • Atualizado em ${esc(DATA.geradoEm||'—')}</text>`);
    rows.forEach((x,i)=>{
      const col=i%2,row=Math.floor(i/2), x0=pad+col*(cardW+gap), y0=headerH+26+row*(cardH+rowGap);
      const [st,cls]=status(x.execucao), pillFill=cls==='final'?'#edf5dc':'#fff1d8', pillText=cls==='final'?'#5d7d25':'#aa6a00';
      const cx=x0+164, cy=y0+208, R=74, circ=2*Math.PI*R, dash=circ*clamp(x.execucao);
      const mX=x0+316;
      chunks.push(`<rect x="${x0}" y="${y0}" width="${cardW}" height="${cardH}" rx="30" fill="#fff" stroke="#d8dfdb" stroke-width="2"/><text x="${x0+30}" y="${y0+38}" font-family="${font}" font-size="12" font-weight="700" fill="#776957">OPERAÇÃO</text><text x="${x0+30}" y="${y0+78}" font-family="${font}" font-size="24" font-weight="700" fill="#202522">${esc(x.operacao)}</text><rect x="${x0+cardW-148}" y="${y0+22}" width="118" height="44" rx="11" fill="${pillFill}"/><text x="${x0+cardW-89}" y="${y0+49}" text-anchor="middle" font-family="${font}" font-size="12" font-weight="700" fill="${pillText}">${st}</text>`);
      chunks.push(`<circle cx="${cx}" cy="${cy}" r="${R}" fill="none" stroke="#ea0000" stroke-width="29"/><circle cx="${cx}" cy="${cy}" r="${R}" fill="none" stroke="#1397c9" stroke-width="29" stroke-dasharray="${dash} ${circ-dash}" transform="rotate(-90 ${cx} ${cy})"/><circle cx="${cx}" cy="${cy}" r="50" fill="#fff"/><text x="${cx}" y="${cy+3}" text-anchor="middle" font-family="${font}" font-size="34" font-weight="700" fill="#202522">${fPct(x.execucao)}</text><text x="${cx}" y="${cy+29}" text-anchor="middle" font-family="${font}" font-size="10" fill="#8a8f8c">EXECUTADO</text>`);
      const metric=(yy,color,label,value,pct)=>chunks.push(`<rect x="${mX}" y="${yy-14}" width="15" height="15" rx="3" fill="${color}"/><text x="${mX+27}" y="${yy-3}" font-family="${font}" font-size="12" fill="#949a97">${label}</text><text x="${mX+27}" y="${yy+25}" font-family="${font}" font-size="19" font-weight="700" fill="#202522">${value}</text>${pct?`<text x="${x0+cardW-40}" y="${yy-3}" text-anchor="end" font-family="${font}" font-size="13" fill="#8c928f">${pct}</text>`:''}`);
      metric(y0+125,'#1397c9','Realizado',fHa(x.realizado),fPct(x.execucao));
      metric(y0+199,'#ea0000','Pendente',fHa(x.pendente),fPct(x.pendPct));
      metric(y0+273,'#aeb8b4','Planejado',fHa(x.planejado),'');
      const barX=x0+30,barY=y0+326,barW=cardW-60;
      chunks.push(`<rect x="${barX}" y="${barY}" width="${barW}" height="12" rx="3" fill="#e8ece9"/><rect x="${barX}" y="${barY}" width="${barW*clamp(x.execucao)}" height="12" rx="3" fill="#1397c9"/><text x="${barX}" y="${y0+370}" font-family="${font}" font-size="12" fill="#8b918e">${Math.round(x.concluidos)} talhões concluídos</text><text x="${x0+cardW-30}" y="${y0+370}" text-anchor="end" font-family="${font}" font-size="12" fill="#8b918e">${Math.round(x.pendTalhoes)} pendentes</text>`);
    });
    chunks.push('</svg>');
    return {svg:chunks.join(''), width:W, height:H};
  }

  function serializeModern(rows, unit) {
    const W=1400, gap=18, pad=14, cardW=(W-pad*2-gap)/2, cardH=380, headH=112, font='Segoe UI,Arial,sans-serif';
    const h=headH+12+Math.ceil(rows.length/2)*(cardH+gap), esc=s=>String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
    const out=[`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${h}" viewBox="0 0 ${W} ${h}"><rect width="100%" height="100%" fill="#f5f8f9"/><rect x="10" y="8" width="1380" height="96" rx="12" fill="#514c48"/><text x="42" y="50" font-family="${font}" font-size="27" font-weight="700" fill="#43c8ff">Boa Esperança</text><text x="42" y="76" font-family="${font}" font-size="15" fill="#a5da24">Agropecuária</text><path d="M310 26v60" stroke="#fff" stroke-opacity=".7" stroke-width="2"/><text x="336" y="51" font-family="${font}" font-size="32" font-weight="800" fill="#fff">ACOMPANHAMENTO DE OPERAÇÕES</text><text x="336" y="79" font-family="${font}" font-size="18" fill="#fff">UNIDADE: ${esc(unit)} • Atualizado em ${esc(DATA.geradoEm||'—')}</text></svg>`];
    out.pop();
    rows.forEach((x,i)=>{const col=i%2,row=Math.floor(i/2),x0=pad+col*(cardW+gap),y0=headH+12+row*(cardH+gap),cx=x0+156,cy=y0+188,r=70,c=2*Math.PI*r,d=c*x.execucao,a=-Math.PI/2+Math.PI*2*x.execucao,ex=cx+r*Math.cos(a),ey=cy+r*Math.sin(a),total=Math.round(x.concluidos+x.pendTalhoes);out.push(`<g><rect x="${x0}" y="${y0}" width="${cardW}" height="${cardH}" rx="14" fill="#fff" stroke="#e4e9ed"/><text x="${x0+22}" y="${y0+30}" font-family="${font}" font-size="14" font-weight="700" fill="#566073">OPERAÇÃO</text><text x="${x0+22}" y="${y0+64}" font-family="${font}" font-size="31" font-weight="900" fill="#101424">${esc(x.operacao)}</text><circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#f70b10" stroke-width="28"/><circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#049ce5" stroke-width="28" stroke-dasharray="${d} ${c-d}" transform="rotate(-90 ${cx} ${cy})"/><line x1="${cx}" y1="${cy}" x2="${cx}" y2="${cy-r}" stroke="#fff" stroke-width="3"/><line x1="${cx}" y1="${cy}" x2="${ex}" y2="${ey}" stroke="#fff" stroke-width="3"/><text x="${cx}" y="${cy+3}" text-anchor="middle" font-family="${font}" font-size="39" font-weight="800">${fmtPct(x.execucao)}</text><text x="${cx}" y="${cy+27}" text-anchor="middle" font-family="${font}" font-size="13" font-weight="700">EXECUTADO</text><text x="${x0+54}" y="${y0+132}" font-family="${font}" font-size="23" font-weight="800" fill="#f70b10">${fmtPct(x.pendPct)}</text><text x="${x0+250}" y="${y0+122}" font-family="${font}" font-size="23" font-weight="800" fill="#049ce5">${fmtPct(x.execucao)}</text><rect x="${x0+350}" y="${y0+90}" width="24" height="24" rx="6" fill="#049ce5"/><text x="${x0+392}" y="${y0+108}" font-family="${font}" font-size="22">Realizado</text><text x="${x0+392}" y="${y0+142}" font-family="${font}" font-size="34" font-weight="800">${fmtHa(x.realizado)}</text><rect x="${x0+350}" y="${y0+164}" width="24" height="24" rx="6" fill="#f70b10"/><text x="${x0+392}" y="${y0+182}" font-family="${font}" font-size="22">Pendente</text><text x="${x0+392}" y="${y0+216}" font-family="${font}" font-size="34" font-weight="800">${fmtHa(x.pendente)}</text><rect x="${x0+350}" y="${y0+238}" width="24" height="24" rx="6" fill="#adb8bb"/><text x="${x0+392}" y="${y0+256}" font-family="${font}" font-size="22">Planejado</text><text x="${x0+392}" y="${y0+290}" font-family="${font}" font-size="34" font-weight="800">${fmtHa(x.planejado)}</text><rect x="${x0+22}" y="${y0+316}" width="${cardW-44}" height="18" rx="6" fill="#e5e9eb"/><rect x="${x0+22}" y="${y0+316}" width="${(cardW-44)*x.execucao}" height="18" rx="6" fill="#049ce5"/><text x="${x0+22}" y="${y0+365}" font-family="${font}" font-size="18"><tspan font-weight="800">${Math.round(x.concluidos)} de ${total}</tspan> talhões concluídos</text><text x="${x0+cardW-22}" y="${y0+365}" text-anchor="end" font-family="${font}" font-size="18">${Math.round(x.pendTalhoes)} pendentes</text></g>`)});out.push('</svg>');return {svg:out.join(''),width:W,height:h};
  }

  function downloadPNG(event) {
    try {
    const rows=currentRows(), unit=unitEl.value||'TODAS';
    if (!rows.length) return alert('Não há operações para exportar.');
    const fileName=`Resumo_Operacoes_${unit}_${new Date().toISOString().slice(0,10).replace(/-/g,'')}.png`;
    const W=1400,pad=14,gap=18,cardW=(W-pad*2-gap)/2,cardH=380,headH=112,H=headH+12+Math.ceil(rows.length/2)*(cardH+gap),scale=2;
    const logo = document.querySelector('.brand-logo');
    const canvas=document.createElement('canvas'),ctx=canvas.getContext('2d'); canvas.width=W*scale;canvas.height=H*scale;ctx.scale(scale,scale);const nativeRoundRect=ctx.roundRect.bind(ctx);ctx.roundRect=(...args)=>{ctx.beginPath();return nativeRoundRect(...args)};ctx.fillStyle='#f5f8f9';ctx.fillRect(0,0,W,H);ctx.fillStyle='#514c48';ctx.roundRect(10,8,1380,96,12);ctx.fill();
    const text=(s,x,y,size,weight='400',color='#101424',align='left')=>{ctx.font=`${weight} ${size}px Segoe UI,Arial`;ctx.fillStyle=color;ctx.textAlign=align;ctx.fillText(s,x,y)};
    if (logo && logo.complete && logo.naturalWidth) ctx.drawImage(logo, 38, 27, 250, 62); else { text('Boa Esperança',42,50,27,'700','#fff');text('Agropecuária',42,76,15,'400','#fff'); }
    text('ACOMPANHAMENTO DE OPERAÇÕES',336,51,32,'800','#fff');text(subtitle(),336,79,18,'400','#fff');
    rows.forEach((x,i)=>{const col=i%2,row=Math.floor(i/2),x0=pad+col*(cardW+gap),y0=headH+12+row*(cardH+gap),cx=x0+156,cy=y0+188,r=70,total=Math.round(x.concluidos+x.pendTalhoes);ctx.fillStyle='#fff';ctx.roundRect(x0,y0,cardW,cardH,14);ctx.fill();ctx.strokeStyle='#e4e9ed';ctx.stroke();text('OPERAÇÃO',x0+22,y0+30,14,'700','#566073');text(x.operacao,x0+22,y0+64,31,'800');ctx.lineWidth=28;ctx.beginPath();ctx.strokeStyle='#f70b10';ctx.arc(cx,cy,r,0,Math.PI*2);ctx.stroke();ctx.beginPath();ctx.strokeStyle='#049ce5';ctx.arc(cx,cy,r,-Math.PI/2,-Math.PI/2+Math.PI*2*x.execucao);ctx.stroke();ctx.lineWidth=3;ctx.strokeStyle='#fff';ctx.beginPath();ctx.moveTo(cx,cy);ctx.lineTo(cx,cy-r);ctx.moveTo(cx,cy);ctx.lineTo(cx+r*Math.cos(-Math.PI/2+Math.PI*2*x.execucao),cy+r*Math.sin(-Math.PI/2+Math.PI*2*x.execucao));ctx.stroke();text(fmtPct(x.execucao),cx,cy+3,39,'800','#101424','center');text('EXECUTADO',cx,cy+27,13,'700','#101424','center');text(fmtPct(x.pendPct),x0+45,y0+132,23,'800','#f70b10');text(fmtPct(x.execucao),x0+228,y0+122,23,'800','#049ce5');[['#049ce5','Realizado',x.realizado,90],['#f70b10','Pendente',x.pendente,164],['#adb8bb','Planejado',x.planejado,238]].forEach(([c,l,v,dy])=>{ctx.fillStyle=c;ctx.roundRect(x0+350,y0+dy,24,24,6);ctx.fill();text(l,x0+392,y0+dy+18,22);text(fmtHa(v),x0+392,y0+dy+52,34,'800')});ctx.fillStyle='#e5e9eb';ctx.roundRect(x0+22,y0+316,cardW-44,18,6);ctx.fill();ctx.fillStyle='#049ce5';ctx.roundRect(x0+22,y0+316,(cardW-44)*x.execucao,18,6);ctx.fill();text(`${Math.round(x.concluidos)} de ${total}`,x0+22,y0+365,18,'800');text('talhões concluídos',x0+122,y0+365,18);text(`${Math.round(x.pendTalhoes)} pendentes`,x0+cardW-22,y0+365,18,'400','#101424','right')});
    const dataUrl=canvas.toDataURL('image/png');
    const binary=atob(dataUrl.slice(dataUrl.indexOf(',')+1));
    const bytes=new Uint8Array(binary.length);
    for(let i=0;i<binary.length;i++) bytes[i]=binary.charCodeAt(i);
    const blob=new Blob([bytes],{type:'image/png'});
    const a=event.currentTarget;
    a.download=fileName;
    const previousUrl=a.dataset.exportUrl;
    if(previousUrl) URL.revokeObjectURL(previousUrl);
    a.href=URL.createObjectURL(blob);
    a.dataset.exportUrl=a.href;
    console.info('[exportacao-png] arquivo preparado', {nome:a.download, bytes:blob.size});
    } catch(error) {
      event.preventDefault();
      console.error('[exportacao-png] falhou', error);
      alert('Não foi possível preparar o PNG: ' + (error && error.message ? error.message : error));
    }
  }

  initUnits();
  $('.calendar').innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"></rect><path d="M7 3v4M17 3v4M3 10h18M8 14h2M14 14h2M8 18h2M14 18h2"></path></svg>';
  $('#downloadBtn').textContent = 'Baixar PNG HQ';
  unitEl.addEventListener('change',render); searchEl.addEventListener('input',render); sortEl.addEventListener('change',render);
  $('#reloadBtn').addEventListener('click',()=>location.reload()); $('#downloadBtn').addEventListener('pointerdown',downloadPNG);
  render();
})();
