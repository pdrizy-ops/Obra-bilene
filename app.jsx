const { useState, useEffect, useMemo, useCallback } = React;
const XLSX = window.XLSX;

/* ─── DESIGN TOKENS ─── */
const C = {
  bg: '#F2F2F4', white: '#FFFFFF', black: '#111111',
  gray100: '#F7F7F8', gray200: '#EBEBED', gray300: '#D4D4D8', gray500: '#9A9A9A', gray700: '#555555',
  lagoon: '#1F7A6C', amber: '#E8B923', terra: '#B5502E',
};
const STATUS = {
  concluido: { label: 'Concluído', color: C.lagoon, bg: '#1F7A6C18' },
  em_curso: { label: 'Em Curso', color: C.amber, bg: '#E8B92318' },
  nao_iniciado: { label: 'Pendente', color: C.gray500, bg: '#9A9A9A18' },
  aviso: { label: 'Acima do Orçamento', color: C.terra, bg: '#B5502E18' },
};
const IVA_RATE = 0.16;
const STORAGE_KEY = 'casalodge-obra-pwa-v1';

/* ─── DADOS BASE (Proposta Otheka Construções, 15/06/2026) ─── */
const DEFAULT_COMPONENTS = [
  { id: 'suite-simples', name: 'Quarto Suite Simples', qty: 3, phases: [
    { code: 'BoQ0', label: 'Caixa de Fundação', unitBudget: 182500 },
    { code: 'BoQ1', label: 'Estruturas de betão', unitBudget: 244012 },
    { code: 'BoQ2', label: 'Alvenaria e Rebocos', unitBudget: 292203 },
    { code: 'BoQ3', label: 'Canalização, Esgotos e Electricidade', unitBudget: 140330 },
    { code: 'BoQ4', label: 'Trabalhos Exteriores, Cobertura e Deck', unitBudget: 484002 },
  ]},
  { id: 'suite-master', name: 'Quarto Suite Master', qty: 1, phases: [
    { code: 'BoQ0', label: 'Caixa de Fundação', unitBudget: 310220 },
    { code: 'BoQ1', label: 'Estruturas de betão', unitBudget: 253000 },
    { code: 'BoQ2', label: 'Alvenaria e Rebocos', unitBudget: 411260 },
    { code: 'BoQ3', label: 'Canalização, Esgotos e Electricidade', unitBudget: 175033 },
    { code: 'BoQ4', label: 'Trabalhos Exteriores, Cobertura e Deck', unitBudget: 644302 },
  ]},
  { id: 'alpendre-braai', name: 'Alpendre-Braai', qty: 1, phases: [
    { code: 'BoQ0', label: 'Caixa de Fundação', unitBudget: 420112 },
    { code: 'BoQ1', label: 'Estruturas de betão', unitBudget: 343102 },
    { code: 'BoQ2', label: 'Alvenaria e Rebocos', unitBudget: 253002 },
    { code: 'BoQ3', label: 'Canalização, Esgotos e Electricidade', unitBudget: 64003 },
    { code: 'BoQ4', label: 'Cobertura e Estrutura de madeira', unitBudget: 634770 },
  ]},
  { id: 'garagem', name: 'Garagem', qty: 1, phases: [
    { code: 'BoQ0', label: 'Caixa de Fundação', unitBudget: 220440 },
    { code: 'BoQ1', label: 'Estruturas de betão', unitBudget: 84338 },
    { code: 'BoQ2', label: 'Alvenaria e Rebocos', unitBudget: 162933 },
    { code: 'BoQ3', label: 'Canalização, Esgotos e Electricidade', unitBudget: 0 },
    { code: 'BoQ4', label: 'Trabalhos Exteriores e Cobertura', unitBudget: 180244 },
  ]},
  { id: 'alpendre-2pisos', name: 'Alpendre em 2 Pisos', qty: 1, phases: [
    { code: 'BoQ0', label: 'Caixa de Fundação', unitBudget: 644033 },
    { code: 'BoQ1', label: 'Estruturas de betão', unitBudget: 823309 },
    { code: 'BoQ2', label: 'Alvenaria e Rebocos', unitBudget: 747730 },
    { code: 'BoQ3', label: 'Canalização, Esgotos e Electricidade', unitBudget: 282000 },
    { code: 'BoQ4', label: 'Cobertura e Estrutura de madeira', unitBudget: 654700 },
  ]},
  { id: 'cozinha-trab', name: 'Cozinha dos Trabalhadores', qty: 1, phases: [
    { code: 'BoQ0', label: 'Caixa de Fundação', unitBudget: 330223 },
    { code: 'BoQ1', label: 'Estruturas de betão', unitBudget: 119230 },
    { code: 'BoQ2', label: 'Alvenaria e Rebocos', unitBudget: 97339 },
    { code: 'BoQ3', label: 'Canalização, Esgotos e Electricidade', unitBudget: 82344 },
    { code: 'BoQ4', label: 'Trabalhos Exteriores e Cobertura', unitBudget: 387990 },
  ]},
  { id: 'piscina', name: 'Piscina', qty: 1, phases: [
    { code: 'BoQ0', label: 'Escavação e Murro de conteção', unitBudget: 388003 },
    { code: 'BoQ1', label: 'Betonagem do Murro de conteção e reboco', unitBudget: 650043 },
    { code: 'BoQ2', label: 'Revestimento e Pintura', unitBudget: 790330 },
    { code: 'BoQ3', label: 'Canalização Backwash e Electricidade', unitBudget: 1364800 },
    { code: 'BoQ4', label: 'Deck de Madeira', unitBudget: 1810004 },
  ]},
];

function uid(prefix = 'id') { return `${prefix}-${Math.random().toString(36).slice(2, 10)}`; }
function nowLabel() {
  return new Date().toLocaleString('pt-PT', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' });
}
function freshState() {
  return {
    components: DEFAULT_COMPONENTS.map(c => ({
      id: c.id, name: c.name, qty: c.qty,
      phases: c.phases.map((p, i) => ({
        id: `${c.id}-${i}`, code: p.code, label: p.label, unitBudget: p.unitBudget,
        status: 'nao_iniciado', progress: 0, labor: 0, materials: [],
      })),
    })),
    activities: [{ id: uid('a'), time: nowLabel(), desc: 'Projecto criado a partir da proposta Otheka Construções' }],
  };
}
function fmt(n) { return (n || 0).toLocaleString('pt-MZ', { minimumFractionDigits: 0, maximumFractionDigits: 0 }); }

/* ─── PERSISTÊNCIA (localStorage — app standalone) ─── */
function loadInitial() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) { /* ignore */ }
  return freshState();
}
function saveState(state) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); return true; }
  catch (e) { return false; }
}

/* ─── CÁLCULOS ─── */
function phaseBudget(p, qty) { return p.unitBudget * qty; }
function materialsTotal(p) { return p.materials.reduce((s, m) => s + m.qty * m.unitCost, 0); }
function phaseActual(p) { return materialsTotal(p) + (p.labor || 0); }
function componentTotals(c) {
  let budgetSub = 0, actual = 0;
  c.phases.forEach(p => { budgetSub += phaseBudget(p, c.qty); actual += phaseActual(p); });
  const iva = budgetSub * IVA_RATE;
  return { budgetSub, iva, budget: budgetSub + iva, actual };
}
function componentProgress(c) {
  if (!c.phases.length) return 0;
  return Math.round(c.phases.reduce((s, p) => s + p.progress, 0) / c.phases.length);
}
function componentStatus(c) {
  const statuses = c.phases.map(p => p.status);
  if (!statuses.length) return 'nao_iniciado';
  if (statuses.every(s => s === 'concluido')) return 'concluido';
  if (statuses.some(s => s === 'em_curso' || s === 'concluido')) return 'em_curso';
  return 'nao_iniciado';
}

/* ─── PRIMITIVOS ─── */
function Card({ children, style = {}, black = false }) {
  return (
    <div style={{ background: black ? C.black : C.white, borderRadius: 20, padding: 24, boxShadow: black ? '0 8px 32px #00000030' : '0 2px 12px #00000009', ...style }}>{children}</div>
  );
}
function Badge({ status }) {
  const s = STATUS[status];
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: s.color, color: 'white', fontSize: 11, fontWeight: 600, padding: '5px 12px', borderRadius: 99, letterSpacing: '0.01em', whiteSpace: 'nowrap' }}>
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'white', opacity: 0.75, flexShrink: 0 }} />{s.label}
    </span>
  );
}
function TapeBar({ pct, over = false }) {
  return (
    <div style={{ height: 10, background: C.gray200, borderRadius: 99, overflow: 'hidden', position: 'relative' }}>
      <div style={{ position: 'absolute', inset: 0, backgroundImage: 'repeating-linear-gradient(90deg, #00000008 0 1px, transparent 1px 10%)' }} />
      <div style={{ width: `${Math.min(pct, 100)}%`, height: '100%', background: over ? C.terra : C.lagoon, borderRadius: 99, transition: 'width .5s' }} />
    </div>
  );
}
function Timeline({ steps, active }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center' }}>
      {steps.map((s, i) => {
        const done = i < active, curr = i === active;
        return (
          <div key={s} style={{ display: 'flex', alignItems: 'center', flex: i < steps.length - 1 ? 1 : 0, minWidth: 0 }}>
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', flexShrink: 0 }}>
              <div style={{ width: 10, height: 10, borderRadius: '50%', flexShrink: 0, background: done ? C.lagoon : curr ? C.amber : C.gray300, border: curr ? `2px solid ${C.amber}` : 'none', boxShadow: curr ? `0 0 0 3px ${C.amber}22` : 'none' }} />
              <span style={{ fontSize: 10, marginTop: 5, whiteSpace: 'nowrap', fontWeight: done ? 500 : curr ? 600 : 400, color: done ? C.gray700 : curr ? C.black : C.gray500 }}>{done ? '✓ ' : ''}{s}</span>
            </div>
            {i < steps.length - 1 && <div style={{ flex: 1, height: 1, background: i < active ? C.lagoon : C.gray200, margin: '0 6px', marginBottom: 18 }} />}
          </div>
        );
      })}
    </div>
  );
}
function Icon({ name, size = 18, color = C.gray700 }) {
  const s = { width: size, height: size, style: { display: 'inline-block', flexShrink: 0 } };
  const icons = {
    home: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M3 8.5L10 2l7 6.5V18H13v-4h-6v4H3V8.5Z" stroke={color} strokeWidth="1.4" strokeLinejoin="round" /></svg>,
    grid: <svg {...s} viewBox="0 0 20 20" fill="none"><rect x="3" y="3" width="6" height="6" rx="1.5" stroke={color} strokeWidth="1.4" /><rect x="11" y="3" width="6" height="6" rx="1.5" stroke={color} strokeWidth="1.4" /><rect x="3" y="11" width="6" height="6" rx="1.5" stroke={color} strokeWidth="1.4" /><rect x="11" y="11" width="6" height="6" rx="1.5" stroke={color} strokeWidth="1.4" /></svg>,
    box: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M3 7l7-4 7 4v6l-7 4-7-4V7Z" stroke={color} strokeWidth="1.4" strokeLinejoin="round" /><path d="M10 3v14M3 7l7 4 7-4" stroke={color} strokeWidth="1.4" strokeLinejoin="round" /></svg>,
    chart: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M3 16h14M6 16V10m4 6V6m4 10V8" stroke={color} strokeWidth="1.4" strokeLinecap="round" /></svg>,
    settings: <svg {...s} viewBox="0 0 20 20" fill="none"><circle cx="10" cy="10" r="3" stroke={color} strokeWidth="1.4" /><path d="M10 2v2m0 12v2M2 10h2m12 0h2m-3.17-4.83-1.42 1.42M6.59 13.41l-1.42 1.42m0-9.66 1.42 1.42m4.82 4.82 1.42 1.42" stroke={color} strokeWidth="1.4" strokeLinecap="round" /></svg>,
    search: <svg {...s} viewBox="0 0 20 20" fill="none"><circle cx="9" cy="9" r="5.5" stroke={color} strokeWidth="1.4" /><path d="M16 16l-3-3" stroke={color} strokeWidth="1.4" strokeLinecap="round" /></svg>,
    eye: <svg {...s} viewBox="0 0 20 20" fill="none"><ellipse cx="10" cy="10" rx="7" ry="4" stroke={color} strokeWidth="1.4" /><circle cx="10" cy="10" r="2" stroke={color} strokeWidth="1.4" /></svg>,
    msg: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M3 4h14a1 1 0 0 1 1 1v8a1 1 0 0 1-1 1H6l-4 2V5a1 1 0 0 1 1-1Z" stroke={color} strokeWidth="1.4" strokeLinejoin="round" /></svg>,
    phone: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M5 3h3l1.5 3.5-2 1.5a9 9 0 0 0 4.5 4.5l1.5-2L17 12v3a2 2 0 0 1-2 2A14 14 0 0 1 3 5a2 2 0 0 1 2-2Z" stroke={color} strokeWidth="1.4" strokeLinejoin="round" /></svg>,
    back: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M12 5L7 10l5 5" stroke={color} strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" /></svg>,
    roof: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M2 10L10 2l8 8" stroke={color} strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" /><path d="M5 10v8h10v-8" stroke={color} strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" /></svg>,
    plus: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M10 4v12M4 10h12" stroke={color} strokeWidth="1.6" strokeLinecap="round" /></svg>,
    trash: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M4 6h12M8 6V4h4v2m-7 0 .7 10a1 1 0 0 0 1 .9h4.6a1 1 0 0 0 1-.9L15 6" stroke={color} strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" /></svg>,
    chev: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M6 8l4 4 4-4" stroke={color} strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" /></svg>,
    download: <svg {...s} viewBox="0 0 20 20" fill="none"><path d="M10 3v10m0 0 3.5-3.5M10 13l-3.5-3.5" stroke={color} strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" /><path d="M4 15v1a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-1" stroke={color} strokeWidth="1.4" strokeLinecap="round" /></svg>,
  };
  return icons[name] ?? null;
}

const NAV = [
  { id: 'home', icon: 'home', label: 'Início' },
  { id: 'comp', icon: 'grid', label: 'Componentes' },
  { id: 'mat', icon: 'box', label: 'Materiais' },
  { id: 'rel', icon: 'chart', label: 'Relatórios' },
  { id: 'def', icon: 'settings', label: 'Definições' },
];

/* ─── SIDEBAR ─── */
function Sidebar({ view, onNav, components, onSelectComp, activeComp, search }) {
  const filtered = components.filter(c => c.name.toLowerCase().includes(search.toLowerCase()));
  return (
    <aside className="app-sidebar" style={{ width: 280, flexShrink: 0, background: C.white, borderRight: `1px solid ${C.gray200}`, display: 'flex', flexDirection: 'column', height: '100vh', position: 'fixed', left: 0, top: 0, zIndex: 50 }}>
      <div style={{ padding: '28px 24px 20px', borderBottom: `1px solid ${C.gray200}` }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 34, height: 34, background: C.black, borderRadius: 10, display: 'flex', alignItems: 'center', justifyContent: 'center' }}><Icon name="roof" size={16} color="white" /></div>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, color: C.black, lineHeight: 1.2 }}>Casa Lodge · Bilene</div>
            <div style={{ fontSize: 11, color: C.gray500, lineHeight: 1.2 }}>Acto Studio</div>
          </div>
        </div>
      </div>
      <nav style={{ padding: '16px 12px 8px' }}>
        {NAV.map(item => {
          const active = view === item.id;
          return (
            <button key={item.id} onClick={() => onNav(item.id)} style={{ display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '10px 12px', borderRadius: 12, background: active ? C.black : 'transparent', border: 'none', cursor: 'pointer', marginBottom: 4 }}>
              <Icon name={item.icon} size={17} color={active ? 'white' : C.gray500} />
              <span style={{ fontSize: 13, fontWeight: active ? 600 : 400, color: active ? 'white' : C.gray700 }}>{item.label}</span>
            </button>
          );
        })}
      </nav>
      <div style={{ padding: '0 12px', flex: 1, overflowY: 'auto' }}>
        <div style={{ fontSize: 10, fontWeight: 600, color: C.gray500, textTransform: 'uppercase', letterSpacing: '0.08em', padding: '8px 12px 6px' }}>Componentes</div>
        {filtered.map(c => {
          const st = componentStatus(c); const s = STATUS[st]; const sel = activeComp === c.id;
          return (
            <button key={c.id} onClick={() => onSelectComp(c.id)} style={{ display: 'flex', alignItems: 'center', gap: 8, width: '100%', padding: '7px 12px', borderRadius: 10, background: sel ? s.bg : 'transparent', border: 'none', cursor: 'pointer', textAlign: 'left', marginBottom: 2, borderLeft: sel ? `3px solid ${s.color}` : '3px solid transparent' }}>
              <span style={{ width: 9, height: 9, borderRadius: '50%', background: s.color, flexShrink: 0, boxShadow: `0 0 0 2px ${s.bg}` }} />
              <span style={{ flex: 1, fontSize: 12, color: C.black, fontWeight: sel ? 600 : 400, lineHeight: 1.3, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{c.name}{c.qty > 1 ? ` ×${c.qty}` : ''}</span>
              <span style={{ fontSize: 11, fontWeight: 600, color: s.color, flexShrink: 0 }}>{componentProgress(c)}%</span>
            </button>
          );
        })}
      </div>
      <div style={{ padding: '16px 24px', borderTop: `1px solid ${C.gray200}`, display: 'flex', alignItems: 'center', gap: 10 }}>
        <div style={{ width: 34, height: 34, borderRadius: '50%', background: C.black, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}><span style={{ fontSize: 14, fontWeight: 700, color: 'white' }}>PH</span></div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: C.black, lineHeight: 1.2 }}>Pedro Halar Jr.</div>
          <div style={{ fontSize: 11, color: C.gray500, lineHeight: 1.2 }}>Director de Obra</div>
        </div>
      </div>
    </aside>
  );
}

function MobileTabbar({ view, onNav }) {
  return (
    <nav className="mobile-tabbar" style={{ position: 'fixed', bottom: 0, left: 0, right: 0, background: C.black, display: 'none', justifyContent: 'space-around', padding: '10px 6px', zIndex: 60, borderTopLeftRadius: 20, borderTopRightRadius: 20 }}>
      {NAV.map(item => {
        const active = view === item.id;
        return (
          <button key={item.id} onClick={() => onNav(item.id)} style={{ background: 'none', border: 'none', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 3, padding: '4px 8px', cursor: 'pointer' }}>
            <div style={{ width: 30, height: 30, borderRadius: '50%', background: active ? 'white' : 'transparent', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <Icon name={item.icon} size={16} color={active ? C.black : '#999'} />
            </div>
          </button>
        );
      })}
    </nav>
  );
}

function Topbar({ title, subtitle, onBack, search, setSearch, onExport }) {
  return (
    <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: 28, flexWrap: 'wrap', gap: 16 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
        {onBack && (
          <button onClick={onBack} style={{ background: C.white, border: `1px solid ${C.gray200}`, borderRadius: 10, padding: '8px 12px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, fontSize: 13, color: C.gray700 }}>
            <Icon name="back" size={15} color={C.gray700} /> Voltar
          </button>
        )}
        <div>
          <div style={{ fontSize: 26, fontWeight: 800, color: C.black, lineHeight: 1.1 }}>{title}</div>
          {subtitle && <div style={{ fontSize: 14, color: C.gray500, marginTop: 2 }}>{subtitle}</div>}
        </div>
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        {setSearch && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, background: C.white, borderRadius: 99, padding: '8px 16px', border: `1px solid ${C.gray200}`, width: 240 }}>
            <Icon name="search" size={15} color={C.gray500} />
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Pesquisar fase ou componente" style={{ border: 'none', outline: 'none', fontSize: 13, color: C.black, width: '100%', background: 'transparent' }} />
          </div>
        )}
        {onExport && (
          <button onClick={onExport} title="Exportar planilha" style={{ width: 38, height: 38, borderRadius: '50%', background: C.black, border: 'none', display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer' }}>
            <Icon name="download" size={16} color="white" />
          </button>
        )}
      </div>
    </div>
  );
}

function ComponentCardSmall({ comp, onClick }) {
  const t = componentTotals(comp);
  const over = t.actual > t.budget;
  const statusKey = over ? 'aviso' : componentStatus(comp);
  const accent = STATUS[statusKey].color;
  const pct = componentProgress(comp);
  return (
    <div onClick={onClick} style={{ background: C.white, borderRadius: 18, padding: 18, cursor: 'pointer', boxShadow: '0 2px 12px #00000009', border: `1px solid ${C.gray200}` }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 10, gap: 8 }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.black, lineHeight: 1.25 }}>{comp.name}{comp.qty > 1 ? ` ×${comp.qty}` : ''}</div>
        <Badge status={statusKey} />
      </div>
      <div style={{ fontSize: 11, color: C.gray500, marginBottom: 12 }}>Orçam. MT {fmt(t.budget)} · Gasto MT {fmt(t.actual)}</div>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6 }}>
        <span style={{ fontSize: 11, color: C.gray500 }}>Progresso físico</span>
        <span style={{ fontSize: 11, fontWeight: 700, color: accent }}>{pct}%</span>
      </div>
      <div style={{ height: 6, background: C.gray200, borderRadius: 99, overflow: 'hidden' }}>
        <div style={{ width: `${Math.min(pct, 100)}%`, height: '100%', background: accent, borderRadius: 99, transition: 'width .5s' }} />
      </div>
    </div>
  );
}

function HomeScreen({ state, onSelectComp, onExport, search, setSearch }) {
  const { components } = state;
  const totals = components.map(componentTotals);
  const totalBudget = totals.reduce((s, t) => s + t.budget, 0);
  const totalActual = totals.reduce((s, t) => s + t.actual, 0);
  const globalPct = totalBudget > 0 ? Math.round((totalActual / totalBudget) * 100) : 0;
  const saldo = totalBudget - totalActual;
  const over = totalActual > totalBudget;

  const activePhase = useMemo(() => {
    for (const c of components) for (const p of c.phases) if (p.status === 'em_curso') return { c, p };
    return null;
  }, [components]);

  const recentMaterials = useMemo(() => {
    const all = [];
    components.forEach(c => c.phases.forEach(p => p.materials.forEach(m => all.push({ ...m, phaseLabel: `${p.code} · ${p.label}` }))));
    return all.sort((a, b) => (b.date || '').localeCompare(a.date || '')).slice(0, 5);
  }, [components]);

  const filtered = components.filter(c => c.name.toLowerCase().includes(search.toLowerCase()));

  return (
    <>
      <Topbar title="Olá, Pedro 👋" subtitle="Aqui está o resumo da obra hoje" search={search} setSearch={setSearch} onExport={onExport} />
      <div className="top-cards" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 16, marginBottom: 28 }}>
        <Card black style={{ gridColumn: 'span 2', display: 'flex', flexDirection: 'column', gap: 20 }}>
          <div>
            <div style={{ fontSize: 12, color: '#AAAAAA', fontWeight: 500, textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 4 }}>Orçamento Total (c/ IVA)</div>
            <div style={{ fontSize: 34, fontWeight: 800, color: 'white', letterSpacing: '-0.02em', lineHeight: 1 }}>{fmt(totalBudget)} MT</div>
          </div>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
            <button onClick={onExport} style={{ background: 'white', color: C.black, border: 'none', borderRadius: 99, padding: '9px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Exportar Planilha</button>
            {activePhase && <button onClick={() => onSelectComp(activePhase.c.id)} style={{ background: 'transparent', color: 'white', border: '1.5px solid #FFFFFF40', borderRadius: 99, padding: '9px 20px', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}>Ver Fase em Curso</button>}
          </div>
        </Card>
        <Card style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          <div>
            <div style={{ fontSize: 11, color: C.gray500, fontWeight: 500, textTransform: 'uppercase', letterSpacing: '0.07em', marginBottom: 4 }}>Gasto Real</div>
            <div style={{ fontSize: 24, fontWeight: 700, color: over ? C.terra : C.black, letterSpacing: '-0.02em', lineHeight: 1.1 }}>{fmt(totalActual)} MT</div>
            <div style={{ fontSize: 12, color: C.gray500, marginTop: 4 }}>Saldo: <span style={{ color: saldo >= 0 ? C.lagoon : C.terra, fontWeight: 600 }}>{saldo >= 0 ? '+' : ''}{fmt(saldo)} MT</span></div>
          </div>
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6 }}>
              <span style={{ fontSize: 11, color: C.gray500 }}>Execução financeira</span>
              <span style={{ fontSize: 11, fontWeight: 700, color: over ? C.terra : C.lagoon }}>{globalPct}%</span>
            </div>
            <TapeBar pct={globalPct} over={over} />
          </div>
        </Card>
      </div>

      {activePhase ? (
        <Card style={{ marginBottom: 28, padding: 24 }}>
          <div style={{ fontSize: 11, fontWeight: 600, color: C.gray500, textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 16 }}>Fase em Curso</div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 32, flexWrap: 'wrap' }}>
            <div style={{ flex: '0 0 auto', minWidth: 200 }}>
              <div style={{ fontSize: 19, fontWeight: 700, color: C.black, lineHeight: 1.2 }}>{activePhase.p.code} · {activePhase.p.label}</div>
              <div style={{ fontSize: 13, color: C.gray500, marginTop: 4, marginBottom: 12 }}>Componente: {activePhase.c.name}</div>
              <Badge status="em_curso" />
            </div>
            <div style={{ flex: 1, minWidth: 240 }}><Timeline steps={['Não iniciada', 'Em curso', 'Concluída']} active={1} /></div>
            <div style={{ flex: '0 0 auto' }}>
              <button onClick={() => onSelectComp(activePhase.c.id)} style={{ background: C.black, color: 'white', border: 'none', borderRadius: 12, padding: '11px 22px', fontSize: 13, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}>Actualizar Progresso</button>
            </div>
          </div>
        </Card>
      ) : (
        <Card style={{ marginBottom: 28, padding: 24, textAlign: 'center', color: C.gray500, fontSize: 13 }}>Nenhuma fase está actualmente "Em curso" — marca uma fase para a veres em destaque aqui.</Card>
      )}

      <div className="home-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 320px', gap: 24, alignItems: 'start' }}>
        <div>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 14 }}><div style={{ fontSize: 16, fontWeight: 700, color: C.black }}>Componentes</div></div>
          <div className="comp-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 14 }}>
            {filtered.map(c => <ComponentCardSmall key={c.id} comp={c} onClick={() => onSelectComp(c.id)} />)}
          </div>
        </div>
        <div>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 14 }}><div style={{ fontSize: 16, fontWeight: 700, color: C.black }}>Registos Recentes</div></div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            {recentMaterials.length === 0 && <Card style={{ padding: 16, fontSize: 12, color: C.gray500 }}>Ainda não há materiais registados.</Card>}
            {recentMaterials.map(r => (
              <Card key={r.id} style={{ padding: '14px 16px', display: 'flex', flexDirection: 'column', gap: 6 }}>
                <div style={{ fontSize: 13, fontWeight: 500, color: C.black, lineHeight: 1.3 }}>{r.name} {r.qty ? `(${r.qty} ${r.unit || 'un'})` : ''}</div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontSize: 11, color: C.gray500 }}>{r.phaseLabel}</span>
                  <span style={{ fontSize: 10, fontWeight: 600, color: C.lagoon, background: '#1F7A6C18', padding: '2px 8px', borderRadius: 99 }}>Registado</span>
                </div>
                <div style={{ fontSize: 11, color: C.gray500 }}>{r.date || '—'}</div>
              </Card>
            ))}
          </div>
        </div>
      </div>
    </>
  );
}

function ComponentsScreen({ state, onSelectComp, onAddComponent, search, setSearch, onExport }) {
  const filtered = state.components.filter(c => c.name.toLowerCase().includes(search.toLowerCase()));
  return (
    <>
      <Topbar title="Componentes" subtitle={`${state.components.length} componentes do projecto`} search={search} setSearch={setSearch} onExport={onExport} />
      <div className="comp-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 14, marginBottom: 20 }}>
        {filtered.map(c => <ComponentCardSmall key={c.id} comp={c} onClick={() => onSelectComp(c.id)} />)}
        <button onClick={onAddComponent} style={{ border: `1.5px dashed ${C.gray300}`, borderRadius: 18, background: 'transparent', color: C.gray500, fontWeight: 600, fontSize: 13, cursor: 'pointer', minHeight: 120, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8 }}>
          <Icon name="plus" size={16} color={C.gray500} /> Adicionar componente
        </button>
      </div>
    </>
  );
}

function DetailScreen({ comp, onBack, onPatchComponent, onPatchPhase, onAddMaterial, onDeleteMaterial, onAddPhase, onDeletePhase, onDeleteComponent }) {
  const [openPhase, setOpenPhase] = useState(comp.phases[0]?.id ?? null);
  const t = componentTotals(comp);
  const over = t.actual > t.budget;
  return (
    <>
      <Topbar title={comp.name} subtitle={`${comp.phases.length} fases · Quantidade ${comp.qty}`} onBack={onBack} />
      <div className="detail-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 340px', gap: 24, alignItems: 'start' }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
          <Card>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12, marginBottom: 8 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <span style={{ fontSize: 12, color: C.gray500, fontWeight: 600 }}>Quantidade de unidades</span>
                <input type="number" min={1} value={comp.qty} onChange={e => onPatchComponent({ qty: Math.max(1, parseInt(e.target.value) || 1) })} style={{ width: 64, padding: '6px 8px', borderRadius: 8, border: `1px solid ${C.gray200}`, fontSize: 13 }} />
              </div>
              <button onClick={onDeleteComponent} style={{ background: 'none', border: 'none', color: C.terra, fontSize: 12, fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4 }}><Icon name="trash" size={13} color={C.terra} /> Remover componente</button>
            </div>
            <div style={{ fontSize: 12, color: C.gray500 }}>Orçam. total <b style={{ color: C.black }}>MT {fmt(t.budget)}</b> · Gasto real <b style={{ color: over ? C.terra : C.lagoon }}>MT {fmt(t.actual)}</b></div>
          </Card>
          {comp.phases.map(p => (
            <PhaseCard key={p.id} phase={p} qty={comp.qty} open={openPhase === p.id}
              onToggle={() => setOpenPhase(openPhase === p.id ? null : p.id)}
              onPatch={patch => onPatchPhase(p.id, patch)}
              onAddMaterial={mat => onAddMaterial(p.id, mat)}
              onDeleteMaterial={mid => onDeleteMaterial(p.id, mid)}
              onDelete={() => onDeletePhase(p.id)} />
          ))}
          <button onClick={onAddPhase} style={{ border: `1.5px dashed ${C.gray300}`, borderRadius: 16, background: 'transparent', color: C.gray500, fontWeight: 600, fontSize: 13, cursor: 'pointer', padding: '14px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8 }}>
            <Icon name="plus" size={16} color={C.gray500} /> Adicionar fase personalizada
          </button>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
          <Card>
            <div style={{ fontSize: 13, fontWeight: 600, color: C.black, marginBottom: 14 }}>Responsável da Obra</div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: C.black, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}><span style={{ fontSize: 18, fontWeight: 700, color: 'white' }}>PH</span></div>
              <div>
                <div style={{ fontSize: 15, fontWeight: 600, color: C.black, lineHeight: 1.2 }}>Pedro Halar Jr.</div>
                <div style={{ fontSize: 12, color: C.gray500 }}>Director de Obra</div>
              </div>
            </div>
            <div style={{ display: 'flex', gap: 10 }}>
              <button style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: C.gray100, border: 'none', borderRadius: 10, padding: '9px', cursor: 'pointer', fontSize: 12, color: C.gray700 }}><Icon name="msg" size={15} color={C.gray700} /> Mensagem</button>
              <button style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: C.gray100, border: 'none', borderRadius: 10, padding: '9px', cursor: 'pointer', fontSize: 12, color: C.gray700 }}><Icon name="phone" size={15} color={C.gray700} /> Chamada</button>
            </div>
          </Card>
          <Card style={{ background: '#F5F8F7' }}>
            <div style={{ fontSize: 13, fontWeight: 600, color: C.black, marginBottom: 16 }}>Resumo Financeiro</div>
            {[
              { label: 'Subtotal (s/IVA)', val: `MT ${fmt(t.budgetSub)}` },
              { label: 'IVA 16%', val: `MT ${fmt(t.iva)}` },
              { label: 'Orçamento total', val: `MT ${fmt(t.budget)}`, bold: true },
              { label: 'Gasto real', val: `MT ${fmt(t.actual)}`, bold: true, color: over ? C.terra : C.lagoon },
              { label: 'Saldo', val: `${t.budget - t.actual >= 0 ? '+' : ''}MT ${fmt(t.budget - t.actual)}`, bold: true, color: t.budget - t.actual >= 0 ? C.lagoon : C.terra },
            ].map(row => (
              <div key={row.label} style={{ display: 'flex', justifyContent: 'space-between', padding: '10px 0', borderBottom: `1px solid ${C.gray200}` }}>
                <span style={{ fontSize: 13, color: C.gray500 }}>{row.label}</span>
                <span style={{ fontSize: 13, fontWeight: row.bold ? 600 : 400, color: row.color ?? C.black }}>{row.val}</span>
              </div>
            ))}
            {over && (
              <div style={{ marginTop: 14, background: '#B5502E12', border: '1px solid #B5502E30', borderRadius: 10, padding: '10px 14px', display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ fontSize: 16 }}>⚠️</span><span style={{ fontSize: 12, color: C.terra, fontWeight: 500 }}>Acima do orçamento em MT {fmt(t.actual - t.budget)}</span>
              </div>
            )}
          </Card>
        </div>
      </div>
    </>
  );
}

function PhaseCard({ phase: p, qty, open, onToggle, onPatch, onAddMaterial, onDeleteMaterial, onDelete }) {
  const budget = phaseBudget(p, qty);
  const actual = phaseActual(p);
  const over = actual > budget;
  const pct = budget > 0 ? Math.round((actual / budget) * 100) : (actual > 0 ? 999 : 0);
  const [form, setForm] = useState({ name: '', qty: 1, unit: '', unitCost: 0, supplier: '', date: '' });

  const submitMaterial = () => {
    if (!form.name.trim()) return;
    onAddMaterial({ id: uid('m'), name: form.name.trim(), qty: Number(form.qty) || 0, unit: form.unit || 'un', unitCost: Number(form.unitCost) || 0, supplier: form.supplier, date: form.date });
    setForm({ name: '', qty: 1, unit: '', unitCost: 0, supplier: '', date: '' });
  };

  return (
    <Card style={{ padding: 0, overflow: 'hidden' }}>
      <div onClick={onToggle} style={{ display: 'flex', alignItems: 'center', gap: 14, padding: '16px 20px', cursor: 'pointer', flexWrap: 'wrap' }}>
        <div style={{ fontSize: 18, fontWeight: 800, color: C.gray300, width: 46, flexShrink: 0 }}>{p.code}</div>
        <div style={{ flex: 1, minWidth: 160 }}><div style={{ fontSize: 14, fontWeight: 700, color: C.black, lineHeight: 1.25 }}>{p.label}</div></div>
        <Badge status={over ? 'aviso' : p.status} />
        <div style={{ textAlign: 'right', fontSize: 12, minWidth: 110 }}>
          <div style={{ color: C.gray500 }}>Orçam. {fmt(budget)}</div>
          <div style={{ fontWeight: 700, color: over ? C.terra : C.black }}>{fmt(actual)}</div>
        </div>
        <Icon name="chev" size={16} color={C.gray500} />
      </div>
      <div style={{ padding: '0 20px 4px' }}><TapeBar pct={Math.min(pct, 100)} over={over} /></div>
      {open && (
        <div style={{ padding: 20, borderTop: `1px solid ${C.gray200}`, marginTop: 12 }}>
          <div style={{ display: 'flex', gap: 14, flexWrap: 'wrap', marginBottom: 18 }}>
            <Field label="Estado">
              <select value={p.status} onChange={e => onPatch({ status: e.target.value })} style={inputStyle}>
                <option value="nao_iniciado">Não iniciada</option>
                <option value="em_curso">Em curso</option>
                <option value="concluido">Concluída</option>
              </select>
            </Field>
            <Field label={`Progresso físico: ${p.progress}%`} wide>
              <input type="range" min={0} max={100} value={p.progress} onChange={e => onPatch({ progress: Number(e.target.value) })} style={{ width: '100%' }} />
            </Field>
            <Field label="Orçam. por unidade (MT)"><input type="number" min={0} value={p.unitBudget} onChange={e => onPatch({ unitBudget: Number(e.target.value) || 0 })} style={inputStyle} /></Field>
            <Field label="Mão-de-obra / outros (MT)"><input type="number" min={0} value={p.labor} onChange={e => onPatch({ labor: Number(e.target.value) || 0 })} style={inputStyle} /></Field>
          </div>
          <div style={{ overflowX: 'auto', marginBottom: 14 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
              <thead><tr>{['Material', 'Qtd.', 'Preço unit.', 'Total', 'Fornecedor', 'Data', ''].map(h => <th key={h} style={{ textAlign: 'left', fontSize: 10, textTransform: 'uppercase', letterSpacing: '.05em', color: C.gray500, borderBottom: `1px solid ${C.gray200}`, padding: '6px 8px' }}>{h}</th>)}</tr></thead>
              <tbody>
                {p.materials.length === 0 && <tr><td colSpan={7} style={{ textAlign: 'center', color: C.gray300, fontStyle: 'italic', padding: 16 }}>Sem materiais registados nesta fase.</td></tr>}
                {p.materials.map(m => (
                  <tr key={m.id}>
                    <td style={{ padding: '7px 8px', borderBottom: `1px solid ${C.gray100}` }}>{m.name}</td>
                    <td style={{ padding: '7px 8px', borderBottom: `1px solid ${C.gray100}`, textAlign: 'right' }}>{m.qty} {m.unit}</td>
                    <td style={{ padding: '7px 8px', borderBottom: `1px solid ${C.gray100}`, textAlign: 'right' }}>{fmt(m.unitCost)}</td>
                    <td style={{ padding: '7px 8px', borderBottom: `1px solid ${C.gray100}`, textAlign: 'right', fontWeight: 600 }}>{fmt(m.qty * m.unitCost)}</td>
                    <td style={{ padding: '7px 8px', borderBottom: `1px solid ${C.gray100}` }}>{m.supplier || '—'}</td>
                    <td style={{ padding: '7px 8px', borderBottom: `1px solid ${C.gray100}` }}>{m.date || '—'}</td>
                    <td style={{ padding: '7px 8px', borderBottom: `1px solid ${C.gray100}` }}><button onClick={() => onDeleteMaterial(m.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: C.gray500 }}><Icon name="trash" size={14} color={C.gray500} /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'flex-end', background: C.gray100, padding: 12, borderRadius: 12 }}>
            <Field label="Material"><input value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="Ex: Cimento 50kg" style={inputStyle} /></Field>
            <Field label="Qtd." small><input type="number" value={form.qty} onChange={e => setForm({ ...form, qty: e.target.value })} style={inputStyle} /></Field>
            <Field label="Unid." small><input value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })} placeholder="saco" style={inputStyle} /></Field>
            <Field label="Preço unit."><input type="number" value={form.unitCost} onChange={e => setForm({ ...form, unitCost: e.target.value })} style={inputStyle} /></Field>
            <Field label="Fornecedor"><input value={form.supplier} onChange={e => setForm({ ...form, supplier: e.target.value })} placeholder="opcional" style={inputStyle} /></Field>
            <Field label="Data"><input type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} style={inputStyle} /></Field>
            <button onClick={submitMaterial} style={{ background: C.black, color: 'white', border: 'none', borderRadius: 10, padding: '9px 16px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>+ Registar</button>
          </div>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 14, flexWrap: 'wrap', gap: 8 }}>
            <span style={{ fontSize: 11, color: C.gray500 }}>Materiais: MT {fmt(materialsTotal(p))} · Mão-de-obra: MT {fmt(p.labor || 0)}</span>
            <button onClick={onDelete} style={{ background: 'none', border: 'none', color: C.terra, fontSize: 12, fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4 }}><Icon name="trash" size={13} color={C.terra} /> Remover fase</button>
          </div>
        </div>
      )}
    </Card>
  );
}
const inputStyle = { padding: '8px 9px', border: `1px solid ${C.gray200}`, borderRadius: 8, fontSize: 13, width: '100%' };
function Field({ label, children, small, wide }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 5, minWidth: small ? 70 : wide ? 200 : 130, flex: wide ? 1 : 'initial' }}>
      <label style={{ fontSize: 10, fontWeight: 600, color: C.gray500, textTransform: 'uppercase', letterSpacing: '.04em' }}>{label}</label>
      {children}
    </div>
  );
}

function MaterialsScreen({ state, search, setSearch, onExport }) {
  const rows = [];
  state.components.forEach(c => c.phases.forEach(p => p.materials.forEach(m => rows.push({ ...m, comp: c.name, phase: `${p.code} · ${p.label}` }))));
  const filtered = rows.filter(r => (r.name + r.comp + r.phase).toLowerCase().includes(search.toLowerCase()));
  const total = filtered.reduce((s, m) => s + m.qty * m.unitCost, 0);
  return (
    <>
      <Topbar title="Materiais" subtitle={`${rows.length} registos em todo o projecto`} search={search} setSearch={setSearch} onExport={onExport} />
      <Card style={{ overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead><tr>{['Material', 'Componente', 'Fase', 'Qtd.', 'Preço unit.', 'Total', 'Fornecedor', 'Data'].map(h => <th key={h} style={{ textAlign: 'left', fontSize: 10, textTransform: 'uppercase', letterSpacing: '.05em', color: C.gray500, borderBottom: `1px solid ${C.gray200}`, padding: '8px' }}>{h}</th>)}</tr></thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={8} style={{ textAlign: 'center', color: C.gray300, fontStyle: 'italic', padding: 20 }}>Sem materiais registados.</td></tr>}
            {filtered.map(m => (
              <tr key={m.id}>
                <td style={{ padding: 8, borderBottom: `1px solid ${C.gray100}` }}>{m.name}</td>
                <td style={{ padding: 8, borderBottom: `1px solid ${C.gray100}`, color: C.gray500 }}>{m.comp}</td>
                <td style={{ padding: 8, borderBottom: `1px solid ${C.gray100}`, color: C.gray500 }}>{m.phase}</td>
                <td style={{ padding: 8, borderBottom: `1px solid ${C.gray100}`, textAlign: 'right' }}>{m.qty} {m.unit}</td>
                <td style={{ padding: 8, borderBottom: `1px solid ${C.gray100}`, textAlign: 'right' }}>{fmt(m.unitCost)}</td>
                <td style={{ padding: 8, borderBottom: `1px solid ${C.gray100}`, textAlign: 'right', fontWeight: 600 }}>{fmt(m.qty * m.unitCost)}</td>
                <td style={{ padding: 8, borderBottom: `1px solid ${C.gray100}` }}>{m.supplier || '—'}</td>
                <td style={{ padding: 8, borderBottom: `1px solid ${C.gray100}` }}>{m.date || '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {filtered.length > 0 && <div style={{ textAlign: 'right', marginTop: 14, fontSize: 13, fontWeight: 700 }}>Total: MT {fmt(total)}</div>}
      </Card>
    </>
  );
}

function ReportsScreen({ state, onExport }) {
  const totals = state.components.map(c => ({ c, t: componentTotals(c) }));
  const g = totals.reduce((s, { t }) => ({ budgetSub: s.budgetSub + t.budgetSub, iva: s.iva + t.iva, budget: s.budget + t.budget, actual: s.actual + t.actual }), { budgetSub: 0, iva: 0, budget: 0, actual: 0 });
  return (
    <>
      <Topbar title="Relatórios" subtitle="Orçamento vs. gasto real por componente" onExport={onExport} />
      <Card style={{ marginBottom: 20 }}>
        {totals.map(({ c, t }) => {
          const max = Math.max(t.budget, t.actual, 1);
          const pct = Math.min(100, (t.actual / max) * 100);
          const over = t.actual > t.budget;
          return (
            <div key={c.id} style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 14 }}>
              <div style={{ width: 170, flexShrink: 0, fontSize: 12, fontWeight: 600, color: C.black }}>{c.name}</div>
              <div style={{ flex: 1, height: 14, background: C.gray200, borderRadius: 8, position: 'relative', overflow: 'hidden' }}>
                <div style={{ position: 'absolute', top: 2, left: 0, height: 10, width: `${pct}%`, background: over ? C.terra : C.lagoon, borderRadius: 6 }} />
              </div>
              <div style={{ width: 190, flexShrink: 0, fontSize: 11, color: C.gray500, textAlign: 'right' }}>MT {fmt(t.actual)} / {fmt(t.budget)}</div>
            </div>
          );
        })}
      </Card>
      <Card>
        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 14 }}>Resumo financeiro global</div>
        {[['Subtotal (s/ IVA)', g.budgetSub], ['IVA (16%)', g.iva], ['Orçamento total', g.budget], ['Gasto real acumulado', g.actual], ['Saldo disponível', g.budget - g.actual]].map(([label, val]) => (
          <div key={label} style={{ display: 'flex', justifyContent: 'space-between', padding: '9px 0', borderBottom: `1px solid ${C.gray200}`, fontSize: 13 }}>
            <span style={{ color: C.gray500 }}>{label}</span><span style={{ fontWeight: 600 }}>MT {fmt(val)}</span>
          </div>
        ))}
      </Card>
    </>
  );
}

function SettingsScreen({ onReset }) {
  return (
    <>
      <Topbar title="Definições" subtitle="Dados e preferências do projecto" />
      <Card style={{ maxWidth: 480 }}>
        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>Dados guardados automaticamente</div>
        <div style={{ fontSize: 12, color: C.gray500, marginBottom: 18 }}>Todas as alterações (fases, materiais, custos) ficam guardadas neste dispositivo, mesmo offline.</div>
        <button onClick={onReset} style={{ background: 'none', border: `1px solid ${C.terra}`, color: C.terra, borderRadius: 10, padding: '9px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Repor dados originais da proposta</button>
      </Card>
    </>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(17,17,17,.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100, padding: 20 }}>
      <div onClick={e => e.stopPropagation()} style={{ background: 'white', borderRadius: 20, padding: 24, width: '100%', maxWidth: 380 }}>
        <div style={{ fontSize: 17, fontWeight: 700, marginBottom: 16 }}>{title}</div>
        {children}
      </div>
    </div>
  );
}

function AddComponentForm({ onCancel, onSave }) {
  const [name, setName] = useState('');
  const [qty, setQty] = useState(1);
  return (
    <div>
      <Field label="Nome"><input value={name} onChange={e => setName(e.target.value)} placeholder="Ex: Muro de vedação" style={inputStyle} /></Field>
      <div style={{ height: 12 }} />
      <Field label="Quantidade"><input type="number" min={1} value={qty} onChange={e => setQty(Number(e.target.value) || 1)} style={inputStyle} /></Field>
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 20 }}>
        <button onClick={onCancel} style={{ background: 'none', border: `1px solid ${C.gray200}`, borderRadius: 10, padding: '9px 16px', fontSize: 13, cursor: 'pointer' }}>Cancelar</button>
        <button onClick={() => name.trim() && onSave(name.trim(), qty)} style={{ background: C.black, color: 'white', border: 'none', borderRadius: 10, padding: '9px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Adicionar</button>
      </div>
    </div>
  );
}

/* ─── APP RAIZ ─── */
function App() {
  const [state, setState] = useState(loadInitial);
  const [view, setView] = useState('home');
  const [activeComp, setActiveComp] = useState(null);
  const [search, setSearch] = useState('');
  const [modal, setModal] = useState(null);
  const [saveNote, setSaveNote] = useState('');
  const [installPrompt, setInstallPrompt] = useState(null);

  useEffect(() => {
    const handler = (e) => { e.preventDefault(); setInstallPrompt(e); };
    window.addEventListener('beforeinstallprompt', handler);
    return () => window.removeEventListener('beforeinstallprompt', handler);
  }, []);

  const persist = useCallback((next) => {
    setState(next);
    const ok = saveState(next);
    setSaveNote(ok ? 'Guardado ✓' : 'Falha ao guardar');
    setTimeout(() => setSaveNote(''), 1200);
  }, []);

  const logActivity = (list, desc) => [{ id: uid('a'), time: nowLabel(), desc }, ...list].slice(0, 40);
  const comp = state.components.find(c => c.id === activeComp);

  const patchComponent = (patch) => persist({ ...state, components: state.components.map(c => c.id === comp.id ? { ...c, ...patch } : c) });
  const patchPhase = (phaseId, patch) => persist({ ...state, components: state.components.map(c => c.id !== comp.id ? c : { ...c, phases: c.phases.map(p => p.id === phaseId ? { ...p, ...patch } : p) }) });
  const addMaterial = (phaseId, mat) => {
    const phase = comp.phases.find(p => p.id === phaseId);
    persist({
      ...state,
      components: state.components.map(c => c.id !== comp.id ? c : { ...c, phases: c.phases.map(p => p.id === phaseId ? { ...p, materials: [...p.materials, mat] } : p) }),
      activities: logActivity(state.activities, `Registado: ${mat.name} · ${mat.qty} ${mat.unit} (${phase?.code})`),
    });
  };
  const deleteMaterial = (phaseId, matId) => persist({ ...state, components: state.components.map(c => c.id !== comp.id ? c : { ...c, phases: c.phases.map(p => p.id === phaseId ? { ...p, materials: p.materials.filter(m => m.id !== matId) } : p) }) });
  const addPhase = () => {
    const label = prompt('Designação da nova fase:');
    if (!label) return;
    const code = prompt('Código da fase (ex: BoQ5):', `BoQ${comp.phases.length}`) || `BoQ${comp.phases.length}`;
    const unitBudget = Number(prompt('Orçamento por unidade (MT):', '0')) || 0;
    const newPhase = { id: uid('ph'), code, label, unitBudget, status: 'nao_iniciado', progress: 0, labor: 0, materials: [] };
    persist({ ...state, components: state.components.map(c => c.id !== comp.id ? c : { ...c, phases: [...c.phases, newPhase] }), activities: logActivity(state.activities, `Fase "${label}" criada em ${comp.name}`) });
  };
  const deletePhase = (phaseId) => {
    if (!confirm('Remover esta fase e todos os seus dados?')) return;
    persist({ ...state, components: state.components.map(c => c.id !== comp.id ? c : { ...c, phases: c.phases.filter(p => p.id !== phaseId) }) });
  };
  const deleteComponent = () => {
    if (!confirm(`Remover "${comp.name}" e todos os seus dados?`)) return;
    setActiveComp(null); setView('comp');
    persist({ ...state, components: state.components.filter(c => c.id !== comp.id) });
  };
  const addComponent = (name, qty) => {
    const newComp = { id: uid(name.toLowerCase().replace(/[^a-z0-9]+/g, '-')), name, qty, phases: [] };
    persist({ ...state, components: [...state.components, newComp], activities: logActivity(state.activities, `Componente "${name}" adicionado`) });
    setModal(null); setActiveComp(newComp.id); setView('detail');
  };
  const resetAll = () => {
    if (!confirm('Repor todos os dados para os valores originais da proposta? Perdes o progresso e materiais registados.')) return;
    persist(freshState());
  };

  const exportSpreadsheet = () => {
    const wb = XLSX.utils.book_new();
    const totals = state.components.map(componentTotals);
    const g = totals.reduce((s, t) => ({ budgetSub: s.budgetSub + t.budgetSub, iva: s.iva + t.iva, budget: s.budget + t.budget, actual: s.actual + t.actual }), { budgetSub: 0, iva: 0, budget: 0, actual: 0 });

    const resumoRows = [['Componente', 'Quantidade', 'Subtotal (s/IVA)', 'IVA (16%)', 'Orçamento total', 'Gasto real', 'Saldo', 'Progresso físico (%)']];
    state.components.forEach(c => { const t = componentTotals(c); resumoRows.push([c.name, c.qty, t.budgetSub, t.iva, t.budget, t.actual, t.budget - t.actual, componentProgress(c)]); });
    resumoRows.push(['TOTAL GERAL', '', g.budgetSub, g.iva, g.budget, g.actual, g.budget - g.actual, '']);
    const wsResumo = XLSX.utils.aoa_to_sheet(resumoRows);
    wsResumo['!cols'] = [{ wch: 34 }, { wch: 10 }, { wch: 16 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, wsResumo, 'Resumo');

    const faseRows = [['Componente', 'Código', 'Fase', 'Estado', 'Progresso (%)', 'Orçam. unitário', 'Orçam. total', 'Mão-de-obra/outros', 'Materiais', 'Gasto real', 'Saldo']];
    state.components.forEach(c => c.phases.forEach(p => {
      const budget = phaseBudget(p, c.qty); const matTotal = materialsTotal(p); const actual = phaseActual(p);
      faseRows.push([c.name, p.code, p.label, STATUS[p.status].label, p.progress, p.unitBudget, budget, p.labor || 0, matTotal, actual, budget - actual]);
    }));
    const wsFases = XLSX.utils.aoa_to_sheet(faseRows);
    wsFases['!cols'] = [{ wch: 30 }, { wch: 8 }, { wch: 34 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, wsFases, 'Fases');

    const matRows = [['Componente', 'Fase', 'Material', 'Quantidade', 'Unidade', 'Preço unitário', 'Total', 'Fornecedor', 'Data']];
    state.components.forEach(c => c.phases.forEach(p => p.materials.forEach(m => matRows.push([c.name, p.label, m.name, m.qty, m.unit, m.unitCost, m.qty * m.unitCost, m.supplier || '', m.date || '']))));
    if (matRows.length === 1) matRows.push(['Sem materiais registados', '', '', '', '', '', '', '', '']);
    const wsMat = XLSX.utils.aoa_to_sheet(matRows);
    wsMat['!cols'] = [{ wch: 30 }, { wch: 30 }, { wch: 24 }, { wch: 10 }, { wch: 8 }, { wch: 14 }, { wch: 14 }, { wch: 20 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, wsMat, 'Materiais');

    XLSX.writeFile(wb, `Casa-Lodge-Bilene-Controlo-Obra-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const selectComp = (id) => { setActiveComp(id); setView('detail'); };
  const goNav = (id) => { setView(id); if (id !== 'detail') setActiveComp(null); };

  let screen = null;
  if (view === 'home') screen = <HomeScreen state={state} onSelectComp={selectComp} onExport={exportSpreadsheet} search={search} setSearch={setSearch} />;
  else if (view === 'comp') screen = <ComponentsScreen state={state} onSelectComp={selectComp} onAddComponent={() => setModal('addComponent')} search={search} setSearch={setSearch} onExport={exportSpreadsheet} />;
  else if (view === 'mat') screen = <MaterialsScreen state={state} search={search} setSearch={setSearch} onExport={exportSpreadsheet} />;
  else if (view === 'rel') screen = <ReportsScreen state={state} onExport={exportSpreadsheet} />;
  else if (view === 'def') screen = <SettingsScreen onReset={resetAll} />;
  else if (view === 'detail' && comp) {
    screen = <DetailScreen comp={comp} onBack={() => { setView('comp'); setActiveComp(null); }}
      onPatchComponent={patchComponent} onPatchPhase={patchPhase}
      onAddMaterial={addMaterial} onDeleteMaterial={deleteMaterial}
      onAddPhase={addPhase} onDeletePhase={deletePhase} onDeleteComponent={deleteComponent} />;
  } else if (view === 'detail' && !comp) {
    screen = null;
    setTimeout(() => setView('home'), 0);
  }

  return (
    <div style={{ display: 'flex', minHeight: '100vh', background: C.bg, fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, sans-serif" }}>
      <Sidebar view={view === 'detail' ? 'comp' : view} onNav={goNav} components={state.components} onSelectComp={selectComp} activeComp={activeComp} search={search} />
      <MobileTabbar view={view === 'detail' ? 'comp' : view} onNav={goNav} />
      <main className="app-main" style={{ marginLeft: 280, flex: 1, padding: '32px 36px', minWidth: 0 }}>
        {installPrompt && (
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#111', color: 'white', borderRadius: 14, padding: '10px 16px', marginBottom: 18, fontSize: 12 }}>
            <span>Instala esta app no teu telemóvel ou computador para acesso rápido e offline.</span>
            <button onClick={async () => { installPrompt.prompt(); await installPrompt.userChoice; setInstallPrompt(null); }} style={{ background: 'white', color: '#111', border: 'none', borderRadius: 99, padding: '6px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer', flexShrink: 0, marginLeft: 12 }}>Instalar</button>
          </div>
        )}
        {screen}
      </main>
      {saveNote && <div style={{ position: 'fixed', bottom: 20, right: 20, background: C.black, color: 'white', padding: '8px 14px', borderRadius: 10, fontSize: 12, zIndex: 200 }}>{saveNote}</div>}
      {modal === 'addComponent' && <Modal title="Novo componente" onClose={() => setModal(null)}><AddComponentForm onCancel={() => setModal(null)} onSave={addComponent} /></Modal>}
    </div>
  );
}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
