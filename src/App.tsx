import { useState, useMemo, Component, type ReactNode } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, CheckCircle, XCircle, AlertCircle, Play, Download, BarChart2, PieChart, ShieldCheck, TrendingUp, X, Settings, Shield } from 'lucide-react';
import './index.css';

// FIX-1 (Vuln 5): Expand merged cell ranges in-place before parsing.
// xlsx stores merged values only in the top-left cell of the range; all
// other cells of the merge are empty by default. Expanding them ensures
// that Tesouro Gerencial exports with merged account-code headers are
// parsed correctly by both the column detector and the data extractor.
const expandMergedCells = (ws: XLSX.WorkSheet): void => {
  const merges = ws['!merges'] || [];
  for (const merge of merges) {
    const originAddr = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
    const originCell = ws[originAddr];
    if (!originCell) continue;
    for (let r = merge.s.r; r <= merge.e.r; r++) {
      for (let c = merge.s.c; c <= merge.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (addr !== originAddr) ws[addr] = { ...originCell };
      }
    }
  }
};

// RITO ORDINÁRIO — Tabela de Validação (Situação Transferegov → Conjunto de Contas SIAFI Válidas)
// Regra: se TODAS as contas detectadas pertencem ao conjunto → "Correto".
// Se QUALQUER conta está fora do conjunto → "Inconsistência (Rito Patológico)".
// Arrays vazios são sentinelas PENDENTES → "Regra Pendente - Revisar" (nunca "Inconsistência").
const validationRules: Record<string, string[]> = {
  // Celebração
  "Proposta de Plano de Trabalho Aprovado":        ["712210101"],
  // Execução — aceita múltiplas contas simultâneas (saldo a liberar + executado + a repassar)
  // 812210202 incluída para instrumentos OBTV onde recursos já foram liberados ao convenente
  // (A COMPROVAR ativo no SIAFI enquanto TG ainda registra "Em execução" — estado transitório normal)
  "Em execução":                                   ["811210100", "812210101", "812210201", "812210202", "811210103"],
  // Aguardando P.C. — foco na obrigação de comprovar o recurso recebido
  "Aguardando prestação de contas":                ["812210202", "811210102"],
  // P.C. em Análise — transição entre entrega dos documentos e o parecer técnico
  // 812210101 e 812210102 são rastros históricos do ciclo (firmamento + aguardando PC)
  // presentes no export TG com saldo zero mas histórico de movimentação (ex: 823645)
  "Prestação de Contas em Análise":                ["812210101", "812210102", "812210103", "812210104", "812210202", "811210102"],
  // P.C. em Complementação — aguardando saneamento de dúvidas técnicas
  "Prestação de Contas em Complementação":         ["812210103"],
  // Demais estágios
  "Cancelado":                                     ["812210108", "811210106"],
  "Convênio Anulado":                              ["812210109"],
  "Instrumento Anulado":                           ["812210101", "812210109"],
  "Convenio Rescindido":                           ["811210109"],
  "Inadimplente":                                  ["812210106"],
  "Prestação de Contas Aprovada":                  ["812210104"],
  "Prestação de Contas Aprovada com Ressalvas":    ["812210104"],
  "Prestação de Contas Comprovada em Análise":     ["812210103", "812210105", "812210107"],
  "Prestação de Contas Concluída":                 ["812210111", "811210110"],
  "Prestação de Contas Iniciada por Antecipação":  [],
  "Prestação de Contas Rejeitada":                 ["812210106"],
};

// SKILL §6 — Dicionário de Tradução Oficial (Tabela de Referência CSV/Imagens)
// Mapeamento De-Para com descrições literais. NÃO alterar sem aprovação da equipe fiscal.
const accountMap: Record<string, string> = {
  '811210100': 'Execução de Convênios e Instrumentos Congêneres',
  '811210102': 'Convênios e Instrumentos Congêneres a Comprovar',
  '811210103': 'Convênios e Instrumentos Congêneres a Receber',
  '811210106': 'Convênios e Instrumentos Congêneres não recebidos',
  '811210109': 'Convênios e Instrumentos Congêneres a Anular',
  '811210110': 'Convênios e Instrumentos Congêneres Arquivados',
  '812210101': 'Convênios e instrumentos a Liberar',
  '812210102': 'Convênios e Instrumentos Congêneres a Comprovar',
  '812210103': 'Convênios e Instrumentos Congêneres a Aprovar',
  '812210104': 'Convênios e Instrumentos Congêneres Aprovados',
  '812210105': 'Convênios e Instrumentos Congêneres Comprovados em Análise',
  '812210106': 'Convênios e instrumentos Congêneres em Inadimplência Efetiva',
  '812210107': 'Convênios e Instrumentos Congêneres em Inadimplência Suspensa',
  '812210108': 'Convênios e Instrumentos Congêneres Cancelados',
  '812210109': 'Convênios e Instrumentos Congêneres Não Liberado/ Devolvido',
  '812210111': 'Convênios e Instrumentos Congêneres Concluídos',
  '812210201': 'A repassar',
  '812210202': 'A Comprovar',
  '712210101': 'Valores Firmados',
};

// Contas dedicadas a instrumentos concluídos — conclusão só é válida quando o evento aparece NESTAS contas
const CONCLUSAO_ACCOUNTS = new Set(['812210111', '811210110']);
// Eventos de conclusão definitiva do instrumento (válidos apenas em CONCLUSAO_ACCOUNTS)
const CONCLUSAO_EVENTS = new Set(['581674', '581338', '580674', '581087']);
// Eventos que indicam remoção de inadimplência efetiva (conta 812210106 torna-se histórica)
const REMOCAO_INAD_EVENTS = new Set(['580742', '581742']);
// Eventos de remoção/estorno de inadimplência suspensa (conta 812210107 torna-se histórica)
const SUSP_INAD_EVENTS = new Set(['580711', '585709']);
// Eventos de estorno (reversão contábil) que indicam anulação processada em 812210101
const ESTORNO_EVENTS = new Set(['585716', '586700']);
// Eventos de aprovação de parcela registrados em 812210103 (aprovação sem encerramento formal)
// Série 58.0.xxx → NL/PORTALCONV (SICONV legado) | Série 58.1.xxx → NS/Transferegov
const APROVACAO_812103_EVENTS = new Set(['580707', '581720']);
// Eventos de inadimplência APÓS entrega da documentação — rito patológico dentro de 812210103
const INAD_APOS_DOC_EVENTS = new Set(['580741', '581741']);
// Situações TG terminais onde 812210103 sem evento rastreável é anomalia (strings pré-normalizadas)
const TERMINAL_TG_NORM = new Set([
  'prestacao de contas aprovada com ressalvas',
  'prestacao de contas aprovada',
  'prestacao de contas concluida',
  'prestacao de contas rejeitada',
  'convenio anulado',
  'instrumento anulado',
  'convenio rescindido',
  'cancelado',
]);

// Tabela de referência de eventos da conta 812210103 — v0.3.0 (2026-05-15)
// Fonte: SIAFI/TABAPOIO-EVENTO-CONEVENTO consultado em 15/05/2026
// Séries paralelas: 58.0.xxx → NL/PORTALCONV (SICONV legado) | 58.1.xxx → NS/Transferegov
const EVENTOS_812210103_META: Record<string, {
  descricao: string;
  tipoDoc: 'NL' | 'NS';
  faseConvenio: string;
  natureza: 'comprovacao' | 'aprovacao' | 'inadimplencia_pc';
  eventoEstorno: string;
  legitimoNaConta: boolean;
}> = {
  '581706': {
    descricao: 'Comprovação — Transf. Voluntárias',
    tipoDoc: 'NS',
    faseConvenio: 'PC apresentada — aguardando análise do concedente',
    natureza: 'comprovacao',
    eventoEstorno: '586706',
    legitimoNaConta: true,
  },
  '580707': {
    descricao: 'Aprovação — TV — Convênio (PORTALCONV/NL)',
    tipoDoc: 'NL',
    faseConvenio: 'PC aprovada — encerramento formal pendente no SIAFI',
    natureza: 'aprovacao',
    eventoEstorno: '585707',
    legitimoNaConta: false,
  },
  '581720': {
    descricao: 'Aprovação Transf. Voluntárias (Transferegov/NS)',
    tipoDoc: 'NS',
    faseConvenio: 'PC aprovada — encerramento formal pendente no SIAFI',
    natureza: 'aprovacao',
    eventoEstorno: '586720',
    legitimoNaConta: false,
  },
  '580741': {
    descricao: 'Inadim.TV — Convênio após doc. (PORTALCONV/NL — privativo UG)',
    tipoDoc: 'NL',
    faseConvenio: 'Inadimplência registrada durante análise da PC',
    natureza: 'inadimplencia_pc',
    eventoEstorno: '585741',
    legitimoNaConta: false,
  },
  '581741': {
    descricao: 'Inadimplência de T.V. Após Doc. (Transferegov/NS — ISF: N/A)',
    tipoDoc: 'NS',
    faseConvenio: 'Inadimplência registrada durante análise da PC',
    natureza: 'inadimplencia_pc',
    eventoEstorno: '586741',
    legitimoNaConta: false,
  },
};

const normalizeText = (text: string) => {
  if (!text) return "";
  return text.trim().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
};

// FIX-4 (Vuln 3): Canonicalize slash separators so "Proposta/Plano de Trabalho Aprovado"
// and "Proposta de Plano de Trabalho Aprovado" resolve to the same lookup key
// without needing a duplicate entry in validationRules.
const normalizeStatusKey = (text: string): string =>
  normalizeText(text).replace(/\s*\/\s*/g, ' de ');

// FIX-7 (Vuln 7): Depth-defense filters for the blind ID scanner.
// These codes are 6-digit numbers that appear in Transferegov exports but
// are NOT transfer IDs. Extend this list as new non-ID codes are discovered.
const KNOWN_EXCLUDED_CODES = new Set([
  "570000", // Código da UG — Skill §2.2 exclusão obrigatória
  "340068", // SIORG — Ministério do Turismo (MTur)
  "340011", // SIORG — entidade predecessor/vinculada MTur
  "110910", // UGE example code — adjust as needed
]);

// Column header keywords that identify columns that cannot contain a transfer ID.
// Used to narrow the blind scanner's search space.
const NON_ID_HEADER_KEYWORDS = [
  "siorg", "uf", "municipio", "município", "ibge", "cnpj", "cpf",
  "cep", "codigo", "código", "natureza", "funcional", "programa",
  "acao", "ação", "orgao", "órgão", "ugestora",
];

// SKILL §2.1: ID must be exactly 6 digits starting with 7, 8, or 9
// Instruments 700xxx–999xxx are valid Transferegov IDs; codes starting with
// 1–6 are UG/SIORG/IBGE codes and are excluded by the blacklist mechanism.
const extractTransferId = (val: string | number | undefined): string | null => {
  if (!val) return null;
  const str = String(val).trim();
  const match = str.match(/(?:^|\D)([789]\d{5})(?:\D|$)/);
  return match ? match[1] : null;
};

// Regex pré-compilado — reutilizado no loop de processamento SIAFI (B5)
const CONTA_9_DIGITS = /^\d{9}$/;

// ── Tema do Dashboard (dark-mode) ─────────────────────────────────────────
// Para reverter ao tema claro, substitua cada valor pelo comentado à direita.
const DASH = {
  cardText:   '#f1f5f9',               // revert → '#1e293b'
  cardMuted:  '#94a3b8',               // revert → '#475569'
  cardSubtle: '#64748b',               // unchanged
  chipBg:     'rgba(255,255,255,0.07)',// revert → '#f8fafc'
  barTrack:   'rgba(255,255,255,0.10)',// revert → '#e2e8f0'
  barBlue:    'linear-gradient(90deg,#1d4ed8,#3b82f6)', // revert → '#3b82f6'
  numBadge:   'rgba(96,165,250,0.15)', // revert → 'transparent'
  numColor:   '#60a5fa',               // revert → '#1e293b'
};
// ──────────────────────────────────────────────────────────────────────────

const PAGE_SIZE = 50;

interface ConciliacaoResult {
  idSiafi: string;
  valorSiafi: string;
  convenenteNome: string;
  convenenteSiafi: string;
  cnpjSiafi: string;
  situacaoRawTg: string;
  situacaoTg: string;
  situacaoSiafiDisplay: string;
  statusConciliacao: string;
  fullData: { "Status de Conciliação": string; "Motivo do Alerta": string };
}

class ErrorBoundary extends Component<
  { children: ReactNode },
  { error: Error | null }
> {
  constructor(props: { children: ReactNode }) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error: Error) {
    return { error };
  }
  componentDidCatch(error: Error, info: { componentStack: string }) {
    console.error('SIACT render error:', error, info.componentStack);
  }
  render() {
    if (this.state.error) {
      return (
        <div style={{ padding: '2rem', margin: '2rem', background: '#1e2d40', borderRadius: 12, border: '1px solid #ef4444' }}>
          <p style={{ color: '#ef4444', fontWeight: 700, fontSize: '1rem', marginBottom: '0.5rem' }}>
            Erro interno — copie o texto abaixo e envie para suporte:
          </p>
          <pre style={{ color: '#fca5a5', fontSize: '0.75rem', whiteSpace: 'pre-wrap', wordBreak: 'break-all', margin: '0 0 1rem', fontFamily: 'monospace' }}>
            {this.state.error.message}{'\n\n'}{this.state.error.stack}
          </pre>
          <button
            onClick={() => this.setState({ error: null })}
            style={{ padding: '0.5rem 1.2rem', background: '#3b82f6', color: '#fff', border: 'none', borderRadius: 8, cursor: 'pointer', fontSize: '0.85rem' }}
          >
            Resetar e tentar novamente
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function App() {
  const [fileTransferegov, setFileTransferegov] = useState<File | null>(null);
  const [fileSiafi, setFileSiafi] = useState<File | null>(null);
  const [results, setResults] = useState<ConciliacaoResult[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [stats, setStats] = useState({
    total: 0, corretos: 0, inconsistencias: 0, naoEncontrados: 0, alertas: 0,
  });
  const [filterNrInstrumento, setFilterNrInstrumento] = useState('');
  const [filterSituacaoTg, setFilterSituacaoTg] = useState('');
  const [filterContaSiafi, setFilterContaSiafi] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [modalOpen, setModalOpen] = useState<'guide' | 'compliance' | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [filterConvenente, setFilterConvenente] = useState('');
  const [currentPage, setCurrentPage] = useState(0);
  const [processingPhase, setProcessingPhase] = useState('');

  const dashboard = useMemo(() => {
    if (results.length === 0) return null;

    // SKILL §5.2: Métrica de Unicidade — contar IDs únicos, não linhas
    const tgFreq: Record<string, Set<string>> = {};
    const uniqueIds = new Set<string>();
    const accUniqueIds: Record<string, Set<string>> = {};

    for (const r of results) {
      const st = String(r.situacaoRawTg || "Não informado").trim() || "Não informado";
      if (!tgFreq[st]) tgFreq[st] = new Set();
      tgFreq[st].add(r.idSiafi);
      uniqueIds.add(r.idSiafi);
      for (const part of String(r.situacaoSiafiDisplay || "").split(' | ')) {
        const m = part.match(/^(\d{9})/);
        if (m) {
          if (!accUniqueIds[m[1]]) accUniqueIds[m[1]] = new Set();
          accUniqueIds[m[1]].add(r.idSiafi);
        }
      }
    }
    const accCounts: Record<string, number> = {};
    for (const [code, ids] of Object.entries(accUniqueIds)) {
      accCounts[code] = ids.size;
    }

    const tgEntries: [string, number][] = Object.entries(tgFreq)
      .map(([k, s]) => [k, s.size] as [string, number])
      .sort((a, b) => b[1] - a[1]).slice(0, 8);
    const tgMax = tgEntries[0]?.[1] || 1;

    const BI_PALETTE = ['#3b82f6','#10b981','#f59e0b','#8b5cf6','#ef4444','#06b6d4','#f97316','#84cc16','#ec4899','#6366f1'];
    const allAccEntries = Object.entries(accCounts)
      .map(([code, n]) => ({ code, n, fullLabel: accountMap[code] || code }))
      .sort((a, b) => b.n - a.n)
      .map((d, i) => ({ ...d, color: BI_PALETTE[i % BI_PALETTE.length] }));
    const allAccTotal = allAccEntries.reduce((s, a) => s + a.n, 0) || 1;
    const accMax = allAccEntries[0]?.n || 1;

    const C = 2 * Math.PI * 50;
    let off = 0;
    const allSegments = allAccEntries.map(d => {
      const frac = d.n / allAccTotal;
      const seg = { ...d, dasharray: `${(frac * C).toFixed(2)} ${C.toFixed(2)}`, dashoffset: -(off * C) };
      off += frac;
      return seg;
    });

    return { tgEntries, tgMax, allSegments, allAccEntries, allAccTotal, accMax,
             uniqueInstruments: uniqueIds.size, totalEntries: results.length };
  }, [results]);

  const uniqueTgSituacoes = useMemo(() => {
    const set = new Set<string>();
    for (const r of results) {
      if (r.situacaoRawTg) set.add(r.situacaoRawTg);
    }
    return [...set].sort();
  }, [results]);

  const uniqueContasSiafi = useMemo(() => {
    const map = new Map<string, string>();
    for (const r of results) {
      const parts = (r.situacaoSiafiDisplay || '').split(' | ');
      for (const p of parts) {
        const m = p.match(/^(\d{9})/);
        if (m && !map.has(m[1])) {
          const full = accountMap[m[1]] || `Conta ${m[1]}`;
          const short = full
            .replace(/^Convênios e [Ii]nstrumentos [Cc]ongêneres\s+/i, '')
            .replace(/^Convênios e instrumentos\s+/i, '');
          map.set(m[1], short);
        }
      }
    }
    return [...map.entries()]
      .sort((a, b) => a[1].localeCompare(b[1]))
      .map(([code, name]) => ({ code, name }));
  }, [results]);

  const filteredResults = useMemo(() => {
    const idTrim = filterNrInstrumento.trim();
    const convTrim = filterConvenente.trim().toLowerCase();
    return results.filter(r => {
      if (idTrim && !String(r.idSiafi ?? '').includes(idTrim)) return false;
      if (filterSituacaoTg && r.situacaoRawTg !== filterSituacaoTg) return false;
      if (filterContaSiafi && !String(r.situacaoSiafiDisplay ?? '').includes(filterContaSiafi)) return false;
      if (filterStatus && r.statusConciliacao !== filterStatus) return false;
      if (convTrim && !String(r.convenenteNome ?? '').toLowerCase().includes(convTrim) &&
          !String(r.cnpjSiafi ?? '').includes(filterConvenente.trim())) return false;
      return true;
    });
  }, [results, filterNrInstrumento, filterSituacaoTg, filterContaSiafi, filterStatus, filterConvenente]);

  const paginatedResults = useMemo(() => {
    const start = currentPage * PAGE_SIZE;
    return filteredResults.slice(start, start + PAGE_SIZE);
  }, [filteredResults, currentPage]);

  const filteredStats = useMemo(() => {
    const seen = new Set<string>();
    let corretos = 0, inconsistencias = 0, naoEncontrados = 0, alertas = 0;
    for (const r of filteredResults) {
      if (seen.has(r.idSiafi)) continue;
      seen.add(r.idSiafi);
      if (r.statusConciliacao === 'Correto') corretos++;
      else if (r.statusConciliacao === 'Inconsistência (Rito Patológico)') inconsistencias++;
      else if (r.statusConciliacao === 'Sem Registro no Transferegov') naoEncontrados++;
      else alertas++;
    }
    const total = seen.size;
    const items = [
      { label: 'Corretos',              val: corretos,         color: '#10b981' },
      { label: 'Inconsistências',       val: inconsistencias,  color: '#ef4444' },
      { label: 'Alertas — Revisar',     val: alertas,          color: '#f97316' },
      { label: 'Sem Ref. Transferegov', val: naoEncontrados,   color: '#f59e0b' },
    ];
    const C = 2 * Math.PI * 62;
    let off = 0;
    const segments = items.filter(d => d.val > 0).map(d => {
      const frac = d.val / (total || 1);
      const seg = { ...d, dasharray: `${(frac * C).toFixed(2)} ${C.toFixed(2)}`, dashoffset: -(off * C) };
      off += frac;
      return seg;
    });
    return { corretos, inconsistencias, naoEncontrados, alertas, total, items, segments };
  }, [filteredResults]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: 'transferegov' | 'siafi') => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (type === 'transferegov') setFileTransferegov(file);
    else setFileSiafi(file);
  };

  const processFiles = async () => {
    if (!fileTransferegov || !fileSiafi) return;
    setIsProcessing(true);
    setErrorMsg(null);
    setProcessingPhase("Lendo arquivos...");

    try {
      const readExcelMatrix = (file: File) => {
        return new Promise<any[][]>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (e) => {
            try {
              const data = new Uint8Array(e.target?.result as ArrayBuffer);
              const workbook = XLSX.read(data, { type: 'array' });
              const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
              // FIX-1 (Vuln 5): Must run before sheet_to_json
              expandMergedCells(firstSheet);
              const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" }) as any[][];
              resolve(rows);
            } catch (err) {
              reject(err);
            }
          };
          reader.onerror = reject;
          reader.readAsArrayBuffer(file);
        });
      };

      const [rowsTg, rowsSiafi] = await Promise.all([
        readExcelMatrix(fileTransferegov),
        readExcelMatrix(fileSiafi),
      ]);

      setProcessingPhase("Processando Transferegov...");
      if (rowsTg.length === 0 || rowsSiafi.length === 0) {
        setErrorMsg("Uma das planilhas está vazia! Verifique os arquivos selecionados.");
        setIsProcessing(false);
        return;
      }

      // --- 1. Process Transferegov ---
      let idColTgIdx = -1;
      let statusColTgIdx = -1;
      let convenenteColTgIdx = -1;
      let tgHeaderRowIdx = -1;

      for (let i = 0; i < Math.min(10, rowsTg.length); i++) {
        const r = rowsTg[i];
        if (!r) continue;
        r.forEach((cell, idx) => {
          const norm = normalizeText(String(cell));
          if (idColTgIdx === -1 && (norm.includes("transferencia") || norm.includes("convenio"))) {
            idColTgIdx = idx;
            tgHeaderRowIdx = i;
          }
          if (statusColTgIdx === -1 && norm.includes("situacao")) {
            statusColTgIdx = idx;
          }
          if (convenenteColTgIdx === -1 &&
            (norm.includes("recebedor") || norm.includes("convenente") || norm.includes("favorecido") || norm.includes("proponente")) &&
            !norm.includes("cnpj") && !norm.includes("uf") && !norm.includes("municipio") && !norm.includes("valor")) {
            convenenteColTgIdx = idx;
          }
        });
      }

      // FIX-7 (Vuln 7): Build the exclusion set from column headers to
      // prevent the blind scanner from picking up SIORG/UF/CNPJ codes.
      const excludedTgCols = new Set<number>();
      if (tgHeaderRowIdx !== -1) {
        const headerRow = rowsTg[tgHeaderRowIdx] || [];
        headerRow.forEach((cell, idx) => {
          const norm = normalizeText(String(cell));
          if (NON_ID_HEADER_KEYWORDS.some(kw => norm.includes(kw))) {
            excludedTgCols.add(idx);
          }
        });
      }

      const tgMap = new Map<string, { status: string; convenente: string; ambiguous?: boolean; fullRow: any[] }>();

      for (let i = 0; i < rowsTg.length; i++) {
        const row = rowsTg[i];
        if (!row || row.length === 0) continue;

        let id = "";
        let status = "";
        let convenente = "";
        let ambiguousRow = false;

        // LAYER 1 — header-based detection (always preferred, zero ambiguity)
        const candidateId = idColTgIdx !== -1 ? extractTransferId(row[idColTgIdx]) : null;
        if (candidateId) {
          id = candidateId;
        } else {
          // LAYER 2 — FIX-7 blind scanner with depth-defense:
          // · Skip columns whose headers are known non-ID types
          // · Skip cells whose value is in the known-exclusion list
          // · Require exactly 1 surviving candidate; if 2+, flag as ambiguous
          const idCandidates: string[] = [];
          for (let colIdx = 0; colIdx < row.length; colIdx++) {
            if (excludedTgCols.has(colIdx)) continue;
            const match = extractTransferId(row[colIdx]);
            if (match && !KNOWN_EXCLUDED_CODES.has(match)) {
              idCandidates.push(match);
            }
          }

          if (idCandidates.length === 1) {
            id = idCandidates[0];
          } else if (idCandidates.length > 1) {
            // Register every candidate as ambiguous so SIAFI lookups are flagged
            ambiguousRow = true;
            idCandidates.forEach(candidate => {
              if (!tgMap.has(candidate)) {
                tgMap.set(candidate, { status: "", convenente: "", ambiguous: true, fullRow: row });
              }
            });
          }
        }

        if (ambiguousRow || !id) continue;

        if (statusColTgIdx !== -1 && String(row[statusColTgIdx]).trim()) {
          status = String(row[statusColTgIdx]).trim();
        } else {
          for (const cell of row) {
            const norm = normalizeText(String(cell));
            if (Object.keys(validationRules).some(rule => normalizeText(rule) === norm)) {
              status = String(cell).trim();
              break;
            }
          }
        }

        if (convenenteColTgIdx !== -1 && row[convenenteColTgIdx] !== undefined) {
          convenente = String(row[convenenteColTgIdx]).trim();
        } else {
          for (const cell of row) {
            const str = String(cell).trim();
            if (str.length > 15 && str === str.toUpperCase() && !str.match(/^\d+$/)) {
              convenente = str;
              break;
            }
          }
        }

        if (status && normalizeText(status) !== "situacao") {
          tgMap.set(id, { status, convenente, fullRow: row });
        }
      }

      setProcessingPhase("Processando SIAFI...");
      // --- 2. Process SIAFI (formato normalizado — novo formato MTur 2026) ---
      // Cada linha = instrumento × conta contábil × evento
      // Fase 1: Localizar colunas pelos cabeçalhos
      let siafiHeaderRowIdx = -1;
      let contaColIdx          = -1;  // "Conta Contábil" → código 9 dígitos
      let valorColIdx          = -1;  // "Transferência - Valor" → valor de referência
      let eventoColIdx         = -1;  // "Evento" → código do evento contábil
      let cnpjColSiafiIdx      = -1;  // CNPJ do convenente
      let convenentColSiafiIdx = -1;  // Nome do convenente

      for (let i = 0; i < Math.min(20, rowsSiafi.length); i++) {
        const r = rowsSiafi[i];
        if (!r) continue;
        r.forEach((cell, idx) => {
          const text = String(cell).trim();
          const norm = normalizeText(text);

          // Exact match to avoid false positive from metadata row
          // "Métrica: Saldo - Moeda Origem (Conta Contábil)" which also contains both words
          if (contaColIdx === -1 && norm === 'conta contabil')
            contaColIdx = idx;

          if (valorColIdx === -1 && norm.includes('transferencia') && norm.includes('valor') && !norm.includes('contrapartida') && !norm.includes('saldo'))
            valorColIdx = idx;

          if (eventoColIdx === -1 && (norm === 'evento' || (norm.startsWith('evento') && text.length <= 25)))
            eventoColIdx = idx;

          if (cnpjColSiafiIdx === -1 && norm.includes('cnpj'))
            cnpjColSiafiIdx = idx;

          if (convenentColSiafiIdx === -1 && norm.includes('convenente') && !norm.includes('cnpj') && !norm.includes('uf') && !norm.includes('municipio'))
            convenentColSiafiIdx = idx;
        });
        if (contaColIdx !== -1 && siafiHeaderRowIdx === -1) siafiHeaderRowIdx = i;
      }

      if (contaColIdx === -1) {
        setErrorMsg("Não foi possível identificar a coluna 'Conta Contábil' na planilha SIAFI. Verifique o formato do arquivo.");
        setIsProcessing(false);
        return;
      }

      // Fallback data-based CNPJ: se header não detectou, busca coluna com 40%+ de valores de 11-14 dígitos
      if (cnpjColSiafiIdx === -1 && siafiHeaderRowIdx !== -1) {
        const sampleRows = rowsSiafi.slice(siafiHeaderRowIdx + 1, siafiHeaderRowIdx + 21).filter(r => r && r.length > 0);
        if (sampleRows.length > 0) {
          const colCount = Math.max(...sampleRows.map(r => r.length));
          for (let col = 0; col < colCount; col++) {
            const hits = sampleRows.filter(r => {
              const d = String(r[col] ?? "").replace(/\D/g, "");
              return d.length >= 11 && d.length <= 14;
            }).length;
            if (hits >= Math.ceil(sampleRows.length * 0.4)) { cnpjColSiafiIdx = col; break; }
          }
        }
      }

      // Fallback data-based Convenente: coluna com texto longo não-numérico em 50%+ das amostras
      if (convenentColSiafiIdx === -1 && siafiHeaderRowIdx !== -1) {
        const sampleRows = rowsSiafi.slice(siafiHeaderRowIdx + 1, siafiHeaderRowIdx + 21).filter(r => r && r.length > 0);
        if (sampleRows.length > 0) {
          const colCount = Math.max(...sampleRows.map(r => r.length));
          for (let col = 0; col < colCount; col++) {
            if (col === cnpjColSiafiIdx || col === contaColIdx) continue;
            const hits = sampleRows.filter(r => {
              const val = String(r[col] ?? "").trim();
              return val.length > 10 && !/^\d[\d.,\s]*$/.test(val) && !/^\d{9}$/.test(val);
            }).length;
            if (hits >= Math.ceil(sampleRows.length * 0.5)) { convenentColSiafiIdx = col; break; }
          }
        }
      }

      // Blacklist: código 7xxxxx que apareça em >50% das linhas da col A é UG, não instrumento
      const colAFreq: Record<string, number> = {};
      let colADataRows = 0;
      for (let i = siafiHeaderRowIdx + 1; i < rowsSiafi.length; i++) {
        const r = rowsSiafi[i];
        if (!r || r.length === 0) continue;
        const raw = String(r[0] ?? "").replace(/\D/g, "");
        if (raw.length === 6 && /[789]/.test(raw[0])) {
          colAFreq[raw] = (colAFreq[raw] || 0) + 1;
          colADataRows++;
        }
      }
      const ugCodeBlacklist = new Set<string>(
        Object.entries(colAFreq)
          .filter(([, n]) => colADataRows > 0 && n / colADataRows > 0.5)
          .map(([k]) => k)
      );

      // Fase 2: Construir mapa de instrumentos (ID → contas + eventos + CNPJ + Convenente)
      type ContraEntry = { display: string; eventos: string[] };
      const instrumentMap = new Map<string, {
        contas: Map<string, ContraEntry>;
        cnpj: string;
        convenente: string;
        valor: string;
      }>();
      let lastIdSiafi = "";
      let lastCnpjSiafi = "";
      let lastConvenenteSiafi = "";

      for (let i = siafiHeaderRowIdx + 1; i < rowsSiafi.length; i++) {
        const row = rowsSiafi[i];
        if (!row || row.length === 0) continue;

        // ID col A — carry-forward para células mescladas
        let idSiafi = extractTransferId(row[0]);
        if (idSiafi && ugCodeBlacklist.has(idSiafi)) idSiafi = null;
        if (idSiafi) lastIdSiafi = idSiafi;
        else idSiafi = lastIdSiafi;
        if (!idSiafi) continue;

        // Conta Contábil — deve ser exatamente 9 dígitos
        const contaRaw = String(row[contaColIdx] ?? "").trim();
        if (!CONTA_9_DIGITS.test(contaRaw)) continue;

        // CNPJ (carry-forward)
        if (cnpjColSiafiIdx !== -1) {
          const cnpjRaw = String(row[cnpjColSiafiIdx] ?? "").trim();
          const d = cnpjRaw.replace(/\D/g, "");
          if (d.length >= 11 && d.length <= 14) lastCnpjSiafi = cnpjRaw;
        }

        // Convenente nome (carry-forward — rejeita valores numéricos)
        if (convenentColSiafiIdx !== -1) {
          const convRaw = String(row[convenentColSiafiIdx] ?? "").trim();
          if (convRaw && convRaw.length > 5 && !/^\d[\d.,\s]*$/.test(convRaw))
            lastConvenenteSiafi = convRaw;
        }

        // Inicializar entrada do instrumento
        if (!instrumentMap.has(idSiafi)) {
          instrumentMap.set(idSiafi, { contas: new Map(), cnpj: lastCnpjSiafi, convenente: lastConvenenteSiafi, valor: "" });
        }
        const entry = instrumentMap.get(idSiafi)!;
        if (!entry.cnpj && lastCnpjSiafi) entry.cnpj = lastCnpjSiafi;
        if (!entry.convenente && lastConvenenteSiafi) entry.convenente = lastConvenenteSiafi;
        if (!entry.valor && valorColIdx !== -1) {
          const rawVal = String(row[valorColIdx] ?? "").trim();
          if (rawVal && rawVal !== "0" && rawVal !== "-9") entry.valor = rawVal;
        }

        // Registrar conta contábil
        if (!entry.contas.has(contaRaw)) {
          const acName = accountMap[contaRaw] || contaRaw;
          entry.contas.set(contaRaw, { display: `${contaRaw} - ${acName}`, eventos: [] });
        }

        // Registrar evento de referência (ignorar "-9" = NAO SE APLICA)
        const eventoCode = eventoColIdx !== -1 ? String(row[eventoColIdx] ?? "").trim() : "";
        if (eventoCode && eventoCode !== "-9" && eventoCode !== "0") {
          const contraEntry = entry.contas.get(contaRaw)!;
          if (!contraEntry.eventos.includes(eventoCode)) {
            contraEntry.eventos.push(eventoCode);
          }
        }
      }

      // Fase 3: Gerar resultados a partir do mapa de instrumentos
      let corretos = 0, inconsistencias = 0, naoEncontrados = 0, alertas = 0;
      const confirmedCorrectIds = new Set<string>();
      const finalResults: ConciliacaoResult[] = [];

      for (const [idSiafi, entry] of instrumentMap) {
        const detectedAccounts = [...entry.contas.entries()].map(([code, info]) => ({
          code,
          display: info.display,
          eventos: info.eventos,
        }));
        // Declarado no escopo externo para ser acessível tanto na lógica de status quanto nos motivos
        const conta812103 = detectedAccounts.find(a => a.code === '812210103');

        const situacaoSiafiDisplay =
          detectedAccounts.length === 0
            ? "Sem Saldo no SIAFI"
            : detectedAccounts.map(a => {
                const refEventos = a.eventos.length > 0
                  ? ` [Ref. Eventos: ${a.eventos.join(" · ")}]`
                  : "";
                return a.display + refEventos;
              }).join(" | ");

        const convenenteSiafi = entry.convenente;
        const cnpjSiafi = entry.cnpj;
        const valorSiafi = entry.valor || "";

        let situacaoTg = "Sem Registro no Transferegov";
        let situacaoRawTg = "Sem Registro no Transferegov";
        let convenenteNome = convenenteSiafi || "Não Informado";
        let statusConciliacao = "";
        let contasForaDoConjunto: string[] = [];

        const tgData = tgMap.get(idSiafi);
        if (tgData && !tgData.ambiguous) {
          situacaoRawTg = tgData.status || (statusColTgIdx !== -1 ? String(tgData.fullRow[statusColTgIdx] ?? "").trim() : "");
        }

        if (tgData) {
          if (tgData.ambiguous) {
            situacaoTg = "Ambiguidade no ID (TG)";
            statusConciliacao = "Revisão Manual - ID Ambíguo";
            if (!confirmedCorrectIds.has(idSiafi)) alertas++;
          } else {
            situacaoTg = tgData.status;
            if (tgData.convenente && tgData.convenente !== "-") convenenteNome = tgData.convenente;

            // Resolução por prioridade de evento: conclusão (conta-específica) > inad. efetiva/suspensa > estorno > padrão
            const allEventos = detectedAccounts.flatMap(a => a.eventos);
            // Conclusão só é válida quando o evento aparece em conta dedicada (812210111 ou 811210110)
            const hasConclusao = detectedAccounts.some(a =>
              CONCLUSAO_ACCOUNTS.has(a.code) && a.eventos.some(e => CONCLUSAO_EVENTS.has(e))
            );
            const hasRemocaoInad = !hasConclusao && detectedAccounts.some(a =>
              a.code === '812210106' && a.eventos.some(e => REMOCAO_INAD_EVENTS.has(e))
            );
            const hasRemocaoSuspInad = !hasConclusao && detectedAccounts.some(a =>
              a.code === '812210107' && a.eventos.some(e => SUSP_INAD_EVENTS.has(e))
            );
            const hasEstorno = allEventos.some(e => ESTORNO_EVENTS.has(e));
            const isInstrumentoAnulado = normalizeStatusKey(situacaoTg) === normalizeStatusKey("Instrumento Anulado");

            // P0 — Motor de eventos 812210103 (v0.3.0)
            // Prioridade: inad_pc > aprovacao > sem_evento (todos excluídos quando hasConclusao=true)
            const situacaoTgNorm = normalizeStatusKey(situacaoTg);
            const has812103InadAposDoc = !hasConclusao && conta812103 !== undefined
              && conta812103.eventos.some(e => INAD_APOS_DOC_EVENTS.has(e));
            const has812103Aprovacao = !hasConclusao && !has812103InadAposDoc
              && conta812103 !== undefined
              && conta812103.eventos.some(e => APROVACAO_812103_EVENTS.has(e));
            const has812103SemEvento = !hasConclusao && !has812103InadAposDoc && !has812103Aprovacao
              && conta812103 !== undefined && conta812103.eventos.length === 0
              && TERMINAL_TG_NORM.has(situacaoTgNorm);

            // Contas efetivas para validação após filtragem de histórico de ciclo de vida
            let accountsForValidation = detectedAccounts;
            if (hasConclusao) {
              // Manter apenas contas de conclusão com evento de conclusão; demais são histórico
              accountsForValidation = detectedAccounts.filter(a =>
                CONCLUSAO_ACCOUNTS.has(a.code) && a.eventos.some(e => CONCLUSAO_EVENTS.has(e))
              );
            } else if (hasRemocaoInad || hasRemocaoSuspInad) {
              // Remover contas de inadimplência cujos eventos registram encerramento
              accountsForValidation = detectedAccounts.filter(a => {
                if (a.code === '812210106' && a.eventos.some(e => REMOCAO_INAD_EVENTS.has(e))) return false;
                if (a.code === '812210107' && a.eventos.some(e => SUSP_INAD_EVENTS.has(e))) return false;
                return true;
              });
            }

            if (detectedAccounts.length === 0) {
              statusConciliacao = "Sem Saldo no SIAFI";
              if (!confirmedCorrectIds.has(idSiafi)) alertas++;
            } else if (has812103InadAposDoc) {
              // P0-A: inadimplência após entrega da documentação — rito patológico em 812210103
              statusConciliacao = "Inconsistência — Inadimplência na fase A Aprovar";
              if (!confirmedCorrectIds.has(idSiafi)) inconsistencias++;
            } else if (has812103Aprovacao) {
              // P0-B: aprovação lançada em 812210103 sem encerramento formal no SIAFI
              statusConciliacao = "Atenção — Aprovação sem encerramento no SIAFI";
              if (!confirmedCorrectIds.has(idSiafi)) alertas++;
            } else if (has812103SemEvento) {
              // P0-C: 812210103 sem evento rastreável (TG=-9) em situação TG terminal
              statusConciliacao = "Revisão Manual — Sem Rastreabilidade de Evento";
              if (!confirmedCorrectIds.has(idSiafi)) alertas++;
            } else if (hasConclusao && normalizeStatusKey(situacaoTg) !== normalizeStatusKey("Prestação de Contas Concluída")) {
              statusConciliacao = "Atenção — Conclusão SIAFI sem encerramento no TG";
              if (!confirmedCorrectIds.has(idSiafi)) alertas++;
            } else if (isInstrumentoAnulado && detectedAccounts.some(a => a.code === '812210101') && !hasEstorno) {
              statusConciliacao = "Atenção — Anulação sem estorno contábil no SIAFI";
              if (!confirmedCorrectIds.has(idSiafi)) alertas++;
            } else {
              const matchedRuleKey = Object.keys(validationRules).find(
                rule => normalizeStatusKey(rule) === normalizeStatusKey(situacaoTg)
              );
              if (!matchedRuleKey) {
                statusConciliacao = "Status Não Mapeado";
                if (!confirmedCorrectIds.has(idSiafi)) inconsistencias++;
              } else {
                const validSet = validationRules[matchedRuleKey];
                if (validSet.length === 0) {
                  statusConciliacao = "Regra Pendente - Revisar";
                  if (!confirmedCorrectIds.has(idSiafi)) alertas++;
                } else {
                  contasForaDoConjunto = accountsForValidation
                    .filter(a => !validSet.includes(a.code))
                    .map(a => a.display);
                  if (contasForaDoConjunto.length === 0) {
                    statusConciliacao = "Correto";
                    corretos++;
                    confirmedCorrectIds.add(idSiafi);
                  } else {
                    statusConciliacao = "Inconsistência (Rito Patológico)";
                    if (!confirmedCorrectIds.has(idSiafi)) inconsistencias++;
                  }
                }
              }
            }
          }
        } else {
          statusConciliacao = "Sem Registro no Transferegov";
          if (!confirmedCorrectIds.has(idSiafi)) naoEncontrados++;
        }

        let motivo = "";
        if (statusConciliacao === "Correto") {
          const allEv = detectedAccounts.flatMap(a => a.eventos);
          const hasConclusaoMotivo = detectedAccounts.some(a =>
            CONCLUSAO_ACCOUNTS.has(a.code) && a.eventos.some(e => CONCLUSAO_EVENTS.has(e))
          );
          const hasRemocaoMotivo = !hasConclusaoMotivo && detectedAccounts.some(a =>
            a.code === '812210106' && a.eventos.some(e => REMOCAO_INAD_EVENTS.has(e))
          );
          const hasRemocaoSuspMotivo = !hasConclusaoMotivo && detectedAccounts.some(a =>
            a.code === '812210107' && a.eventos.some(e => SUSP_INAD_EVENTS.has(e))
          );
          const hasEstornoMotivo = allEv.some(e => ESTORNO_EVENTS.has(e));
          const isAnuladoMotivo = normalizeStatusKey(situacaoTg) === normalizeStatusKey("Instrumento Anulado");
          if (hasConclusaoMotivo) {
            const evConc = [...new Set(detectedAccounts.filter(a => CONCLUSAO_ACCOUNTS.has(a.code)).flatMap(a => a.eventos).filter(e => CONCLUSAO_EVENTS.has(e)))];
            const historico = detectedAccounts.length - detectedAccounts.filter(a => CONCLUSAO_ACCOUNTS.has(a.code) && a.eventos.some(e => CONCLUSAO_EVENTS.has(e))).length;
            motivo = historico > 0
              ? `OK — Conclusão confirmada (ev. ${evConc.join(", ")}). ${historico} conta(s) histórica(s) de ciclo de vida desconsiderada(s).`
              : `OK — Conclusão confirmada por evento SIAFI (ev. ${evConc.join(", ")}).`;
          } else if (hasRemocaoMotivo && hasRemocaoSuspMotivo) {
            const evRem = [...new Set([...allEv.filter(e => REMOCAO_INAD_EVENTS.has(e)), ...allEv.filter(e => SUSP_INAD_EVENTS.has(e))])];
            motivo = `OK — Inadimplência efetiva e suspensa encerradas (ev. ${evRem.join(", ")}); contas 812210106 e 812210107 tratadas como históricas.`;
          } else if (hasRemocaoMotivo) {
            const evRem = [...new Set(allEv.filter(e => REMOCAO_INAD_EVENTS.has(e)))];
            motivo = `OK — Inadimplência efetiva encerrada (ev. ${evRem.join(", ")}); conta 812210106 tratada como histórica.`;
          } else if (hasRemocaoSuspMotivo) {
            const evRem = [...new Set(allEv.filter(e => SUSP_INAD_EVENTS.has(e)))];
            motivo = `OK — Inadimplência suspensa encerrada (ev. ${evRem.join(", ")}); conta 812210107 tratada como histórica.`;
          } else if (hasEstornoMotivo && isAnuladoMotivo) {
            const evEst = [...new Set(allEv.filter(e => ESTORNO_EVENTS.has(e)))];
            motivo = `OK — Anulação contabilmente processada. Estornos registrados (ev. ${evEst.join(", ")}); conta 812210101 encerrada sem saldo ativo.`;
          } else {
            const nContas = detectedAccounts.length;
            motivo = nContas > 1
              ? `OK — ${nContas} contas dentro do conjunto válido (Rito Ordinário): ${situacaoSiafiDisplay}`
              : "OK";
          }
        } else if (statusConciliacao === "Atenção — Anulação sem estorno contábil no SIAFI") {
          motivo = `Instrumento anulado no TG mas conta 812210101 (A Liberar) permanece ativa no SIAFI sem eventos de estorno (${[...ESTORNO_EVENTS].join(", ")}). Verificar lançamento de estorno contábil.`;
        } else if (statusConciliacao === "Atenção — Conclusão SIAFI sem encerramento no TG") {
          const evConc = [...new Set(detectedAccounts.flatMap(a => a.eventos).filter(e => CONCLUSAO_EVENTS.has(e)))];
          motivo = `SIAFI registra conclusão (ev. ${evConc.join(", ")}) mas TG indica '${situacaoTg}'. Verificar atualização no Transferegov.`;
        } else if (statusConciliacao === "Atenção — Aprovação sem encerramento no SIAFI") {
          const aprovEvs = conta812103?.eventos.filter(e => APROVACAO_812103_EVENTS.has(e)) ?? [];
          const aprovMeta = aprovEvs.map(e => EVENTOS_812210103_META[e]?.descricao ?? e).join(', ');
          motivo = `SIAFI registra aprovação de parcela (ev. ${aprovEvs.join(', ')} — ${aprovMeta}) em 812210103 (A Aprovar). Encerramento formal pendente: instrumento deve migrar para 812210104 ou conta de conclusão. TG indica '${situacaoTg}'. Verificar NS de encerramento no CADREDUZTV.`;
        } else if (statusConciliacao === "Inconsistência — Inadimplência na fase A Aprovar") {
          const inadEvs = conta812103?.eventos.filter(e => INAD_APOS_DOC_EVENTS.has(e)) ?? [];
          const inadMeta = inadEvs.map(e => EVENTOS_812210103_META[e]?.descricao ?? e).join(', ');
          motivo = `Inadimplência durante análise da PC (ev. ${inadEvs.join(', ')} — ${inadMeta}) em 812210103 (A Aprovar). Rito patológico: convenente declarado inadimplente após entrega da documentação. Verificar: cancelamento da inadimplência, arquivamento ou providências administrativas pendentes.`;
        } else if (statusConciliacao === "Revisão Manual — Sem Rastreabilidade de Evento") {
          motivo = `Conta 812210103 (A Aprovar) sem evento contábil rastreável (TG retorna -9) com situação TG terminal ('${situacaoTg}'). Possível instrumento SICONV pré-2014 ou saldo residual sem origem identificável. Verificar CADREDUZTV (TRANSF-CADREDUZTV-CONTVREDUZ) e histórico de movimentação.`;
        } else if (statusConciliacao === "Inconsistência (Rito Patológico)") {
          motivo = `Conta(s) fora do conjunto válido para '${situacaoTg}': ${contasForaDoConjunto.join(" | ")}`;
        } else if (statusConciliacao === "Revisão Manual - ID Ambíguo") {
          motivo = "Múltiplos IDs candidatos na mesma linha do Transferegov.";
        } else if (statusConciliacao === "Regra Pendente - Revisar") {
          motivo = `Regra não configurada para: '${situacaoTg}'.`;
        } else if (statusConciliacao === "Status Não Mapeado") {
          motivo = `A situação '${situacaoTg}' não possui mapeamento no sistema.`;
        } else if (statusConciliacao === "Sem Saldo no SIAFI") {
          motivo = "Nenhuma conta contábil detectada no SIAFI.";
        } else if (statusConciliacao === "Sem Registro no Transferegov") {
          motivo = "ID não consta na base do Transferegov.";
        } else {
          motivo = "Desconhecido";
        }

        finalResults.push({
          idSiafi,
          valorSiafi,
          convenenteNome,
          convenenteSiafi,
          cnpjSiafi,
          situacaoRawTg,
          situacaoTg,
          situacaoSiafiDisplay,
          statusConciliacao,
          fullData: { "Status de Conciliação": statusConciliacao, "Motivo do Alerta": motivo },
        });
      }

      setStats({ total: finalResults.length, corretos, inconsistencias, naoEncontrados, alertas });
      setResults(finalResults);
      setCurrentPage(0);
      setProcessingPhase("");
    } catch (err) {
      console.error(err);
      setErrorMsg("Erro ao processar planilhas. Verifique o formato dos arquivos e tente novamente.");
      setProcessingPhase("");
    }

    setIsProcessing(false);
  };

  // SKILL §3 — Exportação com esquema fixo A-G (Validação de Colunas)
  // A: ID  B: Convenente  C: CNPJ  D: Valor (R$)  E: Situação Transferegov
  // F: Situação SIAFI  G: Status  H: Motivo do Alerta
  const exportResults = () => {
    const dataToExport = filteredResults.map(r => ({
      "ID": r.idSiafi,
      "Convenente": r.convenenteSiafi || r.convenenteNome || "",
      "CNPJ": r.cnpjSiafi || "",
      "Valor (R$)": r.valorSiafi || "",
      "Situação Transferegov": r.situacaoRawTg,
      "Situação SIAFI": r.situacaoSiafiDisplay,
      "Status (Conciliação)": r.statusConciliacao,
      "Motivo do Alerta": r.fullData["Motivo do Alerta"],
    }));
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Conciliação");
    XLSX.writeFile(wb, "Resultado_Conciliacao_SIACT.xlsx");
  };

  const getStatusBadgeClass = (status: string): string => {
    if (status === 'Correto') return 'status-correct';
    if (status === 'Sem Registro no Transferegov') return 'status-not-found';
    if (['Regra Pendente - Revisar', 'Sem Saldo no SIAFI',
         'Revisão Manual - ID Ambíguo', 'Status Não Mapeado',
         'Atenção — Conclusão SIAFI sem encerramento no TG',
         'Atenção — Anulação sem estorno contábil no SIAFI',
         'Atenção — Aprovação sem encerramento no SIAFI',
         'Revisão Manual — Sem Rastreabilidade de Evento'].includes(status)) return 'status-alert';
    return 'status-incorrect';
  };

  const getStatusIcon = (status: string) => {
    const style = { display: 'inline' as const, marginRight: '4px', verticalAlign: 'text-bottom' as const };
    if (status === 'Correto') return <CheckCircle size={14} style={style} />;
    if (['Regra Pendente - Revisar', 'Sem Saldo no SIAFI',
         'Revisão Manual - ID Ambíguo', 'Status Não Mapeado',
         'Sem Registro no Transferegov',
         'Atenção — Conclusão SIAFI sem encerramento no TG',
         'Atenção — Anulação sem estorno contábil no SIAFI',
         'Atenção — Aprovação sem encerramento no SIAFI',
         'Revisão Manual — Sem Rastreabilidade de Evento'].includes(status)) return <AlertCircle size={14} style={style} />;
    return <XCircle size={14} style={style} />;
  };

  return (
    <ErrorBoundary>
    <div className="app-container" translate="no">
      <div className="header">
        <h1>SIACT Hub</h1>
        <p>Auditor Digital MTur — Conciliação Transferegov × SIAFI (Tesouro Gerencial)</p>
      </div>

      {/* ── Botões de Informação ── */}
      <div style={{ display: 'flex', justifyContent: 'space-between', gap: '1rem', marginBottom: '0.75rem' }}>
        <button
          onClick={() => setModalOpen('guide')}
          style={{
            display: 'flex', alignItems: 'center', gap: '0.5rem',
            padding: '0.5rem 1.1rem', borderRadius: 8, border: '1px solid #334155',
            background: '#1e293b', color: '#60a5fa', cursor: 'pointer',
            fontSize: '0.82rem', fontWeight: 600,
          }}
        >
          <Settings size={15} /> Guia de Operação e Metas
        </button>
        <button
          onClick={() => setModalOpen('compliance')}
          style={{
            display: 'flex', alignItems: 'center', gap: '0.5rem',
            padding: '0.5rem 1.1rem', borderRadius: 8, border: '1px solid #334155',
            background: '#1e293b', color: '#34d399', cursor: 'pointer',
            fontSize: '0.82rem', fontWeight: 600,
          }}
        >
          <Shield size={15} /> Conformidade e Avaliação Digital
        </button>
      </div>

      <div className="upload-grid">
        <div className={`upload-card ${fileTransferegov ? 'loaded' : ''}`}>
          <input
            type="file"
            className="file-input"
            accept=".xlsx, .xls, .csv"
            onChange={(e) => handleFileUpload(e, 'transferegov')}
          />
          <UploadCloud className="upload-icon" />
          <h3>Planilha TRANSFEREGOV (Referência)</h3>
          <p>{fileTransferegov ? fileTransferegov.name : "Clique ou arraste para selecionar"}</p>
        </div>

        <div className={`upload-card ${fileSiafi ? 'loaded' : ''}`}>
          <input
            type="file"
            className="file-input"
            accept=".xlsx, .xls, .csv"
            onChange={(e) => handleFileUpload(e, 'siafi')}
          />
          <UploadCloud className="upload-icon" />
          <h3>Planilha SIAFI (Alvo)</h3>
          <p>{fileSiafi ? fileSiafi.name : "Clique ou arraste para selecionar"}</p>
        </div>
      </div>

      <button
        className="btn-primary"
        onClick={processFiles}
        disabled={!fileTransferegov || !fileSiafi || isProcessing}
      >
        <Play size={20} />
        {isProcessing ? (processingPhase || "Processando Conciliação...") : "Iniciar Conciliação Automática"}
      </button>

      {errorMsg && (
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.6rem', background: 'rgba(239,68,68,0.12)', border: '1px solid rgba(239,68,68,0.4)', borderRadius: 8, padding: '0.75rem 1rem', marginBottom: '1rem', color: '#fca5a5', fontSize: '0.88rem' }}>
          <XCircle size={16} color="#fca5a5" style={{ flexShrink: 0 }} />
          {errorMsg}
        </div>
      )}

      {dashboard && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '1.25rem', margin: '1.5rem 0 0.5rem' }}>

          {/* ── Card 1: Transferegov — barras horizontais por status ── */}
          <div className="stat-card" style={{ padding: '1.25rem', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '1rem' }}>
              <BarChart2 size={16} color="#60a5fa" />
              <span style={{ fontWeight: 700, fontSize: '0.82rem', color: DASH.cardText }}>
                Transferegov — Volume por Status
              </span>
            </div>
            {dashboard.tgEntries.map(([st, n]) => (
              <div key={st} style={{ width: '100%', marginBottom: '0.6rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.72rem', marginBottom: 3, alignItems: 'center' }}>
                  <span style={{ maxWidth: '76%', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', color: DASH.cardMuted }}>
                    {st || '—'}
                  </span>
                  <span style={{ fontWeight: 800, color: DASH.numColor, fontSize: '0.78rem', background: DASH.numBadge, padding: '1px 7px', borderRadius: 99, flexShrink: 0 }}>{n}</span>
                </div>
                <div style={{ background: DASH.barTrack, borderRadius: 4, height: 8 }}>
                  <div style={{ width: `${Math.round((n / dashboard.tgMax) * 100)}%`, background: DASH.barBlue, height: 8, borderRadius: 4 }} />
                </div>
              </div>
            ))}
          </div>

          {/* ── Card 2: SIAFI — Distribuição de Contas (BI) ── */}
          <div className="stat-card" style={{ padding: '1.25rem', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '0.75rem' }}>
              <PieChart size={16} color="#a78bfa" />
              <span style={{ fontWeight: 700, fontSize: '0.82rem', color: DASH.cardText }}>
                SIAFI — Distribuição de Contas
              </span>
            </div>

            {/* KPI chips */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: '0.4rem', width: '100%', marginBottom: '0.75rem' }}>
              {[
                { label: 'Instrumentos', val: dashboard.uniqueInstruments },
                { label: 'Contas SIAFI', val: dashboard.allAccEntries.length },
                { label: 'Média Contas', val: (dashboard.allAccTotal / Math.max(dashboard.uniqueInstruments, 1)).toFixed(1) },
              ].map(k => (
                <div key={k.label} style={{ background: DASH.chipBg, borderRadius: 8, padding: '0.5rem 0.4rem', textAlign: 'center', border: '1px solid rgba(255,255,255,0.06)' }}>
                  <div style={{ fontSize: '1.15rem', fontWeight: 800, color: DASH.cardText, lineHeight: 1.1 }}>{k.val}</div>
                  <div style={{ fontSize: '0.6rem', color: DASH.cardSubtle, marginTop: 3, textTransform: 'uppercase', letterSpacing: '0.04em' }}>{k.label}</div>
                </div>
              ))}
            </div>

            {/* Donut */}
            <div style={{ display: 'flex', justifyContent: 'center', width: '100%', marginBottom: '0.7rem' }}>
              <svg width="118" height="118" viewBox="0 0 130 130">
                {dashboard.allSegments.length > 0
                  ? dashboard.allSegments.map((seg, i) => (
                      <circle key={i} cx="65" cy="65" r="50"
                        fill="none" stroke={seg.color} strokeWidth="18"
                        strokeDasharray={seg.dasharray}
                        strokeDashoffset={seg.dashoffset}
                        transform="rotate(-90 65 65)"
                      />
                    ))
                  : <circle cx="65" cy="65" r="50" fill="none" stroke={DASH.barTrack} strokeWidth="18" />
                }
                <text x="65" y="61" textAnchor="middle" fontSize="15" fontWeight="800" fill={DASH.cardText}>
                  {dashboard.allAccTotal}
                </text>
                <text x="65" y="75" textAnchor="middle" fontSize="8.5" fill={DASH.cardSubtle}>lançamentos</text>
              </svg>
            </div>

            {/* Barras horizontais por conta */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.42rem', width: '100%' }}>
              {dashboard.allAccEntries.map(d => {
                const short = d.fullLabel
                  .replace(/^Convênios e [Ii]nstrumentos [Cc]ongêneres\s+/i, '')
                  .replace(/^Convênios e instrumentos\s+/i, '');
                const pct = ((d.n / dashboard.allAccTotal) * 100).toFixed(1);
                return (
                  <div key={d.code}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.7rem', marginBottom: 3, alignItems: 'center' }}>
                      <span style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', color: DASH.cardMuted, maxWidth: '70%', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        <span style={{ width: 8, height: 8, borderRadius: '50%', background: d.color, display: 'inline-block', flexShrink: 0 }} />
                        {short}
                      </span>
                      <span style={{ fontWeight: 800, color: DASH.cardText, flexShrink: 0, fontSize: '0.72rem' }}>
                        {d.n} <span style={{ fontWeight: 400, color: DASH.cardSubtle }}>({pct}%)</span>
                      </span>
                    </div>
                    <div style={{ background: DASH.barTrack, borderRadius: 4, height: 7 }}>
                      <div style={{ width: `${Math.round((d.n / dashboard.accMax) * 100)}%`, background: d.color, height: 7, borderRadius: 4 }} />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* ── Card 3: Resultado da Auditoria — reage a filtros ── */}
          <div className="stat-card" style={{ padding: '1.25rem', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', width: '100%', marginBottom: '0.6rem' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <TrendingUp size={16} color="#34d399" />
                <span style={{ fontWeight: 700, fontSize: '0.82rem', color: DASH.cardText }}>Resultado da Auditoria</span>
              </div>
              {(filterNrInstrumento || filterSituacaoTg || filterContaSiafi || filterStatus || filterConvenente) && (
                <span style={{ fontSize: '0.63rem', fontWeight: 700, color: '#60a5fa', background: 'rgba(96,165,250,0.15)', border: '1px solid rgba(96,165,250,0.3)', borderRadius: 99, padding: '2px 8px' }}>
                  Filtrado
                </span>
              )}
            </div>

            {/* Donut de destaque — maior para ocupar o espaço */}
            <div style={{ display: 'flex', justifyContent: 'center', width: '100%', marginBottom: '0.7rem' }}>
              <svg width="170" height="170" viewBox="0 0 170 170">
                {filteredStats.segments.length > 0
                  ? filteredStats.segments.map((seg, i) => (
                      <circle key={i} cx="85" cy="85" r="62"
                        fill="none" stroke={seg.color} strokeWidth="22"
                        strokeDasharray={seg.dasharray}
                        strokeDashoffset={seg.dashoffset}
                        transform="rotate(-90 85 85)"
                      />
                    ))
                  : <circle cx="85" cy="85" r="62" fill="none" stroke={DASH.barTrack} strokeWidth="22" />
                }
                <text x="85" y="78" textAnchor="middle" fontSize="28" fontWeight="900" fill="#34d399">
                  {filteredStats.total > 0 ? `${Math.round((filteredStats.corretos / filteredStats.total) * 100)}%` : '—'}
                </text>
                <text x="85" y="94" textAnchor="middle" fontSize="10" fill={DASH.cardSubtle}>Correto</text>
                <text x="85" y="108" textAnchor="middle" fontSize="9.5" fontWeight="700" fill={DASH.cardMuted}>
                  {filteredStats.total} instrumentos
                </text>
              </svg>
            </div>

            {/* Barras por categoria */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', width: '100%' }}>
              {filteredStats.items.map(d => (
                <div key={d.label}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.72rem', marginBottom: 3, alignItems: 'center' }}>
                    <span style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', color: DASH.cardMuted }}>
                      <span style={{ width: 8, height: 8, borderRadius: '50%', background: d.color, display: 'inline-block', flexShrink: 0 }} />
                      {d.label}
                    </span>
                    <span style={{ fontWeight: 800, color: DASH.cardText, fontSize: '0.75rem' }}>
                      {d.val}&nbsp;<span style={{ fontWeight: 400, color: DASH.cardSubtle }}>
                        ({filteredStats.total > 0 ? ((d.val / filteredStats.total) * 100).toFixed(1) : '0.0'}%)
                      </span>
                    </span>
                  </div>
                  <div style={{ background: DASH.barTrack, borderRadius: 4, height: 8 }}>
                    <div style={{ width: `${filteredStats.total > 0 ? Math.round((d.val / filteredStats.total) * 100) : 0}%`, background: d.color, height: 8, borderRadius: 4 }} />
                  </div>
                </div>
              ))}
            </div>
          </div>

        </div>
      )}

      {results.length > 0 && (
        <div className="results-panel">
          <div className="results-header">
            <h2>Resultados da Auditoria ({stats.total} linhas)</h2>
            <button className="btn-primary" style={{ width: 'auto', marginBottom: 0 }} onClick={exportResults}>
              <Download size={18} /> Exportar Excel ({filteredResults.length} linhas)
            </button>
          </div>

          {/* ── Filtros ── */}
          <div style={{ display: 'flex', gap: '0.75rem', margin: '0.75rem 0', flexWrap: 'wrap' }}>
            <input
              type="text"
              placeholder="Nº Instrumento"
              value={filterNrInstrumento}
              onChange={e => { setFilterNrInstrumento(e.target.value); setCurrentPage(0); }}
              style={{ flex: '0 1 150px', minWidth: 120, padding: '0.45rem 0.75rem', border: '1px solid #334155', borderRadius: 8, fontSize: '0.82rem', outline: 'none', background: '#1e293b', color: '#f1f5f9' }}
            />
            <input
              type="text"
              placeholder="Convenente / CNPJ"
              value={filterConvenente}
              onChange={e => { setFilterConvenente(e.target.value); setCurrentPage(0); }}
              style={{ flex: '0 1 180px', minWidth: 150, padding: '0.45rem 0.75rem', border: '1px solid #334155', borderRadius: 8, fontSize: '0.82rem', outline: 'none', background: '#1e293b', color: '#f1f5f9' }}
            />
            <select
              value={filterSituacaoTg}
              onChange={e => { setFilterSituacaoTg(e.target.value); setCurrentPage(0); }}
              style={{ flex: '1 1 190px', minWidth: 170, padding: '0.45rem 0.75rem', border: '1px solid #334155', borderRadius: 8, fontSize: '0.82rem', background: '#1e293b', color: '#f1f5f9', outline: 'none' }}
            >
              <option value="">Todas as Situações (TG)</option>
              {uniqueTgSituacoes.map(s => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
            <select
              value={filterContaSiafi}
              onChange={e => { setFilterContaSiafi(e.target.value); setCurrentPage(0); }}
              style={{ flex: '1 1 200px', minWidth: 180, padding: '0.45rem 0.75rem', border: '1px solid #334155', borderRadius: 8, fontSize: '0.82rem', background: '#1e293b', color: '#f1f5f9', outline: 'none' }}
            >
              <option value="">Todas as Contas (SIAFI)</option>
              {uniqueContasSiafi.map(({ code, name }) => (
                <option key={code} value={code}>{name}</option>
              ))}
            </select>
            <select
              value={filterStatus}
              onChange={e => { setFilterStatus(e.target.value); setCurrentPage(0); }}
              style={{ flex: '1 1 170px', minWidth: 150, padding: '0.45rem 0.75rem', border: '1px solid #334155', borderRadius: 8, fontSize: '0.82rem', background: '#1e293b', color: '#f1f5f9', outline: 'none' }}
            >
              <option value="">Todos os Status</option>
              <option value="Correto">Correto (Rito Ordinário)</option>
              <option value="Inconsistência (Rito Patológico)">Inconsistência (Rito Patológico)</option>
              <option value="Sem Registro no Transferegov">Sem Registro no Transferegov</option>
              <option value="Regra Pendente - Revisar">Regra Pendente - Revisar</option>
              <option value="Sem Saldo no SIAFI">Sem Saldo no SIAFI</option>
              <option value="Revisão Manual - ID Ambíguo">Revisão Manual - ID Ambíguo</option>
              <option value="Status Não Mapeado">Status Não Mapeado</option>
              <option value="Atenção — Conclusão SIAFI sem encerramento no TG">Conclusão SIAFI sem encerramento no TG</option>
              <option value="Atenção — Anulação sem estorno contábil no SIAFI">Anulação sem estorno contábil no SIAFI</option>
              <option value="Atenção — Aprovação sem encerramento no SIAFI">Aprovação sem encerramento no SIAFI (812210103)</option>
              <option value="Inconsistência — Inadimplência na fase A Aprovar">Inadimplência na fase A Aprovar (812210103)</option>
              <option value="Revisão Manual — Sem Rastreabilidade de Evento">Sem Rastreabilidade de Evento (812210103)</option>
            </select>
            {(filterNrInstrumento || filterSituacaoTg || filterContaSiafi || filterStatus || filterConvenente) && (
              <button
                onClick={() => { setFilterNrInstrumento(''); setFilterSituacaoTg(''); setFilterContaSiafi(''); setFilterStatus(''); setFilterConvenente(''); setCurrentPage(0); }}
                style={{ padding: '0.45rem 0.9rem', border: '1px solid #334155', borderRadius: 8, fontSize: '0.82rem', background: '#1e293b', cursor: 'pointer', color: '#94a3b8' }}
              >
                Limpar filtros
              </button>
            )}
          </div>

          {/* ── Contador ── */}
          <p style={{ fontSize: '0.78rem', color: '#64748b', margin: '0 0 0.5rem', fontWeight: 500 }}>
            Exibindo {new Set(filteredResults.map(r => r.idSiafi)).size} de {new Set(results.map(r => r.idSiafi)).size} instrumentos encontrados
          </p>

          <div className="stats-grid stats-grid-5">
            <div className="stat-card">
              <div className="stat-value text-accent">{stats.total}</div>
              <div className="stat-label">Total SIAFI</div>
            </div>
            <div className="stat-card">
              <div className="stat-value text-success">{stats.corretos}</div>
              <div className="stat-label">Corretos</div>
            </div>
            <div className="stat-card">
              <div className="stat-value text-danger">{stats.inconsistencias}</div>
              <div className="stat-label">Inconsistências</div>
            </div>
            <div className="stat-card">
              <div className="stat-value text-warning">{stats.naoEncontrados}</div>
              <div className="stat-label">Sem Ref. Transferegov</div>
            </div>
            <div className="stat-card">
              <div className="stat-value text-orange">{stats.alertas}</div>
              <div className="stat-label">Alertas - Revisar</div>
            </div>
          </div>

          <div className="table-container">
            <table>
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Convenente</th>
                  <th>CNPJ</th>
                  <th>Valor (R$)</th>
                  <th>Situação Transferegov</th>
                  <th>Situação SIAFI</th>
                  <th>Status Conciliação</th>
                  <th>Motivo</th>
                </tr>
              </thead>
              <tbody>
                {paginatedResults.map((row) => {
                  const motivo = row.fullData["Motivo do Alerta"] || '—';
                  return (
                    <tr key={row.idSiafi}>
                      <td>{row.idSiafi}</td>
                      <td title={row.convenenteSiafi || row.convenenteNome}>
                        {(() => { const c = String(row.convenenteSiafi || row.convenenteNome || ''); return c.length > 24 ? c.substring(0, 24) + '…' : c || '—'; })()}
                      </td>
                      <td>{row.cnpjSiafi || '—'}</td>
                      <td style={{ whiteSpace: 'nowrap' }}>{row.valorSiafi || '—'}</td>
                      <td>{String(row.situacaoRawTg ?? '—')}</td>
                      <td title={String(row.situacaoSiafiDisplay ?? '')}>
                        {(() => { const s = String(row.situacaoSiafiDisplay ?? ''); return s.length > 38 ? s.substring(0, 38) + '…' : s || '—'; })()}
                      </td>
                      <td>
                        <span className={`status-badge ${getStatusBadgeClass(String(row.statusConciliacao ?? ''))}`}>
                          {getStatusIcon(String(row.statusConciliacao ?? ''))}
                          {String(row.statusConciliacao ?? '—')}
                        </span>
                      </td>
                      <td title={motivo} style={{ color: 'var(--text-muted)', fontSize: '0.82rem', maxWidth: 220 }}>
                        {motivo.length > 55 ? motivo.substring(0, 55) + '…' : motivo}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {filteredResults.length > PAGE_SIZE && (
              <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '1rem', marginTop: '1rem', fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                <button
                  onClick={() => setCurrentPage(p => Math.max(0, p - 1))}
                  disabled={currentPage === 0}
                  style={{ padding: '0.35rem 0.9rem', border: '1px solid #334155', borderRadius: 6, background: currentPage === 0 ? 'transparent' : '#1e293b', color: currentPage === 0 ? '#475569' : '#94a3b8', cursor: currentPage === 0 ? 'default' : 'pointer', fontSize: '0.82rem' }}
                >
                  ← Anterior
                </button>
                <span>
                  Página {currentPage + 1} de {Math.ceil(filteredResults.length / PAGE_SIZE)}
                  <span style={{ color: '#475569', marginLeft: '0.5rem' }}>({filteredResults.length} registros)</span>
                </span>
                <button
                  onClick={() => setCurrentPage(p => Math.min(Math.ceil(filteredResults.length / PAGE_SIZE) - 1, p + 1))}
                  disabled={currentPage >= Math.ceil(filteredResults.length / PAGE_SIZE) - 1}
                  style={{ padding: '0.35rem 0.9rem', border: '1px solid #334155', borderRadius: 6, background: currentPage >= Math.ceil(filteredResults.length / PAGE_SIZE) - 1 ? 'transparent' : '#1e293b', color: currentPage >= Math.ceil(filteredResults.length / PAGE_SIZE) - 1 ? '#475569' : '#94a3b8', cursor: currentPage >= Math.ceil(filteredResults.length / PAGE_SIZE) - 1 ? 'default' : 'pointer', fontSize: '0.82rem' }}
                >
                  Próxima →
                </button>
              </div>
            )}
          </div>
        </div>
      )}
    </div>

    {/* ── Modais Informativos ── */}
    {modalOpen && (
      <div
        onClick={() => setModalOpen(null)}
        style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.65)', zIndex: 1000,
          display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '1.5rem',
        }}
      >
        <div
          onClick={e => e.stopPropagation()}
          style={{
            background: '#1e293b', borderRadius: 14, maxWidth: 580, width: '100%',
            boxShadow: '0 20px 60px rgba(0,0,0,0.5)', border: '1px solid #334155',
            fontFamily: 'Arial, sans-serif',
          }}
        >
          {/* Header */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '1.25rem 1.5rem', borderBottom: '1px solid #334155' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
              {modalOpen === 'guide'
                ? <Settings size={20} color="#60a5fa" />
                : <Shield size={20} color="#34d399" />
              }
              <span style={{ fontWeight: 700, fontSize: '1rem', color: '#f1f5f9' }}>
                {modalOpen === 'guide' ? 'Guia de Operação e Metas' : 'Conformidade e Avaliação Digital'}
              </span>
            </div>
            <button
              onClick={() => setModalOpen(null)}
              style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8', display: 'flex', alignItems: 'center', padding: 4, borderRadius: 6 }}
            >
              <X size={20} />
            </button>
          </div>

          {/* Body */}
          <div style={{ padding: '1.5rem', color: '#cbd5e1', fontSize: '0.875rem', lineHeight: 1.7 }}>
            {modalOpen === 'guide' ? (
              <>
                <p style={{ color: '#94a3b8', marginBottom: '1.25rem' }}>
                  O SIACT Hub é uma ferramenta de auditoria digital projetada para a conciliação automatizada entre as bases de dados Transferegov e SIAFI (Tesouro Gerencial).
                </p>
                <p style={{ fontWeight: 700, color: '#e2e8f0', marginBottom: '0.5rem' }}>Funcionalidades</p>
                <ul style={{ margin: '0 0 1.25rem 1.1rem', padding: 0 }}>
                  <li>Cruzamento de dados estruturados em tempo real.</li>
                  <li>Identificação automática de divergências de situação contábil.</li>
                  <li>Aplicação de hierarquia de eventos SIAFI (P1–P5 + motor P0 para conta 812210103).</li>
                  <li>Motor 812210103 (v0.3.0): aprovação sem encerramento, inadimplência na fase PC e sem rastreabilidade de evento.</li>
                  <li>Geração de relatórios de conformidade em Excel.</li>
                </ul>
                <p style={{ fontWeight: 700, color: '#e2e8f0', marginBottom: '0.5rem' }}>Metas de Desempenho</p>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem', marginBottom: '1.25rem' }}>
                  {[
                    { label: 'Redução do tempo de conciliação manual', value: '80%' },
                    { label: 'Acurácia no processamento de registros vinculados', value: '100%' },
                  ].map(m => (
                    <div key={m.label} style={{ background: '#0f172a', borderRadius: 10, padding: '0.75rem 1rem', border: '1px solid #334155' }}>
                      <div style={{ fontSize: '1.5rem', fontWeight: 800, color: '#60a5fa' }}>{m.value}</div>
                      <div style={{ fontSize: '0.75rem', color: '#64748b', marginTop: 2 }}>{m.label}</div>
                    </div>
                  ))}
                </div>
                <p style={{ color: '#94a3b8', fontSize: '0.82rem' }}>
                  <strong style={{ color: '#e2e8f0' }}>Conclusão esperada:</strong> Otimizar a fiscalização financeira e garantir a integridade dos dados auditados pelo MTur.
                </p>
              </>
            ) : (
              <>
                <p style={{ color: '#94a3b8', marginBottom: '1.25rem' }}>
                  Este sistema foi desenvolvido em estrita observância aos pilares da Estratégia de Governo Digital (SGD/MGI).
                </p>
                <p style={{ fontWeight: 700, color: '#e2e8f0', marginBottom: '0.75rem' }}>Diretrizes da Secretaria de Governo Digital</p>
                {[
                  { icon: '🔗', title: 'Interoperabilidade', desc: 'Comunicação transparente entre Transferegov e SIAFI, sem duplicidade de esforços.' },
                  { icon: '🔒', title: 'Segurança da Informação', desc: 'Processamento 100% client-side — nenhum dado é transmitido para servidores externos.' },
                  { icon: '📋', title: 'Transparência Ativa', desc: 'Estruturação de dados que facilita a prestação de contas e o controle social.' },
                ].map(item => (
                  <div key={item.title} style={{ display: 'flex', gap: '0.75rem', marginBottom: '0.9rem', alignItems: 'flex-start' }}>
                    <span style={{ fontSize: '1.1rem', marginTop: 1 }}>{item.icon}</span>
                    <div>
                      <div style={{ fontWeight: 700, color: '#e2e8f0', fontSize: '0.85rem' }}>{item.title}</div>
                      <div style={{ color: '#94a3b8', fontSize: '0.8rem' }}>{item.desc}</div>
                    </div>
                  </div>
                ))}
                <div style={{ marginTop: '1.25rem', background: '#0f172a', borderRadius: 10, padding: '0.85rem 1rem', border: '1px solid #334155' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '0.4rem' }}>
                    <ShieldCheck size={14} color="#34d399" />
                    <span style={{ fontWeight: 700, fontSize: '0.82rem', color: '#34d399' }}>Avaliação AIE — Risco Baixo</span>
                  </div>
                  <p style={{ fontSize: '0.78rem', color: '#64748b', margin: 0 }}>
                    O sistema atende aos critérios da Portaria SGD/MGI nº 473/2026, com supervisão humana obrigatória em todas as decisões e processamento efêmero sem armazenamento de dados.
                  </p>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    )}

    <footer style={{ textAlign: 'center', padding: '2rem 0 1rem', color: '#475569', fontSize: '0.72rem', borderTop: '1px solid #1f2937', marginTop: '2rem' }}>
      SIACT Hub · v0.3.0 · Build 2026-05-15 · MTur — Coordenação de Análise Financeira de Prestação de Contas
    </footer>

    </ErrorBoundary>
  );
}
