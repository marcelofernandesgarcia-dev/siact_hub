import { useState, useMemo, Component, type ReactNode } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, CheckCircle, XCircle, AlertCircle, Play, Download, BarChart2, PieChart, ShieldCheck, TrendingUp } from 'lucide-react';
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
  "Em execução":                                   ["811210100", "812210101", "812210201", "811210103"],
  // Aguardando P.C. — foco na obrigação de comprovar o recurso recebido
  "Aguardando prestação de contas":                ["812210202", "811210102"],
  // P.C. em Análise — transição entre entrega dos documentos e o parecer técnico
  "Prestação de Contas em Análise":                ["812210103", "812210104", "812210202", "811210102"],
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
  '812210103': 'Convênios e Instrumentos Congêneres a Aprovar',
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

// SKILL §2.1: ID must be exactly 6 digits starting with 7 (700000–799999)
const extractTransferId = (val: string | number | undefined): string | null => {
  if (!val) return null;
  const str = String(val).trim();
  const match = str.match(/(?:^|\D)(7\d{5})(?:\D|$)/);
  return match ? match[1] : null;
};

// SKILL §5.1: Contas Críticas — ordem conforme roteiro
const CRITICAL_ACCS = [
  { code: "812210211", label: "Concluído",   color: "#10b981" },
  { code: "812210109", label: "Anulado",     color: "#ef4444" },
  { code: "812210104", label: "Aprovado",    color: "#3b82f6" },
  { code: "811210109", label: "Extinto",     color: "#8b5cf6" },
  { code: "811210100", label: "Em Execução", color: "#f59e0b" },
];

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
  const [results, setResults] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [stats, setStats] = useState({
    total: 0, corretos: 0, inconsistencias: 0, naoEncontrados: 0, alertas: 0,
  });
  const [filterNrInstrumento, setFilterNrInstrumento] = useState('');
  const [filterSituacaoSiafi, setFilterSituacaoSiafi] = useState('');
  const [filterStatus, setFilterStatus] = useState('');

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
      const m = String(r.situacaoSiafiDisplay || "").match(/^(\d{9})/);
      if (m) {
        if (!accUniqueIds[m[1]]) accUniqueIds[m[1]] = new Set();
        accUniqueIds[m[1]].add(r.idSiafi);
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

    const donutSlices = CRITICAL_ACCS
      .map(a => ({ ...a, n: accCounts[a.code] || 0 }))
      .filter(a => a.n > 0);
    const donutTotal = donutSlices.reduce((s, a) => s + a.n, 0) || 1;

    const C = 2 * Math.PI * 50;
    let off = 0;
    const segments = donutSlices.map(d => {
      const frac = d.n / donutTotal;
      const seg = {
        ...d,
        dasharray: `${(frac * C).toFixed(2)} ${C.toFixed(2)}`,
        dashoffset: -(off * C),
      };
      off += frac;
      return seg;
    });

    // Count accounting entries per unique instrument ID
    const entriesPerIdMap = new Map<string, number>();
    for (const r of results) {
      entriesPerIdMap.set(r.idSiafi, (entriesPerIdMap.get(r.idSiafi) || 0) + 1);
    }
    const entryVals = [...entriesPerIdMap.values()];
    const singleEntry   = entryVals.filter(n => n === 1).length;
    const multipleEntry = entryVals.filter(n => n > 1).length;

    return { tgEntries, tgMax, segments, donutSlices, donutTotal,
             uniqueInstruments: uniqueIds.size, totalEntries: results.length,
             singleEntry, multipleEntry };
  }, [results]);

  const filteredResults = useMemo(() => {
    const idTrim = filterNrInstrumento.trim();
    const siafTrim = filterSituacaoSiafi.trim().toLowerCase();
    return results.filter(r => {
      if (idTrim && !String(r.idSiafi ?? '').includes(idTrim)) return false;
      if (siafTrim && !String(r.situacaoSiafiDisplay ?? '').toLowerCase().includes(siafTrim)) return false;
      if (filterStatus && r.statusConciliacao !== filterStatus) return false;
      return true;
    });
  }, [results, filterNrInstrumento, filterSituacaoSiafi, filterStatus]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: 'transferegov' | 'siafi') => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (type === 'transferegov') setFileTransferegov(file);
    else setFileSiafi(file);
  };

  const processFiles = async () => {
    if (!fileTransferegov || !fileSiafi) return;
    setIsProcessing(true);

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

      if (rowsTg.length === 0 || rowsSiafi.length === 0) {
        alert("Uma das planilhas está vazia!");
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
        if (idColTgIdx !== -1 && extractTransferId(row[idColTgIdx])) {
          id = extractTransferId(row[idColTgIdx])!;
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

      // --- 2. Process SIAFI (formato normalizado — novo formato MTur 2026) ---
      // Cada linha = instrumento × conta contábil × evento
      // Fase 1: Localizar colunas pelos cabeçalhos
      let siafiHeaderRowIdx = -1;
      let contaColIdx      = -1;  // "Conta Contábil" → código 9 dígitos como VALOR
      let valorColIdx      = -1;  // "Transferência - Valor" → valor de referência
      let eventoColIdx     = -1;  // "Evento" → código do evento contábil
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
        alert("Não foi possível identificar a coluna 'Conta Contábil' na planilha SIAFI. Verifique o formato do arquivo.");
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
        if (raw.length === 6 && raw[0] === '7') {
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
        if (!contaRaw.match(/^\d{9}$/)) continue;

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
          instrumentMap.set(idSiafi, { contas: new Map(), cnpj: lastCnpjSiafi, convenente: lastConvenenteSiafi });
        }
        const entry = instrumentMap.get(idSiafi)!;
        if (!entry.cnpj && lastCnpjSiafi) entry.cnpj = lastCnpjSiafi;
        if (!entry.convenente && lastConvenenteSiafi) entry.convenente = lastConvenenteSiafi;

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
      const finalResults: any[] = [];

      for (const [idSiafi, entry] of instrumentMap) {
        const detectedAccounts = [...entry.contas.entries()].map(([code, info]) => ({
          code,
          display: info.display,
          eventos: info.eventos,
        }));

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
        const situacaoRawSiafi = situacaoSiafiDisplay;

        let situacaoTg = "Sem Registro no Transferegov";
        let situacaoRawTg = "Sem Registro no Transferegov";
        let convenenteNome = convenenteSiafi || "Não Informado";
        let statusConciliacao = "";
        let contasForaDoConjunto: string[] = [];

        if (tgMap.has(idSiafi) && !tgMap.get(idSiafi)!.ambiguous) {
          const tgEntry = tgMap.get(idSiafi)!;
          situacaoRawTg = tgEntry.status || String(tgEntry.fullRow[15] ?? "").trim();
        }

        if (tgMap.has(idSiafi)) {
          const tgData = tgMap.get(idSiafi)!;
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
          convenenteNome,
          convenenteSiafi,
          cnpjSiafi,
          situacaoRawTg,
          situacaoRawSiafi,
          situacaoTg,
          situacaoSiafiDisplay,
          statusConciliacao,
          fullData: { "Status de Conciliação": statusConciliacao, "Motivo do Alerta": motivo },
        });
      }

      setStats({ total: finalResults.length, corretos, inconsistencias, naoEncontrados, alertas });
      setResults(finalResults);
    } catch (err) {
      console.error(err);
      alert("Erro ao processar planilhas.");
    }

    setIsProcessing(false);
  };

  // SKILL §3 — Exportação com esquema fixo A-F (Validação de Colunas)
  // A: ID (6 dígitos)  B: Convenente  C: CNPJ  D: Situação Transferegov
  // E: Situação SIAFI (nomes por extenso do accountMap)  F: Status (Conciliação)
  // Carry-forward de CNPJ/Convenente garantido pelos campos r.cnpjSiafi / r.convenenteSiafi.
  const exportResults = () => {
    const dataToExport = filteredResults.map(r => {
      // Coluna E: se há múltiplas contas, o display já está concatenado com ' | '
      // Garantir que cada segmento use o nome do accountMap (já aplicado na construção)
      const situacaoSiafi = r.situacaoSiafiDisplay;

      return {
        // A — ID do Instrumento (6 dígitos, início 7)
        "ID": r.idSiafi,
        // B — Convenente com carry-forward (nunca vazio)
        "Convenente": r.convenenteSiafi || r.convenenteNome || "",
        // C — CNPJ com carry-forward (nunca vazio)
        "CNPJ": r.cnpjSiafi || "",
        // D — Situação Transferegov (Col P do arquivo TG)
        "Situação Transferegov": r.situacaoRawTg,
        // E — Situação SIAFI: [Código] - [Nome] com separador ' | ' para múltiplas contas
        "Situação SIAFI": situacaoSiafi,
        // F — Status da Conciliação
        "Status (Conciliação)": r.statusConciliacao,
      };
    });
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
         'Atenção — Anulação sem estorno contábil no SIAFI'].includes(status)) return 'status-alert';
    return 'status-incorrect'; // Inconsistência (Rito Patológico) e outros
  };

  const getStatusIcon = (status: string) => {
    const style = { display: 'inline' as const, marginRight: '4px', verticalAlign: 'text-bottom' as const };
    if (status === 'Correto') return <CheckCircle size={14} style={style} />;
    if (['Regra Pendente - Revisar', 'Sem Saldo no SIAFI',
         'Revisão Manual - ID Ambíguo', 'Status Não Mapeado',
         'Sem Registro no Transferegov',
         'Atenção — Conclusão SIAFI sem encerramento no TG',
         'Atenção — Anulação sem estorno contábil no SIAFI'].includes(status)) return <AlertCircle size={14} style={style} />;
    return <XCircle size={14} style={style} />; // Inconsistência (Rito Patológico)
  };

  return (
    <ErrorBoundary>
    <div className="app-container" translate="no">
      <div className="header">
        <h1>SIACT Hub</h1>
        <p>Auditor Digital MTur — Conciliação Transferegov × SIAFI (Tesouro Gerencial)</p>
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
        {isProcessing ? "Processando Conciliação..." : "Iniciar Conciliação Automática"}
      </button>

      {dashboard && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '1.25rem', margin: '1.5rem 0 0.5rem' }}>

          {/* ── Card 1: Transferegov — barras horizontais por status ── */}
          <div className="stat-card" style={{ padding: '1.25rem', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '1rem' }}>
              <BarChart2 size={16} color="#3b82f6" />
              <span style={{ fontWeight: 700, fontSize: '0.82rem', color: '#1e293b' }}>
                Transferegov — Volume por Status
              </span>
            </div>
            {dashboard.tgEntries.map(([st, n]) => (
              <div key={st} style={{ width: '100%', marginBottom: '0.55rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.7rem', marginBottom: 2 }}>
                  <span style={{ maxWidth: '78%', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', color: '#475569' }}>
                    {st || '—'}
                  </span>
                  <span style={{ fontWeight: 700, color: '#1e293b' }}>{n}</span>
                </div>
                <div style={{ background: '#e2e8f0', borderRadius: 3, height: 6 }}>
                  <div style={{ width: `${Math.round((n / dashboard.tgMax) * 100)}%`, background: '#3b82f6', height: 6, borderRadius: 3 }} />
                </div>
              </div>
            ))}
          </div>

          {/* ── Card 2: SIAFI — donut das contas críticas ── */}
          <div className="stat-card" style={{ padding: '1.25rem', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '1rem' }}>
              <PieChart size={16} color="#8b5cf6" />
              <span style={{ fontWeight: 700, fontSize: '0.82rem', color: '#1e293b' }}>
                SIAFI — Contas Críticas
              </span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'center', width: '100%', marginBottom: '0.75rem' }}>
              <svg width="130" height="130" viewBox="0 0 130 130">
                {dashboard.segments.length > 0
                  ? dashboard.segments.map((seg, i) => (
                      <circle key={i} cx="65" cy="65" r="50"
                        fill="none" stroke={seg.color} strokeWidth="20"
                        strokeDasharray={seg.dasharray}
                        strokeDashoffset={seg.dashoffset}
                        transform="rotate(-90 65 65)"
                      />
                    ))
                  : <circle cx="65" cy="65" r="50" fill="none" stroke="#e2e8f0" strokeWidth="20" />
                }
                <text x="65" y="61" textAnchor="middle" fontSize="15" fontWeight="800" fill="#1e293b">
                  {dashboard.donutTotal}
                </text>
                <text x="65" y="75" textAnchor="middle" fontSize="9" fill="#64748b">lançamentos</text>
              </svg>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.3rem', width: '100%' }}>
              {dashboard.donutSlices.map(d => (
                <div key={d.code} style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.7rem', alignItems: 'center' }}>
                  <span style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', color: '#475569' }}>
                    <span style={{ width: 9, height: 9, borderRadius: '50%', background: d.color, display: 'inline-block', flexShrink: 0 }} />
                    {d.label}
                  </span>
                  <span style={{ fontWeight: 700, color: '#1e293b' }}>
                    {d.n}&nbsp;<span style={{ fontWeight: 400, color: '#94a3b8' }}>({((d.n / dashboard.donutTotal) * 100).toFixed(1)}%)</span>
                  </span>
                </div>
              ))}
            </div>
          </div>

          {/* ── Card 3: Integridade dos Dados ── */}
          <div className="stat-card" style={{ padding: '1.25rem', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '1rem' }}>
              <ShieldCheck size={16} color="#10b981" />
              <span style={{ fontWeight: 700, fontSize: '0.82rem', color: '#1e293b' }}>
                Integridade dos Dados
              </span>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.9rem', width: '100%' }}>

              {/* Counters row */}
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.5rem' }}>
                <div style={{ background: '#f8fafc', borderRadius: 8, padding: '0.6rem' }}>
                  <div style={{ fontSize: '0.62rem', color: '#64748b', marginBottom: 2 }}>Instrumentos Únicos</div>
                  <div style={{ fontSize: '1.6rem', fontWeight: 800, color: '#1e293b', lineHeight: 1 }}>
                    {dashboard.uniqueInstruments}
                  </div>
                </div>
                <div style={{ background: '#f8fafc', borderRadius: 8, padding: '0.6rem' }}>
                  <div style={{ fontSize: '0.62rem', color: '#64748b', marginBottom: 2 }}>Lançamentos (SIAFI)</div>
                  <div style={{ fontSize: '1.6rem', fontWeight: 800, color: '#1e293b', lineHeight: 1 }}>
                    {dashboard.totalEntries}
                  </div>
                </div>
              </div>

              {/* Complexity bar chart: 1 entry vs multiple */}
              <div style={{ width: '100%' }}>
                <div style={{ fontSize: '0.68rem', fontWeight: 700, color: '#475569', marginBottom: '0.55rem', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                  Complexidade por Instrumento
                </div>
                <div style={{ marginBottom: '0.45rem' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.7rem', marginBottom: 2 }}>
                    <span style={{ color: '#475569' }}>Simples (1 lançamento)</span>
                    <span style={{ fontWeight: 700, color: '#15803d' }}>{dashboard.singleEntry}</span>
                  </div>
                  <div style={{ background: '#e2e8f0', borderRadius: 3, height: 6 }}>
                    <div style={{ width: `${dashboard.uniqueInstruments ? Math.round((dashboard.singleEntry / dashboard.uniqueInstruments) * 100) : 0}%`, background: '#10b981', height: 6, borderRadius: 3 }} />
                  </div>
                </div>
                <div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.7rem', marginBottom: 2 }}>
                    <span style={{ color: '#475569' }}>Complexos (2+ lançamentos)</span>
                    <span style={{ fontWeight: 700, color: '#f97316' }}>{dashboard.multipleEntry}</span>
                  </div>
                  <div style={{ background: '#e2e8f0', borderRadius: 3, height: 6 }}>
                    <div style={{ width: `${dashboard.uniqueInstruments ? Math.round((dashboard.multipleEntry / dashboard.uniqueInstruments) * 100) : 0}%`, background: '#f97316', height: 6, borderRadius: 3 }} />
                  </div>
                </div>
              </div>

              {/* Inflation signal */}
              {dashboard.multipleEntry > 0 ? (
                <div style={{ background: '#fff7ed', border: '1px solid #fed7aa', borderRadius: 8, padding: '0.65rem' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', marginBottom: '0.2rem' }}>
                    <TrendingUp size={11} color="#f97316" />
                    <span style={{ fontSize: '0.68rem', fontWeight: 700, color: '#f97316' }}>
                      {dashboard.multipleEntry} instrumento{dashboard.multipleEntry > 1 ? 's' : ''} com múltiplos lançamentos
                    </span>
                  </div>
                  <span style={{ fontSize: '0.65rem', color: '#9a3412' }}>
                    +{dashboard.totalEntries - dashboard.uniqueInstruments} lançamentos excedentes — revisar células mescladas no SIAFI.
                  </span>
                </div>
              ) : (
                <div style={{ background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 8, padding: '0.65rem' }}>
                  <span style={{ fontSize: '0.7rem', color: '#15803d', fontWeight: 600 }}>✓ 1 lançamento por instrumento — sem inflação</span>
                </div>
              )}
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
              placeholder="Nº Instrumento (ex: 712345)"
              value={filterNrInstrumento}
              onChange={e => setFilterNrInstrumento(e.target.value)}
              style={{ flex: '1 1 160px', minWidth: 140, padding: '0.45rem 0.75rem', border: '1px solid #cbd5e1', borderRadius: 8, fontSize: '0.82rem', outline: 'none' }}
            />
            <input
              type="text"
              placeholder="Situação SIAFI (ex: Aprovado)"
              value={filterSituacaoSiafi}
              onChange={e => setFilterSituacaoSiafi(e.target.value)}
              style={{ flex: '2 1 200px', minWidth: 180, padding: '0.45rem 0.75rem', border: '1px solid #cbd5e1', borderRadius: 8, fontSize: '0.82rem', outline: 'none' }}
            />
            <select
              value={filterStatus}
              onChange={e => setFilterStatus(e.target.value)}
              style={{ flex: '1 1 180px', minWidth: 160, padding: '0.45rem 0.75rem', border: '1px solid #cbd5e1', borderRadius: 8, fontSize: '0.82rem', background: '#fff', outline: 'none' }}
            >
              <option value="">Todos os Status</option>
              <option value="Correto">✅ Correto (Rito Ordinário)</option>
              <option value="Inconsistência (Rito Patológico)">⛔ Inconsistência (Rito Patológico)</option>
              <option value="Sem Registro no Transferegov">🔍 Sem Registro no Transferegov</option>
              <option value="Regra Pendente - Revisar">⏳ Regra Pendente - Revisar</option>
              <option value="Sem Saldo no SIAFI">⚠️ Sem Saldo no SIAFI</option>
              <option value="Revisão Manual - ID Ambíguo">🔀 Revisão Manual - ID Ambíguo</option>
              <option value="Status Não Mapeado">❓ Status Não Mapeado</option>
              <option value="Atenção — Conclusão SIAFI sem encerramento no TG">⚠️ Conclusão SIAFI sem encerramento no TG</option>
              <option value="Atenção — Anulação sem estorno contábil no SIAFI">⚠️ Anulação sem estorno contábil no SIAFI</option>
            </select>
            {(filterNrInstrumento || filterSituacaoSiafi || filterStatus) && (
              <button
                onClick={() => { setFilterNrInstrumento(''); setFilterSituacaoSiafi(''); setFilterStatus(''); }}
                style={{ padding: '0.45rem 0.9rem', border: '1px solid #e2e8f0', borderRadius: 8, fontSize: '0.82rem', background: '#f8fafc', cursor: 'pointer', color: '#64748b' }}
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
                  <th>ID (A)</th>
                  <th>Convenente (B)</th>
                  <th>CNPJ (C)</th>
                  <th>Situação Transferegov (D)</th>
                  <th>Situação SIAFI (E)</th>
                  <th>Status Conciliação (F)</th>
                </tr>
              </thead>
              <tbody>
                {filteredResults.slice(0, 100).map((row, idx) => (
                  <tr key={idx}>
                    {/* A — ID */}
                    <td>{row.idSiafi}</td>
                    {/* B — Convenente (carry-forward) */}
                    <td title={row.convenenteSiafi || row.convenenteNome}>
                      {(() => {
                        const conv = String(row.convenenteSiafi || row.convenenteNome || '');
                        return conv.length > 28 ? conv.substring(0, 28) + '…' : conv || '—';
                      })()}
                    </td>
                    {/* C — CNPJ (carry-forward) */}
                    <td>{row.cnpjSiafi || '—'}</td>
                    {/* D — Situação Transferegov */}
                    <td>{String(row.situacaoRawTg ?? '—')}</td>
                    {/* E — Situação SIAFI (nomes por extenso do accountMap) */}
                    <td title={String(row.situacaoSiafiDisplay ?? '')}>
                      {(() => {
                        const siafi = String(row.situacaoSiafiDisplay ?? '');
                        return siafi.length > 45 ? siafi.substring(0, 45) + '…' : siafi || '—';
                      })()}
                    </td>
                    {/* F — Status Conciliação */}
                    <td>
                      <span className={`status-badge ${getStatusBadgeClass(String(row.statusConciliacao ?? ''))}`}>
                        {getStatusIcon(String(row.statusConciliacao ?? ''))}
                        {String(row.statusConciliacao ?? '—')}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {filteredResults.length > 100 && (
              <p style={{ textAlign: 'center', marginTop: '1rem', color: 'var(--text-muted)' }}>
                Mostrando as primeiras 100 linhas de {filteredResults.length}. Use o botão Exportar para obter o conjunto completo filtrado.
              </p>
            )}
          </div>
        </div>
      )}
    </div>
    </ErrorBoundary>
  );
}
