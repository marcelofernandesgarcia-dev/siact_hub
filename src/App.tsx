import { useState } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, CheckCircle, XCircle, AlertCircle, Play, Download } from 'lucide-react';
import './index.css';

// Mapping rules based on the user's provided table
const validationRules: Record<string, string[]> = {
  "Aguardando prestação de contas": ["812210202", "811210102"],
  "Cancelado": ["812210108", "811210106"],
  "Convênio Anulado": ["812210109"],
  "Convenio Rescindido": ["811210109"],
  "Em execução": ["811210100", "812210101", "812210201", "811210103"],
  "Prestação de Contas em Complementação": [],
  "Inadimplente": ["812210106"],
  "Prestação de Contas Aprovada": ["812210104"],
  "Prestação de Contas Aprovada com Ressalvas": [],
  "Prestação de Contas Comprovada em Análise": ["812210103", "812210105", "812210107"],
  "Prestação de Contas Concluída": ["812210211"],
  "Prestação de Contas em Análise": ["812210203"],
  "Prestação de Contas Iniciada por Antecipação": [],
  "Prestação de Contas Rejeitada": ["812210106"],
  "Proposta de Plano de Trabalho Aprovado": ["712210101"],
  "Proposta/Plano de Trabalho Aprovado": ["712210101"]
};

const accountNames: Record<string, string> = {
  "812210202": "A Comprovar",
  "811210102": "Convenios e Instrumentos Congeneres a Comprovar",
  "812210108": "Convênios e Instrumentos Congeneres Cancelados",
  "811210106": "Convenios e Instrumentos Congeneres não recebidos", 
  "812210109": "Convênios e Instrumentos Congeneres Não Liberado/ Devolvido",
  "811210109": "Convenios e Instrumentos Congeneres Extintos",
  "811210100": "Execução Convenio e Instrumentos Congeneres",
  "812210101": "Convenios e instrumentos a Liberar",
  "812210201": "A repassar",
  "811210103": "Convenios e Instrumentos Congeneres a Receber",
  "812210106": "Convenios e instrumentos Congeneres em Inadimplencia Efetiva",
  "812210104": "Convênios e Instrumentos Congeneres Aprovado",
  "812210103": "Convenios e Instrumentos Congeneres a Aprovar",
  "812210105": "Convenios e Instrumentos Congeneres Impugnados",
  "812210107": "Convênios e Instrumentos Congeneres em Inadimplencias Suspensa",
  "812210211": "Concluido",
  "812210203": "Comprovado",
  "712210101": "Valores Firmados",
  "811210110": "Convenios e Instrumentos Congeneres Arquivados"
};

const normalizeText = (text: string) => {
  if (!text) return "";
  return text.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
};

// Extract exactly 6 digits from string
const extractTransferId = (val: string | number | undefined): string | null => {
  if (!val) return null;
  const str = String(val).trim();
  const match = str.match(/(?:^|\D)(\d{6})(?:\D|$)/);
  return match ? match[1] : null;
};

export default function App() {
  const [fileTransferegov, setFileTransferegov] = useState<File | null>(null);
  const [fileSiafi, setFileSiafi] = useState<File | null>(null);
  const [results, setResults] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [stats, setStats] = useState({ total: 0, corretos: 0, inconsistencias: 0, naoEncontrados: 0 });

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
        readExcelMatrix(fileSiafi)
      ]);

      if (rowsTg.length === 0 || rowsSiafi.length === 0) {
        alert("Uma das planilhas está vazia!");
        setIsProcessing(false);
        return;
      }

      // 1. Process Transferegov
      let idColTgIdx = -1;
      let statusColTgIdx = -1;
      let convenenteColTgIdx = -1;

      for (let i = 0; i < Math.min(10, rowsTg.length); i++) {
        const r = rowsTg[i];
        if (!r) continue;
        r.forEach((cell, idx) => {
          const norm = normalizeText(String(cell));
          if (idColTgIdx === -1 && (norm.includes("transferencia") || norm.includes("convenio"))) {
            idColTgIdx = idx;
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

      const tgMap = new Map<string, { status: string, convenente: string }>();
      for (let i = 0; i < rowsTg.length; i++) {
        const row = rowsTg[i];
        if (!row || row.length === 0) continue;

        let id = "";
        let status = "";
        let convenente = "";
        
        // Find ID
        if (idColTgIdx !== -1 && extractTransferId(row[idColTgIdx])) {
           id = extractTransferId(row[idColTgIdx])!;
        } else {
           for (const cell of row) {
              const match = extractTransferId(cell);
              if (match) { id = match; break; }
           }
        }

        if (id) {
           // Find Status
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

           // Find Convenente
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
              tgMap.set(id, { status, convenente });
           }
        }
      }

      // 2. Process SIAFI (Cross-tab handling)
      let idColSiafiIdx = -1;
      let convenenteColSiafiIdx = -1;
      let accountCols: { idx: number, code: string, headerName: string }[] = [];
      let siafiHeaderRowIdx = 0;
      let maxAccountColsFound = 0;

      for (let i = 0; i < Math.min(20, rowsSiafi.length); i++) {
        const r = rowsSiafi[i];
        if (!r) continue;
        
        let currentAccountCols: { idx: number, code: string, headerName: string }[] = [];
        let currentIdColIdx = -1;
        let currentConvenenteColIdx = -1;

        r.forEach((cell, idx) => {
          const text = String(cell).trim();
          const norm = normalizeText(text);

          if (currentIdColIdx === -1 && (norm === "transferencia" || norm.includes("transferência") || norm.includes("convenio")) && !norm.includes("convenente")) {
            currentIdColIdx = idx;
          }
          if (currentConvenenteColIdx === -1 && (norm.includes("transferencia-convenente") || norm.includes("favorecido") || norm === "convenente")) {
            currentConvenenteColIdx = idx;
          }
          
          const match = text.match(/\b\d{9}\b/);
          if (match) {
            currentAccountCols.push({ idx, code: match[0], headerName: text });
          }
        });

        if (currentAccountCols.length > maxAccountColsFound) {
          maxAccountColsFound = currentAccountCols.length;
          accountCols = currentAccountCols;
          siafiHeaderRowIdx = i;
          if (currentIdColIdx !== -1) idColSiafiIdx = currentIdColIdx;
          if (currentConvenenteColIdx !== -1) convenenteColSiafiIdx = currentConvenenteColIdx;
        } else if (currentAccountCols.length === maxAccountColsFound) {
          if (currentIdColIdx !== -1) idColSiafiIdx = currentIdColIdx;
          if (currentConvenenteColIdx !== -1) convenenteColSiafiIdx = currentConvenenteColIdx;
        }
      }

      if (idColSiafiIdx === -1) {
        for (let i = siafiHeaderRowIdx + 1; i < Math.min(siafiHeaderRowIdx + 15, rowsSiafi.length); i++) {
          const r = rowsSiafi[i];
          const foundIdx = r.findIndex(cell => extractTransferId(cell));
          if (foundIdx !== -1) {
            idColSiafiIdx = foundIdx;
            break;
          }
        }
      }

      if (idColSiafiIdx === -1 || accountCols.length === 0) {
        alert("Não foi possível identificar o formato da planilha do SIAFI.");
        setIsProcessing(false);
        return;
      }

      let corretos = 0, inconsistencias = 0, naoEncontrados = 0;
      const finalResults: any[] = [];

      for (let i = siafiHeaderRowIdx + 1; i < rowsSiafi.length; i++) {
        const row = rowsSiafi[i];
        if (!row || row.length === 0) continue;

        const idSiafi = extractTransferId(row[idColSiafiIdx]);
        if (!idSiafi) continue;

        let contaSiafiRaw = "-";
        let contaCode = "";

        for (const ac of accountCols) {
          const val = row[ac.idx];
          if (val !== undefined && val !== null && val !== "") {
            const strVal = String(val).trim();
            if (strVal !== "-" && strVal !== "0" && strVal !== "0,00" && strVal !== "0.00") {
              contaCode = ac.code;
              contaSiafiRaw = ac.headerName;
              break;
            }
          }
        }

        // Try getting convenente directly from SIAFI as a fallback
        let convenenteSiafi = "";
        if (convenenteColSiafiIdx !== -1 && row[convenenteColSiafiIdx] !== undefined) {
           convenenteSiafi = String(row[convenenteColSiafiIdx]).trim();
        }

        let situacaoTg = "Sem Registro no Transferegov";
        let convenenteNome = convenenteSiafi || "Não Informado";
        let statusConciliacao = "Inconsistência";
        let situacaoSiafiDisplay = "Nenhuma Conta Detectada";

        if (contaCode) {
           const accountName = accountNames[contaCode] || contaSiafiRaw;
           situacaoSiafiDisplay = `${contaCode} - ${accountName}`;
        }

        if (tgMap.has(idSiafi)) {
          const tgData = tgMap.get(idSiafi)!;
          situacaoTg = tgData.status;
          // Prefer Transferegov Convenente, but use SIAFI if TG is empty, or fallback to unknown
          if (tgData.convenente && tgData.convenente !== "-") {
              convenenteNome = tgData.convenente;
          }
          
          if (contaCode) {
            const matchedRuleKey = Object.keys(validationRules).find(
              rule => normalizeText(rule) === normalizeText(situacaoTg)
            );

            if (matchedRuleKey) {
              const validAccounts = validationRules[matchedRuleKey];
              if (validAccounts.includes(contaCode)) {
                statusConciliacao = "Correto";
                corretos++;
              } else {
                statusConciliacao = "Inconsistência";
                inconsistencias++;
              }
            } else {
              statusConciliacao = "Inconsistência";
              inconsistencias++;
            }
          } else {
             statusConciliacao = "Inconsistência";
             inconsistencias++;
          }
        } else {
          naoEncontrados++;
          inconsistencias++;
        }

        // Clean Object generation for the Excel export
        // We place the most important audit columns FIRST so they appear at the beginning of the sheet
        const resultRow: any = {
           "Nº Transferência": idSiafi,
           "Convenente (Recebedor)": convenenteNome,
           "Situação (Transferegov)": situacaoTg,
           "Situação (SIAFI)": situacaoSiafiDisplay,
           "Status Conciliação": statusConciliacao,
        };

        // Then we append the original SIAFI data
        const headers = rowsSiafi[siafiHeaderRowIdx] || [];
        for (let colIdx = 0; colIdx < Math.max(headers.length, row.length); colIdx++) {
           let headerName = headers[colIdx] ? String(headers[colIdx]).trim() : "";
           if (!headerName) headerName = `Coluna_${colIdx + 1}`;
           
           // Ensure unique keys so we don't overwrite our clean columns or other merged columns
           if (resultRow.hasOwnProperty(headerName)) {
               headerName = `${headerName}_SIAFI`;
           }
           
           let finalKey = headerName;
           let counter = 1;
           while (resultRow.hasOwnProperty(finalKey)) {
               finalKey = `${headerName}_${counter}`;
               counter++;
           }
           
           resultRow[finalKey] = row[colIdx];
        }

        finalResults.push({
          idSiafi,
          convenenteNome,
          situacaoTg,
          situacaoSiafiDisplay,
          statusConciliacao,
          fullData: resultRow
        });
      }

      setStats({ 
        total: finalResults.length, 
        corretos, 
        inconsistencias, 
        naoEncontrados 
      });
      setResults(finalResults);
    } catch (err) {
      console.error(err);
      alert("Erro ao processar planilhas.");
    }

    setIsProcessing(false);
  };

  const exportResults = () => {
    const dataToExport = results.map(r => r.fullData);
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Conciliação Completa");
    XLSX.writeFile(wb, "Resultado_Conciliacao_SIACT.xlsx");
  };

  return (
    <div className="app-container">
      <div className="header">
        <h1>SIACT Hub</h1>
        <p>Sistema de Conciliação Transferegov x SIAFI (Tesouro Gerencial)</p>
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

      {results.length > 0 && (
        <div className="results-panel">
          <div className="results-header">
            <h2>Resultados da Auditoria ({stats.total} linhas)</h2>
            <button className="btn-primary" style={{ width: 'auto', marginBottom: 0 }} onClick={exportResults}>
              <Download size={18} /> Exportar Excel Completo
            </button>
          </div>

          <div className="stats-grid">
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
          </div>

          <div className="table-container">
            <table>
              <thead>
                <tr>
                  <th>Nº Transferência</th>
                  <th>Convenente</th>
                  <th>Situação (Transferegov)</th>
                  <th>Situação (SIAFI)</th>
                  <th>Status Conciliação</th>
                </tr>
              </thead>
              <tbody>
                {results.slice(0, 100).map((row, idx) => {
                  return (
                    <tr key={idx}>
                      <td>{row.idSiafi}</td>
                      <td title={row.convenenteNome}>
                         {row.convenenteNome.length > 30 ? row.convenenteNome.substring(0, 30) + "..." : row.convenenteNome}
                      </td>
                      <td>{row.situacaoTg}</td>
                      <td title={row.situacaoSiafiDisplay}>
                         {row.situacaoSiafiDisplay.length > 40 ? row.situacaoSiafiDisplay.substring(0, 40) + "..." : row.situacaoSiafiDisplay}
                      </td>
                      <td>
                        <span className={`status-badge ${
                          row.statusConciliacao === 'Correto' ? 'status-correct' : 
                          'status-incorrect'
                        }`}>
                          {row.statusConciliacao === 'Correto' && <CheckCircle size={14} style={{display:'inline', marginRight:'4px', verticalAlign:'text-bottom'}} />}
                          {row.statusConciliacao === 'Inconsistência' && <XCircle size={14} style={{display:'inline', marginRight:'4px', verticalAlign:'text-bottom'}} />}
                          {row.statusConciliacao}
                        </span>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
            {results.length > 100 && (
              <p style={{ textAlign: 'center', marginTop: '1rem', color: 'var(--text-muted)' }}>
                Mostrando as primeiras 100 linhas. Exporte para Excel para ver a planilha SIAFI completa com a conciliação.
              </p>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
