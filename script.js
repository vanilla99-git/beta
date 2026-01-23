const REGEX_ENDERECO = /^[A-Z]\d{2}[A-Z]$/;
const REGEX_CARRETA = /^1K[0-9A-Z]+$/;
const DESTINOS_IGNORADOS = new Set(["BSB", "DFV", "CBS"]); // Cargas da loja
const DESTINOS_LOJA = DESTINOS_IGNORADOS; // Alias para melhor legibilidade
const PREFIXOS_IGNORADOS = ["PAG", "PKC"];

const DESTINO_EQUIV_GROUPS = [
  { key: "SDU_QNR", label: "SDU/QNR", codes: new Set(["SDU", "QNR"]) },
  { key: "SP_REGIAO", label: "SP", codes: new Set(["CGH", "VJD", "SSZ", "BNU", "QOZ", "MGF", "SBC", "SPO", "SPA"]) },
];

const SHC_GLOSSARIO = {
  "CAO": { significado: "Cargo Aircraft Only", descricao: "Carga transportada apenas em aviões cargueiros." },
  "ICE": { significado: "Dry Ice", descricao: "Gelo seco (usado para refrigeração, requer ventilação)." },
  "MAG": { significado: "Magnetized Material", descricao: "Material magnetizado que pode interferir na bússola do avião." },
  "RCL": { significado: "Cryogenic Liquids", descricao: "Líquidos criogênicos." },
  "RFL": { significado: "Flammable Liquid", descricao: "Líquidos inflamáveis." },
  "RNG": { significado: "Non-Flammable Gas", descricao: "Gás não inflamável e não tóxico." },
  "RPB": { significado: "Toxic Substance", descricao: "Substância tóxica/venenosa." },
  "ELI": { significado: "Lithium Batteries", descricao: "Baterias de lítio (Iônico ou Metal), um dos códigos mais comuns hoje." },
  "ELM": { significado: "Lithium Batteries", descricao: "Baterias de lítio (Iônico ou Metal), um dos códigos mais comuns hoje." },
  "PER": { significado: "Perishable", descricao: "Carga perecível em geral." },
  "PEM": { significado: "Meat", descricao: "Carne fresca ou congelada." },
  "PES": { significado: "Seafood", descricao: "Frutos do mar e peixes." },
  "PEF": { significado: "Flowers", descricao: "Flores e plantas cortadas." },
  "COL": { significado: "Cool Storage", descricao: "Exige armazenamento em câmara fria (2°C a 8°C)." },
  "FRO": { significado: "Frozen Storage", descricao: "Exige armazenamento congelado (abaixo de 0°C)." },
  "AVI": { significado: "Live Animal in Hold", descricao: "Animal transportado no porão (comum na aviação comercial)." },
  "AVC": { significado: "Live Animal - Cold Blooded", descricao: "Animais de sangue frio (répteis, peixes)." },
  "AVW": { significado: "Live Animal - Warm Blooded", descricao: "Animais de sangue quente." },
  "VAL": { significado: "Valuable Cargo", descricao: "Carga de alto valor (ouro, pedras preciosas, notas)." },
  "VUN": { significado: "Vulnerable Cargo", descricao: "Carga vulnerável a roubos (eletrônicos, smartphones)." },
  "HUM": { significado: "Human Remains", descricao: "Restos mortais (esquife/caixão)." },
  "DIP": { significado: "Diplomatic Mail", descricao: "Mala diplomática (não pode ser aberta sem autorização)." },
  "HEA": { significado: "Heavy Cargo", descricao: "Carga pesada (geralmente acima de 150kg por volume)." },
  "WET": { significado: "Wet Cargo", descricao: "Carga úmida ou molhada." },
  "COM": { significado: "Combustible", descricao: "Material combustível." },
  "RDS": { significado: "Diagnostic Specimens", descricao: "Refere-se a amostras biológicas coletadas para diagnóstico ou investigação (ex: amostras de sangue, urina ou tecidos). Embora sejam materiais biológicos, são classificados como substâncias biológicas de Categoria B (UN3373) e possuem exigências de embalagem específicas, mas menos restritivas que o código INF (Infectious Substances)." },
  "LHO": { significado: "Living Human Organs", descricao: "Identifica órgãos humanos destinados a transplante. É uma das cargas de maior prioridade na aviação, exigindo manuseio imediato e coordenação direta entre a rampa e a tripulação para garantir que o tempo de viabilidade do órgão seja respeitado." },
};

const state = {
  fileName: null,
  delimiter: ";",
  hasHeader: true,
  // Defaults alinhados ao layout padrão do CSV informado pelo usuário:
  // Identificação, Origem, Destino, Peças, Peso, Localização, SHC, Cliente, Emissão, Serviço, ...
  colId: 0,
  colDest: 2,
  colLocal: 5,
  colPeso: 4,
  colPecas: 3,
  colData: 8,
  colServico: 9,
  colShc: 6,
  colPri: null,
  colSla: null,
  colPosse: null,
  colEmissao: null,
  rawRows: [],
  rows: [],
};

const stateVoos = {
  fileName: null,
  rawRows: [],
  voos: [],
  filtroData: "hoje", // "hoje" ou "amanha"
  filtroTerminal: "T1", // "T1", "T2", "T3"
};

function $(id) {
  return document.getElementById(id);
}

function setText(id, text) {
  const el = $(id);
  if (el) el.textContent = text ?? "";
}

function buildBadge(text, className) {
  const span = document.createElement("span");
  span.className = "badge rounded-pill" + (className ? " " + className : "");
  span.textContent = text;
  return span;
}

function buildShcElement(shcCode) {
  const span = document.createElement("span");
  span.className = "mono";
  span.textContent = shcCode;
  span.style.cursor = "help";
  span.style.borderBottom = "1px dotted #5a5a5a";
  
  const info = SHC_GLOSSARIO[shcCode];
  if (info) {
    span.setAttribute("data-bs-toggle", "tooltip");
    span.setAttribute("data-bs-html", "true");
    span.setAttribute("data-bs-placement", "top");
    span.setAttribute("title", `<strong>${escapeHtml(shcCode)}</strong><br/><em>${escapeHtml(info.significado)}</em><br/>${escapeHtml(info.descricao)}`);
  }
  
  return span;
}

function buildShcElements(shcLabel) {
  if (!shcLabel || shcLabel === "-" || shcLabel.trim() === "") return null;
  
  const container = document.createElement("span");
  container.className = "mono";
  
  const codes = shcLabel.split("/").map(s => s.trim()).filter(s => s && s !== "-");
  const validCodes = codes.filter(code => SHC_GLOSSARIO[code]);
  
  if (validCodes.length === 0) {
    container.textContent = shcLabel;
    return container;
  }
  
  if (validCodes.length === 1) {
    return buildShcElement(validCodes[0]);
  }
  
  // Múltiplos SHCs: criar um tooltip com lista
  container.textContent = shcLabel;
  container.style.cursor = "help";
  container.style.borderBottom = "1px dotted #5a5a5a";
  container.setAttribute("data-bs-toggle", "tooltip");
  container.setAttribute("data-bs-html", "true");
  container.setAttribute("data-bs-placement", "top");
  
  const tooltipContent = validCodes.map(code => {
    const info = SHC_GLOSSARIO[code];
    return `<div style="margin-bottom: 6px; padding-bottom: 6px; border-bottom: 1px solid #ddd;">
      <strong>${escapeHtml(code)}</strong> — ${escapeHtml(info.significado)}<br/>
      <small style="color: #666;">${escapeHtml(info.descricao)}</small>
    </div>`;
  }).join("");
  
  container.setAttribute("title", `<div style="max-width: 300px;">${tooltipContent}</div>`);
  
  return container;
}

function serviceLabelForRow(row) {
  if (!row) return "";
  const key = (row.servicoGroup || servicoGroupKey(row.servico || "") || "").toString();
  const label = servicoLabelFromKey(key);
  if (label) return label;
  return (row.servico || "").toString().trim();
}

function escapeHtml(s) {
  return (s ?? "")
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function stripDiacritics(s) {
  try {
    return (s ?? "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  } catch {
    return (s ?? "").toString();
  }
}

function normalizeUpper(s) {
  return stripDiacritics(s).toUpperCase().trim();
}

function destinoEquivGroup(destino) {
  const d = normalizeUpper(destino);
  if (!d) return null;
  for (const g of DESTINO_EQUIV_GROUPS) {
    if (g.codes.has(d)) return g;
  }
  return null;
}

function destinoEquivKey(destino) {
  const g = destinoEquivGroup(destino);
  return g ? g.key : normalizeUpper(destino);
}

function destinoEquivLabelFromKey(key) {
  const k = (key || "").toString();
  const g = DESTINO_EQUIV_GROUPS.find((x) => x.key === k);
  return g ? g.label : k;
}

function destinoLabelFromRawSet(destinosRawSet) {
  const vals = Array.from(destinosRawSet || [])
    .map((x) => normalizeUpper(x))
    .filter(Boolean);
  if (!vals.length) return "";

  const keys = new Set(vals.map((d) => destinoEquivKey(d)));
  if (keys.size > 1) return "MISTO";

  if (vals.length === 1) return vals[0];
  return destinoEquivLabelFromKey(vals.map((d) => destinoEquivKey(d))[0]);
}

function normalizeHeaderCell(h) {
  return normalizeUpper(h).replace(/[^A-Z0-9]+/g, " ").trim();
}

function normalizeId(raw) {
  return normalizeUpper(raw).replace(/\s+/g, "");
}

function normalizeEndereco(raw) {
  return normalizeUpper(raw).replace(/\s+/g, "");
}

function normalizarLocal(raw) {
  const s = fixUtf8Mojibake((raw || "").toString()).trim();
  if (!s) return "";

  const mAddr = s.match(/([A-Za-z]\d{2}[A-Za-z])/);
  if (mAddr) return mAddr[1].toUpperCase().replace(/\s+/g, "");

  const mCarreta = s.match(/\b(1k[0-9a-z]+)\b/i);
  if (mCarreta) return mCarreta[1].toUpperCase().replace(/\s+/g, "");

  const head = s.split(":")[0].trim();
  return head.toUpperCase().replace(/\s+/g, "");
}

function parseMultiplosEnderecos(raw) {
  const s = fixUtf8Mojibake((raw || "").toString()).trim();
  if (!s) return new Map();

  const result = new Map();
  const pattern = /([A-Z]\d{2}[A-Z])\s*:\s*(\d+)\s*Vol/gi;
  let match;

  while ((match = pattern.exec(s)) !== null) {
    const addr = match[1].toUpperCase();
    const vols = parseInt(match[2], 10);
    if (Number.isFinite(vols) && vols > 0) {
      result.set(addr, vols);
    }
  }

  return result;
}

function formatLocalizacao(raw) {
  const parsed = parseMultiplosEnderecos(raw);
  if (parsed.size === 0) return raw; // Fallback ao original se não conseguir parsear
  
  const entries = Array.from(parsed.entries()).sort((a, b) => a[0].localeCompare(b[0]));
  return entries.map(([addr, vols]) => `${addr} x${vols}`).join(", ");
}

function extrairData(raw) {
  return (raw || "").toString().trim().split(/\s+/)[0] || "";
}

function parsePtNumber(raw) {
  const s = (raw || "")
    .toString()
    .replace(/\s+/g, "")
    .replace(/kg/gi, "")
    .replace(/\./g, "")
    .replace(",", ".")
    .replace(/[^\d.-]/g, "");
  const v = parseFloat(s);
  return Number.isFinite(v) ? v : 0;
}

function parseIntSafe(raw) {
  const s = (raw ?? "").toString().replace(/[^\d-]/g, "");
  const v = parseInt(s, 10);
  return Number.isFinite(v) ? v : 0;
}

function parsePtDateToMs(raw) {
  const s = (raw ?? "").toString().trim();
  const m = s.match(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/);
  if (!m) return null;
  const dd = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  let yyyy = parseInt(m[3], 10);
  if (!Number.isFinite(dd) || !Number.isFinite(mm) || !Number.isFinite(yyyy)) return null;
  if (yyyy < 100) yyyy += 2000;
  const d = new Date(yyyy, mm - 1, dd);
  const t = d.getTime();
  return Number.isFinite(t) ? t : null;
}

function priFromServico(servicoNorm) {
  const s = normalizeUpper(servicoNorm);
  if (!s) return "P3";

  const p1 = ["SAUDE", "URGENTE", "CHEGOL"];
  if (p1.some((k) => s.includes(k))) return "P1";

  const p2 = ["GCE", "ECG", "RAPIDO", "RAPIDO/KG"];
  if (p2.some((k) => s.includes(k))) return "P2";

  const p5 = ["SER", "SERV", "SERVICO"];
  if (p5.some((k) => s === k || s.startsWith(k))) return "P5";

  const p3 = ["ECONOMICO", "MODA"];
  if (p3.some((k) => s.includes(k))) return "P3";

  return "P3";
}

function normalizePriCell(raw, servicoNorm) {
  const s = normalizeUpper(raw);
  const m = s.match(/\bP?\s*([1235])\b/);
  if (m) return `P${m[1]}`;
  if (s === "P1" || s === "P2" || s === "P3" || s === "P5") return s;
  return priFromServico(servicoNorm);
}

function parseSlaCell(raw) {
  const s = normalizeUpper(raw).replace(/\s+/g, " ").trim();
  if (!s) return { raw: "", vencida: false, pctRemaining: null };
  const vencida = s.includes("VENC") || s.includes("ATRAS") || s.includes("EXPIR");

  let pct = null;
  const mPct = s.match(/(\d{1,3})(?:[.,](\d+))?\s*%/);
  if (mPct) {
    const v = parseFloat(`${mPct[1]}.${mPct[2] || ""}`);
    if (Number.isFinite(v)) pct = Math.max(0, Math.min(100, v));
  } else {
    const mNum = s.match(/^\s*(\d{1,3})(?:[.,](\d+))?\s*$/);
    if (mNum) {
      const v = parseFloat(`${mNum[1]}.${mNum[2] || ""}`);
      if (Number.isFinite(v)) pct = Math.max(0, Math.min(100, v));
    }
  }

  return { raw: s, vencida, pctRemaining: pct };
}

function parsePosseHours(raw) {
  const s = normalizeUpper(raw).replace(/\s+/g, " ").trim();
  if (!s) return null;

  const mDias = s.match(/([\d.,]+)\s*(D|DIA|DIAS)\b/);
  if (mDias) {
    const v = parseFloat(mDias[1].replace(",", "."));
    return Number.isFinite(v) ? v * 24 : null;
  }

  const mHoras = s.match(/([\d.,]+)\s*(H|HORA|HORAS)\b/);
  if (mHoras) {
    const v = parseFloat(mHoras[1].replace(",", "."));
    return Number.isFinite(v) ? v : null;
  }

  const mNum = s.match(/(\d+(?:[.,]\d+)?)/);
  if (mNum) {
    const v = parseFloat(mNum[1].replace(",", "."));
    return Number.isFinite(v) ? v : null;
  }

  return null;
}

function formatNumber(v) {
  if (!Number.isFinite(v)) return "0";
  const isInt = Math.abs(v - Math.round(v)) < 1e-9;
  return isInt ? String(Math.round(v)) : v.toFixed(1);
}

function detectDelimiter(line) {
  const s = line || "";
  const candidates = [";", ",", "\t"];
  let best = ";";
  let bestCount = -1;
  for (const d of candidates) {
    const count = s.split(d).length - 1;
    if (count > bestCount) {
      bestCount = count;
      best = d;
    }
  }
  return best;
}

function parseCsv(text, delimiter) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const c = text[i];

    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
      continue;
    }

    if (c === '"') {
      inQuotes = true;
      continue;
    }
    if (c === delimiter) {
      row.push(field);
      field = "";
      continue;
    }
    if (c === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      continue;
    }
    if (c === "\r") continue;
    field += c;
  }

  row.push(field);
  rows.push(row);
  while (rows.length && rows[rows.length - 1].every((v) => (v || "").trim() === "")) rows.pop();
  return rows;
}

function readFileAsBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Falha ao ler o arquivo."));
    reader.onload = () => resolve(reader.result);
    reader.readAsArrayBuffer(file);
  });
}

function decodeBuffer(buf, encoding) {
  return new TextDecoder(encoding, { fatal: false }).decode(buf);
}

function mojibakeScore(text) {
  const s = (text || "").slice(0, 4096);
  const m = (re) => (s.match(re) || []).length;
  return m(/Ã./g) + m(/Â./g) + m(/â./g) + m(/\uFFFD/g);
}

function fixUtf8Mojibake(s) {
  const text = (s ?? "").toString();
  if (!text) return text;
  if (!/[ÃÂâ\uFFFD]/.test(text)) return text;
  try {
    const bytes = new Uint8Array(Array.from(text, (ch) => ch.charCodeAt(0) & 0xff));
    const decoded = new TextDecoder("utf-8", { fatal: false }).decode(bytes);
    if (mojibakeScore(decoded) < mojibakeScore(text)) return decoded;
  } catch {}
  return text;
}

function guessHasHeader(rows) {
  if (!rows || rows.length < 2) return false;
  const first = rows[0] || [];
  const joined = first.map((c) => normalizeHeaderCell(c)).join(" ");
  const keywords = ["IDENTIFICACAO", "DESTINO", "LOCALIZACAO", "LOCAL", "PESO", "KG", "PECAS", "DATA", "EMISSAO", "SERVICO", "SHC"];
  if (keywords.some((k) => joined.includes(k))) return true;

  let alphaCells = 0;
  for (const c of first) {
    const s = (c ?? "").toString();
    if (/[A-Za-zÀ-ÿ]/.test(s)) alphaCells++;
  }
  return alphaCells >= Math.max(2, Math.floor(first.length / 3));
}

function detectColumnsFromHeader() {
  if (!state.hasHeader) return;
  const header = state.rawRows && state.rawRows.length ? state.rawRows[0] || [] : [];
  if (!header.length) return;
  const norm = header.map((h) => normalizeHeaderCell(fixUtf8Mojibake(h)));

  const find = (tokens) => {
    for (let i = 0; i < norm.length; i++) {
      const cell = norm[i];
      if (!cell) continue;
      for (const t of tokens) {
        if (cell === t || cell.includes(t)) return i;
      }
    }
    return null;
  };

  state.colId = find(["IDENTIFICACAO", "AWB", "ID"]) ?? state.colId;
  state.colDest = find(["DESTINO"]) ?? state.colDest;
  state.colLocal = find(["LOCALIZACAO", "LOCAL"]) ?? state.colLocal;
  state.colPeso = find(["PESO", "KG"]) ?? state.colPeso;
  state.colPecas = find(["PECAS", "PECA", "VOLS", "VOLUMES", "VOLUME"]);
  state.colData = find(["DATA", "EMISSAO"]) ?? state.colData;
  state.colServico = find(["SERVICO"]) ?? state.colServico;
  state.colShc = find(["SHC"]);
  state.colPri = find(["PRI", "PRIORIDADE", "PRIOR"]) ?? state.colPri;
  state.colSla = find(["SLA", "STATUS SLA", "SAUDE SLA", "SAUDE"]) ?? state.colSla;
  state.colPosse = find(["POSSE", "TEMPO POSSE", "TEMPO DE POSSE", "TEMPO NO ESTOQUE", "ESTOQUE", "AGING", "IDADE"]) ?? state.colPosse;
  state.colEmissao = find(["EMISSAO", "DATA EMISSAO", "DT EMISSAO"]) ?? state.colEmissao;
}

function normalizeServicoCell(raw) {
  const s = normalizeUpper(raw).replace(/\s+/g, " ");
  if (!s) return "";
  return s.replace(/\s*\/\s*/g, "/");
}

function servicoGroupKey(servicoNorm) {
  const s = normalizeServicoCell(servicoNorm);
  if (!s) return "";

  if (s === "ECONOMICO" || s === "ECONÔMICO" || s === "MODA") return "ECO_MODA";

  const expSet = new Set(["GCE", "ECG", "SERV", "SER", "SERVICO", "SERVIÇO", "RAPIDO", "RÁPIDO", "RAPIDO/KG", "RÁPIDO/KG"]);
  if (expSet.has(s)) return "EXPRESSO";
  if (s.startsWith("RAPIDO") || s.startsWith("RÁPIDO")) return "EXPRESSO";
  if (s.startsWith("SERV")) return "EXPRESSO";

  return s;
}

function servicoLabelFromKey(servKey) {
  const k = (servKey || "").toString();
  if (!k) return "";
  if (k === "ECO_MODA") return "ECO/MODA";
  if (k === "EXPRESSO") return "EXPRESSO";
  return k;
}

function parseShcSet(raw) {
  const s = normalizeUpper(raw);
  if (!s) return new Set();
  const parts = s
    .split(/[\s,|+/;]+/g)
    .map((x) => x.trim())
    .filter(Boolean);
  return new Set(parts);
}

function shcLabelFromSet(set) {
  return set && set.size ? Array.from(set).sort().join(", ") : "-";
}

function analyzeShcConflict(tokensSet) {
  const tokens = new Set(Array.from(tokensSet || []).map((x) => normalizeUpper(x)));
  if (tokens.size === 0) return false;

  const GENERAL_OK = new Set(["COM", "HEA", "PER"]);
  const LITHIUM = new Set(["ELI", "ELM"]);

  const hasLithium = Array.from(LITHIUM).some((t) => tokens.has(t));
  if (hasLithium) {
    for (const t of tokens) {
      if (LITHIUM.has(t)) continue;
      if (GENERAL_OK.has(t)) continue;
      return true;
    }
  }

  const hasWet = tokens.has("WET");
  if (hasWet) {
    const WET_OK = new Set(["WET", "COM", "HEA"]);
    for (const t of tokens) {
      if (!WET_OK.has(t)) return true;
    }
  }

  return false;
}

function isEnderecado(row) {
  return REGEX_ENDERECO.test(normalizeEndereco(row.localOriginal || ""));
}

// Verifica se um endereço normalizado é um bolsão válido (não carreta)
function isValidBolsaoAddress(normalized) {
  if (!REGEX_ENDERECO.test(normalized)) return false;
  if (!isPracaAddress(normalized)) return false;
  return true;
}

// Verifica se uma localização bruta é carreta
function isCarretaLocation(raw) {
  return /^1k/i.test((raw || "").trim());
}

function isIgnoredRow({ id, destino, localOriginal }, options = {}) {
  const { allowIgnoredDestinos = false } = options;
  const destUp = normalizeUpper(destino);
  if (!allowIgnoredDestinos && DESTINOS_IGNORADOS.has(destUp)) return true;

  const idUp = normalizeId(id);
  for (const p of PREFIXOS_IGNORADOS) {
    if (idUp.startsWith(p)) return true;
  }

  const loc = normalizarLocal(localOriginal);
  const locOk = loc === "" || REGEX_ENDERECO.test(loc) || REGEX_CARRETA.test(loc);
  return !locOk;
}

function mapRowsFromRaw() {
  const rows = [];
  const start = state.hasHeader ? 1 : 0;

  for (let i = start; i < state.rawRows.length; i++) {
    const row = parseRowFromRaw(i);
    if (!row) continue;
    rows.push(row);
  }

  state.rows = rows;
}

function parseRowFromRaw(index, options = {}) {
  const { allowIgnoredDestinos = false } = options;
  const cols = state.rawRows[index] || [];
  const id = (cols[state.colId] || "").toString().trim();
  if (!id) return null;

  const destino = normalizeUpper(cols[state.colDest] || "");
  const localOriginalRaw = (cols[state.colLocal] ?? "").toString();
  const localOriginal = normalizarLocal(localOriginalRaw);
  const peso = parsePtNumber(cols[state.colPeso] || "");
  const pecasRaw = state.colPecas == null ? "" : cols[state.colPecas];
  const pecas = Math.max(1, parseIntSafe(pecasRaw));
  const data = extrairData(cols[state.colData] || "");
  const emissaoCol = state.colEmissao == null ? state.colData : state.colEmissao;
  const emissao = extrairData(cols[emissaoCol] || "");
  const emissaoMs = parsePtDateToMs(emissao);
  const servico = normalizeServicoCell(cols[state.colServico] || "");
  const servicoGroup = servicoGroupKey(servico);
  const shcRaw = state.colShc == null ? "" : (cols[state.colShc] || "").toString();
  const shcTokens = parseShcSet(shcRaw);
  const priRaw = state.colPri == null ? "" : (cols[state.colPri] ?? "").toString();
  const pri = normalizePriCell(priRaw, servico);
  const slaRaw = state.colSla == null ? "" : (cols[state.colSla] ?? "").toString();
  const sla = parseSlaCell(slaRaw);
  const posseRaw = state.colPosse == null ? "" : (cols[state.colPosse] ?? "").toString();
  let posseHours = state.colPosse == null ? null : parsePosseHours(posseRaw);
  if (posseHours == null && emissaoMs != null) posseHours = Math.max(0, (Date.now() - emissaoMs) / (60 * 60 * 1000));

  if (isIgnoredRow({ id, destino, localOriginal }, { allowIgnoredDestinos })) return null;

  return {
    key: `${index}`,
    rawIndex: index,
    id,
    destino,
    localOriginal,
    localOriginalRaw,
    peso,
    pecas,
    data,
    emissao,
    emissaoMs,
    servico,
    servicoGroup,
    shcTokens,
    pri,
    priRaw,
    slaRaw: sla.raw,
    slaVencida: sla.vencida,
    slaPctRemaining: sla.pctRemaining,
    posseHours,
  };
}

function collectPrintableRows() {
  const rows = [];
  const start = state.hasHeader ? 1 : 0;
  for (let i = start; i < state.rawRows.length; i++) {
    const row = parseRowFromRaw(i, { allowIgnoredDestinos: true });
    if (!row) continue;
    rows.push(row);
  }
  return rows;
}

function buildRowDefinitions(bolsao, cols, maxNum) {
  if (bolsao === "A") return buildBolsaoARowDefs(cols, maxNum);
  if (bolsao === "B") return buildBolsaoBRowDefs(cols, maxNum);
  if (bolsao === "D") return buildBolsaoDRowDefs(cols, maxNum);
  return buildDefaultRowDefs(cols, maxNum);
}

function buildDefaultRowDefs(cols, maxNum) {
  const rows = [];
  for (let n = maxNum; n >= 1; n--) {
    rows.push({ num: n, allowedLetters: new Set(cols), spacerAfter: false });
  }
  return rows;
}

function buildBolsaoARowDefs(cols, maxNum) {
  const limitedRows = new Set([5, 11, 17, 18]);
  const spacerAfter = new Set([19, 17, 15, 13, 11, 9, 7, 5, 3]);
  const lettersWithoutH = cols.filter((col) => col !== "H");
  const rows = [];
  for (let n = maxNum; n >= 1; n--) {
    const allowedLetters = new Set(limitedRows.has(n) ? lettersWithoutH : cols);
    rows.push({ num: n, allowedLetters, spacerAfter: spacerAfter.has(n) });
  }
  return rows;
}

function buildBolsaoBRowDefs(cols, maxNum) {
  const limitedRows = new Set([7, 8, 9]);
  const spacerAfter = new Set([9, 5, 3]);
  const lettersWithoutGH = cols.filter((col) => col !== "G" && col !== "H");
  const rows = [];
  for (let n = maxNum; n >= 1; n--) {
    const allowedLetters = new Set(limitedRows.has(n) ? lettersWithoutGH : cols);
    rows.push({ num: n, allowedLetters, spacerAfter: spacerAfter.has(n) });
  }
  return rows;
}

function buildBolsaoDRowDefs(cols, maxNum) {
  const limitedRows = new Set([7, 8]);
  const spacerAfter = new Set([11, 9, 7, 5, 3]);
  const lettersWithoutJ = cols.filter((col) => col !== "J");
  const rows = [];
  for (let n = maxNum; n >= 1; n--) {
    const allowedLetters = new Set(limitedRows.has(n) ? lettersWithoutJ : cols);
    rows.push({ num: n, allowedLetters, spacerAfter: spacerAfter.has(n) });
  }
  return rows;
}

function setError(msg) {
  setText("errorBox", msg || "");
}

function setActionMsg(msg) {
  setText("actionMsg", msg || "");
}

function updateAppStamp() {
  const fileInfo = state.fileName ? `${state.fileName}` : "Sem arquivo";
  const delimLabel = state.delimiter === "\t" ? "TAB" : state.delimiter;
  setText("appStamp", `${fileInfo} | ${state.rows.length} linhas | ${delimLabel}`);
}

function updateKpis() {
  const rows = filteredRows();
  const total = rows.length;
  const ok = rows.filter(isEnderecado).length;
  const pend = total - ok;
  const peso = rows.reduce((acc, r) => acc + (Number.isFinite(r.peso) ? r.peso : 0), 0);

  setText("kpiTotal", String(total));
  setText("kpiOk", String(ok));
  setText("kpiPend", String(pend));
  setText("kpiPeso", formatNumber(peso));
}

function updateServicoFilter() {
  const el = $("filterServico");
  if (!el) return;
  const current = normalizeUpper(el.value || "TODOS");
  const services = Array.from(
    new Set(
      state.rows
        .map((r) => normalizeUpper(serviceLabelForRow(r)))
        .filter((v) => v && v !== "-")
    )
  ).sort();

  el.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = "TODOS";
  optAll.textContent = "Todos";
  el.appendChild(optAll);

  for (const service of services) {
    const opt = document.createElement("option");
    opt.value = service;
    opt.textContent = service;
    el.appendChild(opt);
  }

  el.value = services.includes(current) ? current : "TODOS";
}

function filteredRows() {
  const servFilter = normalizeUpper($("filterServico")?.value || "");
  const hasServFilter = servFilter && servFilter !== "TODOS";

  return state.rows.filter((r) => {
    if (isEnderecado(r)) return false;
    // Excluir cargas em carreta
    if (isCarretaLocation(r.localOriginal || "")) return false;
    // Excluir cargas da loja (CBS, DFV, BSB)
    const destNorm = normalizeUpper(r.destino || "");
    if (DESTINOS_LOJA.has(destNorm)) return false;
    if (hasServFilter) {
      const label = normalizeUpper(serviceLabelForRow(r));
      if (label !== servFilter) return false;
    }
    return true;
  });
}

function analyzeEnderecoGroup(rows) {
  const destinosRaw = new Set();
  const destinosKey = new Set();
  const datas = new Set();
  const servicoGroups = new Set();
  const shcKeys = new Set();
  const shcTokens = new Set();
  let kg = 0;

  for (const r of rows) {
    if (r.destino) {
      const d = normalizeUpper(r.destino);
      if (d) {
        destinosRaw.add(d);
        destinosKey.add(destinoEquivKey(d));
      }
    }
    datas.add((r.data || "").toString());

    const isAllocated = REGEX_ENDERECO.test(normalizeEndereco(r.localOriginal || ""));
    const serv = (r.servico || "").toString().trim();
    if (!(isAllocated && serv === "")) {
      const grp = (r.servicoGroup || servicoGroupKey(serv)).toString();
      if (grp) servicoGroups.add(grp);
    }

    const shcKey = shcLabelFromSet(r.shcTokens).replaceAll(", ", "+");
    if (shcKey && shcKey !== "-") shcKeys.add(shcKey);
    for (const t of r.shcTokens || []) shcTokens.add(t);

    kg += Number.isFinite(r.peso) ? r.peso : 0;
  }

  const dataMixed = datas.size > 1;
  
  // Verificar se serviço é realmente diferente
  // URGENTE e URGENTE/KG são considerados o mesmo
  const normalizeServiceGroup = (sg) => {
    const s = (sg || "").toString().toUpperCase();
    if (s.includes("URGENTE")) return "URGENTE";
    return s;
  };
  const normalizedServGroups = new Set(Array.from(servicoGroups).map(normalizeServiceGroup));
  const servMixed = normalizedServGroups.size > 1;
  
  const destMixed = destinosKey.size > 1;
  const shcMixed = shcKeys.size > 1;
  const shcConflict = analyzeShcConflict(shcTokens);
  const anyMixed = dataMixed || servMixed || destMixed || shcMixed || shcConflict;

  const singleOrMixed = (set) => {
    const vals = Array.from(set).filter((x) => (x ?? "").toString().trim() !== "");
    if (vals.length === 0) return "";
    if (vals.length === 1) return vals[0];
    return "MISTO";
  };

  return {
    destinos: destinosRaw,
    destKey: destinosKey.size === 1 ? Array.from(destinosKey)[0] : "",
    datas,
    servicoGroups,
    shcTokens,
    kg,
    destLabel: destinoLabelFromRawSet(destinosRaw),
    dataLabel: singleOrMixed(datas),
    servLabel: singleOrMixed(servicoGroups),
    dataMixed,
    servMixed,
    destMixed,
    shcMixed,
    shcConflict,
    anyMixed,
  };
}

function collectEnderecoItems() {
  const aggs = new Map();
  for (const r of state.rows) {
    const endereco = normalizeEndereco(r.localOriginal || "");
    if (!REGEX_ENDERECO.test(endereco)) continue;
    const list = aggs.get(endereco) || [];
    list.push(r);
    aggs.set(endereco, list);
  }

  return Array.from(aggs.entries()).map(([endereco, rows]) => {
    const analysis = analyzeEnderecoGroup(rows);
    return { endereco, rows, ...analysis };
  });
}

function updateProgressBar(id, value, total, labelId) {
  const el = $(id);
  if (!el) return;
  const safeTotal = Math.max(1, total || 1);
  const pct = Math.min(100, Math.round((value / safeTotal) * 100));
  el.style.width = `${pct}%`;
  el.setAttribute("aria-valuenow", pct);
  if (labelId) setText(labelId, `${pct}%`);
}

function getEnderecoStats() {
  // Contar volumes na triagem usando filteredRows
  const cargasTriagem = filteredRows();
  const volumesTriagem = cargasTriagem.reduce((sum, r) => {
    const pecas = Math.max(1, Number.isFinite(r.pecas) ? r.pecas : 1);
    return sum + pecas;
  }, 0);
  const pesoTriagem = cargasTriagem.reduce((sum, r) => sum + (Number.isFinite(r.peso) ? r.peso : 0), 0);
  
  // Contar volumes endereçados (no mesmo escopo: em bolsões, não carreta, não loja)
  const cargasEnderecoadas = [];
  let volumesEnderecoados = 0;
  let pesoEnderecoado = 0;
  for (const r of state.rows) {
    // Mesmos filtros de filteredRows
    const destNorm = normalizeUpper(r.destino || "");
    if (DESTINOS_LOJA.has(destNorm)) continue;
    if (isCarretaLocation(r.localOriginal || "")) continue;
    if (isCarretaLocation(r.localOriginalRaw || "")) continue; // Também verificar raw
    
    // Mas que ESTEJAM endereçados
    if (!isEnderecado(r)) continue;
    
    // E estejam em bolsão válido
    const addr = normalizeEndereco(r.localOriginal || "");
    if (!isValidBolsaoAddress(addr)) continue;
    
    const pecas = Math.max(1, Number.isFinite(r.pecas) ? r.pecas : 1);
    volumesEnderecoados += pecas;
    pesoEnderecoado += Number.isFinite(r.peso) ? r.peso : 0;
    cargasEnderecoadas.push({
      id: r.id,
      pecas,
      destino: r.destino || "-",
      endereco: addr || "-",
      localizacao: r.localOriginalRaw || "-"
    });
  }
  
  return {
    triagem: volumesTriagem,
    cargasTriagem: cargasTriagem,
    enderecoados: volumesEnderecoados,
    cargasEnderecoadas: cargasEnderecoadas,
    pesoTriagem,
    pesoEnderecoado,
    total: volumesEnderecoados + volumesTriagem
  };
}

// Helper: calcula contraste e retorna cor (preto ou branco) baseado no fundo
function getContrastColor(hexBg) {
  // Remove # se existir
  const hex = hexBg.replace("#", "");
  // Converte para RGB
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  // Calcula luminância relativa (WCAG formula)
  const luminancia = (r * 0.299 + g * 0.587 + b * 0.114) / 255;
  // Se luminância > 0.5, fundo é claro, usa preto; caso contrário usa branco
  return luminancia > 0.5 ? "#111" : "#fff";
}

function renderEnderecoStats() {
  const stats = getEnderecoStats();
  
  // 1. Cargas endereçadas (barra horizontal empilhada 100%)
  const cargasContainer = $("statsCargasChart");
  if (cargasContainer) {
    cargasContainer.innerHTML = "";
    const cargasTotal = stats.enderecoados + stats.triagem;
    const cargasEndPercent = cargasTotal > 0 ? Math.round((stats.enderecoados / cargasTotal) * 100) : 0;
    const cargasTriagemPercent = 100 - cargasEndPercent;
    
    // Criar barra horizontal empilhada
    const barWrapper = document.createElement("div");
    barWrapper.style.display = "flex";
    barWrapper.style.flexDirection = "column";
    barWrapper.style.gap = "8px";
    
    const barContainer = document.createElement("div");
    barContainer.style.display = "flex";
    barContainer.style.height = "40px";
    barContainer.style.borderRadius = "4px";
    barContainer.style.overflow = "visible";
    barContainer.style.border = "1px solid #ccc";
    barContainer.style.background = "white";
    barContainer.style.position = "relative";
    
    // Segmento Endereçado
    const barEndereco = document.createElement("div");
    barEndereco.style.flex = `${cargasEndPercent}%`;
    barEndereco.style.background = "#0d6efd";
    barEndereco.style.display = "flex";
    barEndereco.style.alignItems = "center";
    barEndereco.style.justifyContent = "center";
    barEndereco.style.color = "#fff";
    barEndereco.style.fontWeight = "700";
    barEndereco.style.fontSize = "0.8rem";
    barEndereco.style.overflow = "hidden";
    barContainer.appendChild(barEndereco);
    
    // Segmento Triagem
    const barTriagem = document.createElement("div");
    barTriagem.style.flex = `${cargasTriagemPercent}%`;
    barTriagem.style.background = "#d9d9d9";
    barTriagem.style.display = "flex";
    barTriagem.style.alignItems = "center";
    barTriagem.style.justifyContent = "center";
    barTriagem.style.color = "#111";
    barTriagem.style.fontWeight = "700";
    barTriagem.style.fontSize = "0.8rem";
    barTriagem.style.overflow = "hidden";
    barContainer.appendChild(barTriagem);
    
    // Overlay com percentual centralizado - cor com contraste inteligente
    const percentOverlay = document.createElement("div");
    percentOverlay.style.position = "absolute";
    percentOverlay.style.top = "50%";
    percentOverlay.style.left = "50%";
    percentOverlay.style.transform = "translate(-50%, -50%)";
    // Detectar qual segmento vai ser sobreposto para calcular contraste
    const bgColor1 = cargasEndPercent >= 50 ? "#0d6efd" : "#d9d9d9";
    percentOverlay.style.color = getContrastColor(bgColor1);
    percentOverlay.style.fontWeight = "700";
    percentOverlay.style.fontSize = "0.9rem";
    percentOverlay.style.textShadow = "0 0 3px rgba(255,255,255,0.8)";
    percentOverlay.textContent = `${cargasEndPercent}%`;
    barContainer.appendChild(percentOverlay);
    
    barWrapper.appendChild(barContainer);
    
    // Legenda
    const legend = document.createElement("div");
    legend.style.marginTop = "4px";
    legend.style.fontSize = "0.7rem";
    legend.style.display = "grid";
    legend.style.gap = "2px";
    legend.innerHTML = `
      <div style="display: flex; gap: 4px; align-items: center;">
        <div style="width: 8px; height: 8px; background: #0d6efd;"></div>
        <span>Endereçadas: ${stats.enderecoados} (${cargasEndPercent}%)</span>
      </div>
      <div style="display: flex; gap: 4px; align-items: center;">
        <div style="width: 8px; height: 8px; background: #d9d9d9;"></div>
        <span>Pendentes: ${stats.triagem} (${cargasTriagemPercent}%)</span>
      </div>
    `;
    barWrapper.appendChild(legend);
    
    cargasContainer.appendChild(barWrapper);
  }

  // 2. Peso total endereçado (barra horizontal empilhada 100%)
  const pesoContainer = $("statsPesoChart");
  if (pesoContainer) {
    pesoContainer.innerHTML = "";
    const pesoTotal = stats.pesoEnderecoado + stats.pesoTriagem;
    const pesoEndPercent = pesoTotal > 0 ? Math.round((stats.pesoEnderecoado / pesoTotal) * 100) : 0;
    const pesoTriagemPercent = 100 - pesoEndPercent;
    
    // Criar barra horizontal empilhada
    const barWrapper = document.createElement("div");
    barWrapper.style.display = "flex";
    barWrapper.style.flexDirection = "column";
    barWrapper.style.gap = "8px";
    
    const barContainer = document.createElement("div");
    barContainer.style.display = "flex";
    barContainer.style.height = "40px";
    barContainer.style.borderRadius = "4px";
    barContainer.style.overflow = "visible";
    barContainer.style.border = "1px solid #ccc";
    barContainer.style.background = "white";
    barContainer.style.position = "relative";
    
    // Segmento Endereçado
    const barEndereco = document.createElement("div");
    barEndereco.style.flex = `${pesoEndPercent}%`;
    barEndereco.style.background = "#fd7e14";
    barEndereco.style.display = "flex";
    barEndereco.style.alignItems = "center";
    barEndereco.style.justifyContent = "center";
    barEndereco.style.color = "#111";
    barEndereco.style.fontWeight = "700";
    barEndereco.style.fontSize = "0.8rem";
    barEndereco.style.overflow = "hidden";
    barContainer.appendChild(barEndereco);
    
    // Segmento Triagem
    const barTriagem = document.createElement("div");
    barTriagem.style.flex = `${pesoTriagemPercent}%`;
    barTriagem.style.background = "#d9d9d9";
    barTriagem.style.display = "flex";
    barTriagem.style.alignItems = "center";
    barTriagem.style.justifyContent = "center";
    barTriagem.style.color = "#111";
    barTriagem.style.fontWeight = "700";
    barTriagem.style.fontSize = "0.8rem";
    barTriagem.style.overflow = "hidden";
    barContainer.appendChild(barTriagem);
    
    // Overlay com percentual centralizado - cor com contraste inteligente
    const percentOverlay = document.createElement("div");
    percentOverlay.style.position = "absolute";
    percentOverlay.style.top = "50%";
    percentOverlay.style.left = "50%";
    percentOverlay.style.transform = "translate(-50%, -50%)";
    // Detectar qual segmento vai ser sobreposto para calcular contraste
    const bgColor2 = pesoEndPercent >= 50 ? "#fd7e14" : "#d9d9d9";
    percentOverlay.style.color = getContrastColor(bgColor2);
    percentOverlay.style.fontWeight = "700";
    percentOverlay.style.fontSize = "0.9rem";
    percentOverlay.style.textShadow = "0 0 3px rgba(255,255,255,0.8)";
    percentOverlay.textContent = `${pesoEndPercent}%`;
    barContainer.appendChild(percentOverlay);
    
    barWrapper.appendChild(barContainer);
    
    // Legenda
    const legend = document.createElement("div");
    legend.style.marginTop = "4px";
    legend.style.fontSize = "0.7rem";
    legend.style.display = "grid";
    legend.style.gap = "2px";
    legend.innerHTML = `
      <div style="display: flex; gap: 4px; align-items: center;">
        <div style="width: 8px; height: 8px; background: #fd7e14;"></div>
        <span>Endereçado: ${formatNumber(stats.pesoEnderecoado)} kg (${pesoEndPercent}%)</span>
      </div>
      <div style="display: flex; gap: 4px; align-items: center;">
        <div style="width: 8px; height: 8px; background: #d9d9d9;"></div>
        <span>Triagem: ${formatNumber(stats.pesoTriagem)} kg (${pesoTriagemPercent}%)</span>
      </div>
    `;
    barWrapper.appendChild(legend);
    
    pesoContainer.appendChild(barWrapper);
  }

  // 3. Misturas (barra horizontal empilhada 100%)
  const mixedContainer = $("statsMixedChart");
  if (mixedContainer) {
    mixedContainer.innerHTML = "";
    const items = collectEnderecoItems();
    const mixed = items.filter((it) => it.anyMixed).length;
    const clean = items.length - mixed;
    const mixedTotal = mixed + clean;
    const mixedPercent = mixedTotal > 0 ? Math.round((mixed / mixedTotal) * 100) : 0;
    const cleanPercent = 100 - mixedPercent;
    
    // Criar barra horizontal empilhada
    const barWrapper = document.createElement("div");
    barWrapper.style.display = "flex";
    barWrapper.style.flexDirection = "column";
    barWrapper.style.gap = "8px";
    
    const barContainer = document.createElement("div");
    barContainer.style.display = "flex";
    barContainer.style.height = "40px";
    barContainer.style.borderRadius = "4px";
    barContainer.style.overflow = "visible";
    barContainer.style.border = "1px solid #ccc";
    barContainer.style.background = "white";
    barContainer.style.position = "relative";
    
    // Segmento Com Mistura
    const barMixed = document.createElement("div");
    barMixed.style.flex = `${mixedPercent}%`;
    barMixed.style.background = "#06c6f0";
    barMixed.style.display = "flex";
    barMixed.style.alignItems = "center";
    barMixed.style.justifyContent = "center";
    barMixed.style.color = "#fff";
    barMixed.style.fontWeight = "700";
    barMixed.style.fontSize = "0.8rem";
    barMixed.style.overflow = "hidden";
    barContainer.appendChild(barMixed);
    
    // Segmento Limpos
    const barClean = document.createElement("div");
    barClean.style.flex = `${cleanPercent}%`;
    barClean.style.background = "#d9d9d9";
    barClean.style.display = "flex";
    barClean.style.alignItems = "center";
    barClean.style.justifyContent = "center";
    barClean.style.color = "#111";
    barClean.style.fontWeight = "700";
    barClean.style.fontSize = "0.8rem";
    barClean.style.overflow = "hidden";
    barContainer.appendChild(barClean);
    
    // Overlay com percentual centralizado - cor com contraste inteligente
    const percentOverlay = document.createElement("div");
    percentOverlay.style.position = "absolute";
    percentOverlay.style.top = "50%";
    percentOverlay.style.left = "50%";
    percentOverlay.style.transform = "translate(-50%, -50%)";
    // Detectar qual segmento vai ser sobreposto para calcular contraste
    const bgColor3 = mixedPercent >= 50 ? "#06c6f0" : "#d9d9d9";
    percentOverlay.style.color = getContrastColor(bgColor3);
    percentOverlay.style.fontWeight = "700";
    percentOverlay.style.fontSize = "0.9rem";
    percentOverlay.style.textShadow = "0 0 3px rgba(255,255,255,0.8)";
    percentOverlay.textContent = `${mixedPercent}%`;
    barContainer.appendChild(percentOverlay);
    
    barWrapper.appendChild(barContainer);
    
    // Legenda
    const legend = document.createElement("div");
    legend.style.marginTop = "4px";
    legend.style.fontSize = "0.7rem";
    legend.style.display = "grid";
    legend.style.gap = "2px";
    legend.innerHTML = `
      <div style="display: flex; gap: 4px; align-items: center;">
        <div style="width: 8px; height: 8px; background: #06c6f0;"></div>
        <span>Com mistura: ${mixed} (${mixedPercent}%)</span>
      </div>
      <div style="display: flex; gap: 4px; align-items: center;">
        <div style="width: 8px; height: 8px; background: #d9d9d9;"></div>
        <span>Limpos: ${clean} (${cleanPercent}%)</span>
      </div>
    `;
    barWrapper.appendChild(legend);
    
    mixedContainer.appendChild(barWrapper);
  }

  // 4. Conflito SHC (barra horizontal empilhada 100%)
  const conflictContainer = $("statsConflictChart");
  if (conflictContainer) {
    conflictContainer.innerHTML = "";
    const items = collectEnderecoItems();
    const conflict = items.filter((it) => it.shcConflict).length;
    const safe = items.length - conflict;
    const conflictTotal = conflict + safe;
    const conflictPercent = conflictTotal > 0 ? Math.round((conflict / conflictTotal) * 100) : 0;
    const safePercent = 100 - conflictPercent;
    
    // Criar barra horizontal empilhada
    const barWrapper = document.createElement("div");
    barWrapper.style.display = "flex";
    barWrapper.style.flexDirection = "column";
    barWrapper.style.gap = "8px";
    
    const barContainer = document.createElement("div");
    barContainer.style.display = "flex";
    barContainer.style.height = "40px";
    barContainer.style.borderRadius = "4px";
    barContainer.style.overflow = "visible";
    barContainer.style.border = "1px solid #ccc";
    barContainer.style.background = "white";
    barContainer.style.position = "relative";
    
    // Segmento Com Conflito
    const barConflict = document.createElement("div");
    barConflict.style.flex = `${conflictPercent}%`;
    barConflict.style.background = "#e74c3c";
    barConflict.style.display = "flex";
    barConflict.style.alignItems = "center";
    barConflict.style.justifyContent = "center";
    barConflict.style.color = "#fff";
    barConflict.style.fontWeight = "700";
    barConflict.style.fontSize = "0.8rem";
    barConflict.style.overflow = "hidden";
    barContainer.appendChild(barConflict);
    
    // Segmento Seguros
    const barSafe = document.createElement("div");
    barSafe.style.flex = `${safePercent}%`;
    barSafe.style.background = "#d9d9d9";
    barSafe.style.display = "flex";
    barSafe.style.alignItems = "center";
    barSafe.style.justifyContent = "center";
    barSafe.style.color = "#111";
    barSafe.style.fontWeight = "700";
    barSafe.style.fontSize = "0.8rem";
    barSafe.style.overflow = "hidden";
    barContainer.appendChild(barSafe);
    
    // Overlay com percentual centralizado - cor com contraste inteligente
    const percentOverlay = document.createElement("div");
    percentOverlay.style.position = "absolute";
    percentOverlay.style.top = "50%";
    percentOverlay.style.left = "50%";
    percentOverlay.style.transform = "translate(-50%, -50%)";
    // Detectar qual segmento vai ser sobreposto para calcular contraste
    const bgColor4 = conflictPercent >= 50 ? "#e74c3c" : "#d9d9d9";
    percentOverlay.style.color = getContrastColor(bgColor4);
    percentOverlay.style.fontWeight = "700";
    percentOverlay.style.fontSize = "0.9rem";
    percentOverlay.style.textShadow = "0 0 3px rgba(255,255,255,0.8)";
    percentOverlay.textContent = `${conflictPercent}%`;
    barContainer.appendChild(percentOverlay);
    
    barWrapper.appendChild(barContainer);
    
    // Legenda
    const legend = document.createElement("div");
    legend.style.marginTop = "4px";
    legend.style.fontSize = "0.7rem";
    legend.style.display = "grid";
    legend.style.gap = "2px";
    legend.innerHTML = `
      <div style="display: flex; gap: 4px; align-items: center;">
        <div style="width: 8px; height: 8px; background: #e74c3c;"></div>
        <span>Com conflito: ${conflict} (${conflictPercent}%)</span>
      </div>
      <div style="display: flex; gap: 4px; align-items: center;">
        <div style="width: 8px; height: 8px; background: #d9d9d9;"></div>
        <span>Seguros: ${safe} (${safePercent}%)</span>
      </div>
    `;
    barWrapper.appendChild(legend);
    
    conflictContainer.appendChild(barWrapper);
  }
}

function buildRecommendationCandidates() {
  const addrToRows = new Map();
  for (const r of state.rows) {
    const addr = normalizeEndereco(r.localOriginal || "");
    if (!REGEX_ENDERECO.test(addr)) continue;
    const list = addrToRows.get(addr) || [];
    list.push(r);
    addrToRows.set(addr, list);
  }

  const tripleToCandidates = new Map(); // tripleKey -> [{ endereco, analysis }]
  for (const [addr, rows] of addrToRows.entries()) {
    const a = analyzeEnderecoGroup(rows);
    if (a.destMixed || a.dataMixed || a.servMixed) continue;
    if (!a.destKey || !a.dataLabel || !a.servLabel) continue;
    const tripleKey = `${a.destKey}|${a.dataLabel}|${a.servLabel}`;
    const list = tripleToCandidates.get(tripleKey) || [];
    list.push({ endereco: addr, analysis: a });
    tripleToCandidates.set(tripleKey, list);
  }

  return tripleToCandidates;
}

function buildCargasGroups(rows, recommendationCandidates) {
  const groups = new Map(); // groupKey -> agg

  for (const r of rows) {
    const servKey = (r.servicoGroup || servicoGroupKey(r.servico || "") || "").toString();
    const groupKey = `${r.destino}|${r.data}|${servKey}`;

    let g = groups.get(groupKey);
    if (!g) {
      g = {
        groupKey,
        destino: r.destino,
        data: r.data,
        servKey,
        servicoLabel: servicoLabelFromKey(servKey),
        keys: [],
        ids: [],
        kg: 0,
        anyPendentes: false,
        localSet: new Set(),
        shcTokens: new Set(),
        recommendedEndereco: null,
      };
      groups.set(groupKey, g);
    }

    g.keys.push(r.key);
    g.ids.push(r.id);
    g.kg += Number.isFinite(r.peso) ? r.peso : 0;
    g.anyPendentes = g.anyPendentes || !isEnderecado(r);
    if (r.localOriginal) g.localSet.add(normalizeEndereco(r.localOriginal));
    for (const t of r.shcTokens || []) g.shcTokens.add(t);
  }

  const list = Array.from(groups.values()).map((g) => {
    const tripleKey = `${destinoEquivKey(g.destino)}|${g.data}|${g.servKey}`;
    const candidates = recommendationCandidates.get(tripleKey) || [];

    let best = null; // { endereco, score }
    for (const c of candidates) {
      const combined = new Set(c.analysis?.shcTokens ? Array.from(c.analysis.shcTokens) : []);
      for (const t of g.shcTokens) combined.add(t);
      if (analyzeShcConflict(combined)) continue;

      const score = (c.analysis?.shcMixed ? 3 : 0) + (c.analysis?.shcConflict ? 10 : 0);
      if (!best || score < best.score || (score === best.score && c.endereco.localeCompare(best.endereco) < 0)) {
        best = { endereco: c.endereco, score };
      }
    }

    const reco = best ? best.endereco : null;
    const localVals = Array.from(g.localSet).filter(Boolean);
    const localLabel = localVals.length === 1 ? localVals[0] : localVals.length > 1 ? "MISTO" : "";
    const shcLabel = shcLabelFromSet(g.shcTokens);

    return {
      ...g,
      count: g.keys.length,
      firstId: g.ids[0] || "",
      localLabel,
      shcLabel,
      recommendedEndereco: reco,
    };
  });

  const score = (g) => {
    let s = 0;
    if (g.recommendedEndereco) s += 200;
    if (g.anyPendentes) s += 10;
    s += Math.min(g.count, 50);
    return s;
  };

  list.sort((a, b) => score(b) - score(a) || a.destino.localeCompare(b.destino) || a.data.localeCompare(b.data) || a.servKey.localeCompare(b.servKey));
  return list;
}

function renderCargaModal(row) {
  if (!row) return;

  const shcList = shcLabelFromSet(row.shcTokens);
  setText("cargaModalTitle", `Carga ${row.id}`);
  setText("cargaModalSub", `${row.destino} | ${isEnderecado(row) ? "OK" : "PEND"} | Kg ${formatNumber(row.peso)} | SHC: ${shcList}`);

  const body = $("cargaModalBody");
  if (!body) return;

  const raw = state.rawRows[row.rawIndex] || [];
  const header = state.hasHeader ? state.rawRows[0] || [] : [];
  const labels = state.hasHeader && header.length ? header : raw.map((_, i) => `Col ${i}`);

  const frag = document.createDocumentFragment();

  const top = document.createElement("div");
  top.className = "row g-2";
  top.innerHTML = `
    <div class="col-6 col-lg-3"><div class="small text-muted">ID</div><div class="mono">${escapeHtml(row.id)}</div></div>
    <div class="col-6 col-lg-3"><div class="small text-muted">Destino</div><div class="mono">${escapeHtml(row.destino)}</div></div>
    <div class="col-6 col-lg-3"><div class="small text-muted">Local</div><div class="mono">${escapeHtml(row.localOriginal || "")}</div></div>
    <div class="col-6 col-lg-3"><div class="small text-muted">Peso (kg)</div><div class="mono">${escapeHtml(formatNumber(row.peso))}</div></div>
  `;
  frag.appendChild(top);

  const hr = document.createElement("hr");
  hr.className = "my-3";
  frag.appendChild(hr);

  const tableWrap = document.createElement("div");
  tableWrap.className = "table-responsive";
  const tbl = document.createElement("table");
  tbl.className = "table table-sm mb-0";
  const tbody = document.createElement("tbody");

  for (let i = 0; i < labels.length; i++) {
    const key = fixUtf8Mojibake((labels[i] ?? `Col ${i}`).toString()).trim() || `Col ${i}`;
    const val = fixUtf8Mojibake((raw[i] ?? "").toString());
    if (!key && !val) continue;

    const tr = document.createElement("tr");
    const th = document.createElement("th");
    th.className = "text-muted";
    th.style.width = "40%";
    th.textContent = key;
    const td = document.createElement("td");
    td.className = "mono";
    td.textContent = val ?? "";
    tr.appendChild(th);
    tr.appendChild(td);
    tbody.appendChild(tr);
  }

  tbl.appendChild(tbody);
  tableWrap.appendChild(tbl);
  frag.appendChild(tableWrap);

  body.innerHTML = "";
  body.appendChild(frag);
}

function openCargaModalByKey(key) {
  const row = state.rows.find((r) => r.key === key) || null;
  if (!row) return;
  renderCargaModal(row);

  const el = $("cargaModal");
  if (!el || typeof bootstrap === "undefined") return;
  bootstrap.Modal.getOrCreateInstance(el).show();
}

function openCargaModalById(id) {
  const idNorm = normalizeId(id);
  const row = state.rows.find((r) => normalizeId(r.id) === idNorm) || null;
  if (!row) return;
  openCargaModalByKey(row.key);
}

function renderEnderecoModal(endereco, rows) {
  setText("enderecoModalTitle", `Endereço ${endereco}`);
  const a = analyzeEnderecoGroup(rows);
  setText("enderecoModalSub", `${rows.length} carga(s) | Kg ${formatNumber(a.kg)}`);
  setText(
    "enderecoModalResumo",
    `Destino: ${a.destLabel || "-"} | Data: ${a.dataLabel || "-"} | Serviço: ${a.servLabel || "-"} | SHC: ${shcLabelFromSet(a.shcTokens)}`
  );

  const alerts = [];
  if (a.destMixed) alerts.push("Destino misturado");
  if (a.dataMixed) alerts.push("Data misturada");
  if (a.servMixed) alerts.push("Serviço misturado");
  if (a.shcMixed) alerts.push("SHC misturado");
  if (a.shcConflict) alerts.push("Possível conflito SHC (ELI/ELM/WET)");
  setText("enderecoModalAlertas", alerts.length ? alerts.join(" | ") : "OK");

  const tbody = $("tblEnderecoAwbs")?.querySelector("tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  const sorted = rows.slice().sort((x, y) => normalizeId(x.id).localeCompare(normalizeId(y.id)));
  const frag = document.createDocumentFragment();

  for (const r of sorted) {
    const tr = document.createElement("tr");

    const tdId = document.createElement("td");
    tdId.className = "mono";
    const idLink = document.createElement("a");
    idLink.href = "#";
    idLink.className = "text-reset ds-link";
    idLink.textContent = r.id;
    idLink.addEventListener("click", (e) => {
      e.preventDefault();
      openCargaModalByKey(r.key);
    });
    tdId.appendChild(idLink);

    const tdDest = document.createElement("td");
    tdDest.className = "mono";
    tdDest.textContent = r.destino || "";

    const tdData = document.createElement("td");
    tdData.className = "mono";
    tdData.textContent = r.data || "";

    const tdServ = document.createElement("td");
    tdServ.className = "mono";
    tdServ.textContent = r.servico || "";

    const tdShc = document.createElement("td");
    tdShc.className = "mono";
    tdShc.textContent = shcLabelFromSet(r.shcTokens);

    const tdKg = document.createElement("td");
    tdKg.className = "mono text-end";
    tdKg.textContent = formatNumber(r.peso);

    // Coluna de endereço sugerido
    const tdSugerido = document.createElement("td");
    tdSugerido.className = "mono";
    const enderecoSugerido = buscarEnderecoParaAwb(r, endereco);
    
    if (enderecoSugerido) {
      const link = document.createElement("a");
      link.href = "#";
      link.className = "text-reset ds-link";
      link.textContent = enderecoSugerido;
      link.addEventListener("click", (e) => {
        e.preventDefault();
        openEnderecoModal(enderecoSugerido);
      });
      tdSugerido.appendChild(link);
    } else {
      tdSugerido.style.color = "var(--ds-muted)";
      tdSugerido.style.fontStyle = "italic";
      tdSugerido.textContent = "Sem sugestão";
    }

    tr.appendChild(tdId);
    tr.appendChild(tdDest);
    tr.appendChild(tdData);
    tr.appendChild(tdServ);
    tr.appendChild(tdShc);
    tr.appendChild(tdKg);
    tr.appendChild(tdSugerido);

    frag.appendChild(tr);
  }

  tbody.appendChild(frag);
}

function buscarEnderecoParaAwb(row, enderecoAtual) {
  // Buscar um endereço que tenha AWBs com mesmo destino, data e serviço
  const items = collectEnderecoItems();
  
  for (const item of items) {
    // Pular o endereço atual
    if (item.endereco === enderecoAtual) continue;
    
    // Procurar uma AWB com os mesmos dados
    for (const itemRow of item.rows) {
      if (
        normalizeUpper(itemRow.destino) === normalizeUpper(row.destino) &&
        itemRow.data === row.data &&
        normalizeServicoCell(itemRow.servico) === normalizeServicoCell(row.servico)
      ) {
        return item.endereco;
      }
    }
  }
  
  return null;
}

function openEnderecoModal(endereco) {
  const addr = normalizeEndereco(endereco);
  if (!REGEX_ENDERECO.test(addr)) return;
  const rows = state.rows.filter((r) => normalizeEndereco(r.localOriginal || "") === addr);
  renderEnderecoModal(addr, rows);

  const el = $("enderecoModal");
  if (!el || typeof bootstrap === "undefined") return;
  bootstrap.Modal.getOrCreateInstance(el).show();
}

function renderGrupoModal(group) {
  setText("grupoModalTitle", "Grupo de cargas");
  setText("grupoModalSub", `${group.count} carga(s) | Kg ${formatNumber(group.kg)}`);
  setText("grupoModalResumo", `Destino: ${group.destino || "-"} | Data: ${group.data || "-"} | Serviço: ${group.servicoLabel || "-"}`);

  const recoEl = $("grupoModalReco");
  if (recoEl) {
    recoEl.innerHTML = "";
    if (group.recommendedEndereco) {
      const a = document.createElement("a");
      a.href = "#";
      a.className = "text-reset ds-link";
      a.textContent = group.recommendedEndereco;
      a.addEventListener("click", (e) => {
        e.preventDefault();
        openEnderecoModal(group.recommendedEndereco);
      });
      recoEl.appendChild(a);
    } else {
      recoEl.textContent = "-";
    }
  }

  const tbody = $("tblGrupoAwbs")?.querySelector("tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  const rows = group.keys.map((k) => state.rows.find((r) => r.key === k)).filter(Boolean);
  rows.sort((x, y) => normalizeId(x.id).localeCompare(normalizeId(y.id)));

  const frag = document.createDocumentFragment();
  for (const r of rows) {
    const tr = document.createElement("tr");

    const tdId = document.createElement("td");
    tdId.className = "mono";
    const idLink = document.createElement("a");
    idLink.href = "#";
    idLink.className = "text-reset ds-link";
    idLink.textContent = r.id;
    idLink.addEventListener("click", (e) => {
      e.preventDefault();
      openCargaModalByKey(r.key);
    });
    tdId.appendChild(idLink);

    const tdDest = document.createElement("td");
    tdDest.className = "mono";
    tdDest.textContent = r.destino || "";

    const tdData = document.createElement("td");
    tdData.className = "mono";
    tdData.textContent = r.data || "";

    const tdServ = document.createElement("td");
    tdServ.className = "mono";
    tdServ.textContent = r.servico || "";

    const tdShc = document.createElement("td");
    tdShc.className = "mono";
    tdShc.textContent = shcLabelFromSet(r.shcTokens);

    const tdKg = document.createElement("td");
    tdKg.className = "mono text-end";
    tdKg.textContent = formatNumber(r.peso);

    tr.appendChild(tdId);
    tr.appendChild(tdDest);
    tr.appendChild(tdData);
    tr.appendChild(tdServ);
    tr.appendChild(tdShc);
    tr.appendChild(tdKg);

    frag.appendChild(tr);
  }

  tbody.appendChild(frag);
}

function openGrupoModalByGroupKey(groupKey) {
  const recommendationCandidates = buildRecommendationCandidates();
  const groups = buildCargasGroups(filteredRows(), recommendationCandidates);
  const group = groups.find((g) => g.groupKey === groupKey);
  if (!group) return;
  renderGrupoModal(group);

  const el = $("grupoModal");
  if (!el || typeof bootstrap === "undefined") return;
  bootstrap.Modal.getOrCreateInstance(el).show();
}

function renderCargasPendentesTable() {
  const tbody = $("tblCargas")?.querySelector("tbody");
  if (!tbody) return;

  const rows = filteredRows();
  const recommendationCandidates = buildRecommendationCandidates();
  const groups = buildCargasGroups(rows, recommendationCandidates);

  // Ordenação: Saúde > Urgentes > Expressos > Ecos, depois por endereço recomendado
  const servicePriority = (servKey) => {
    const k = (servKey || "").toUpperCase();
    if (k.includes("SAUDE")) return 0;
    if (k.includes("URGENTE")) return 1;
    if (k === "EXPRESSO") return 2;
    return 3; // Eco/Moda e outros
  };

  groups.sort((a, b) => {
    const priA = servicePriority(a.servKey);
    const priB = servicePriority(b.servKey);
    if (priA !== priB) return priA - priB;
    
    // Depois por endereço recomendado (com recomendação vem antes)
    if ((a.recommendedEndereco ? 1 : 0) !== (b.recommendedEndereco ? 1 : 0)) {
      return (b.recommendedEndereco ? 1 : 0) - (a.recommendedEndereco ? 1 : 0);
    }
    
    if (a.recommendedEndereco && b.recommendedEndereco) {
      return a.recommendedEndereco.localeCompare(b.recommendedEndereco);
    }
    
    return a.destino.localeCompare(b.destino);
  });

  const MAX_RENDER = 2500;
  const limited = groups.slice(0, MAX_RENDER);

  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();

  for (const g of limited) {
    const tr = document.createElement("tr");
    if (g.recommendedEndereco) tr.classList.add("ds-reco");

    const tdId = document.createElement("td");
    tdId.className = "mono";
    if (g.count > 1) {
      const grpLink = document.createElement("a");
      grpLink.href = "#";
      grpLink.className = "text-reset ds-link";
      grpLink.textContent = `x${g.count}`;
      grpLink.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        openGrupoModalByGroupKey(g.groupKey);
      });
      tdId.appendChild(grpLink);
    } else {
      const idLink = document.createElement("a");
      idLink.href = "#";
      idLink.className = "text-reset ds-link";
      idLink.textContent = g.firstId || "";
      idLink.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (g.keys[0]) openCargaModalByKey(g.keys[0]);
      });
      tdId.appendChild(idLink);
    }

    const tdDest = document.createElement("td");
    tdDest.className = "mono";
    tdDest.textContent = g.destino;

    const tdData = document.createElement("td");
    tdData.className = "mono";
    tdData.textContent = g.data;

    const tdServ = document.createElement("td");
    tdServ.className = "mono";
    tdServ.textContent = g.servicoLabel;

    const tdShc = document.createElement("td");
    tdShc.className = "mono";
    // Parse SHC labels e criar com tooltips
    const shcParts = (g.shcLabel || "-").split(", ");
    if (shcParts[0] === "-") {
      tdShc.textContent = "-";
    } else {
      const shcFrag = document.createDocumentFragment();
      for (let i = 0; i < shcParts.length; i++) {
        if (i > 0) {
          const comma = document.createElement("span");
          comma.textContent = ", ";
          shcFrag.appendChild(comma);
        }
        const code = shcParts[i].trim();
        shcFrag.appendChild(buildShcElement(code));
      }
      tdShc.appendChild(shcFrag);
    }

    const tdReco = document.createElement("td");
    tdReco.className = "mono";
    if (g.recommendedEndereco) {
      const a = document.createElement("a");
      a.href = "#";
      a.className = "text-reset ds-link";
      a.textContent = g.recommendedEndereco;
      a.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        openEnderecoModal(g.recommendedEndereco);
      });
      tdReco.appendChild(a);
    } else {
      tdReco.textContent = "";
    }

    const tdKg = document.createElement("td");
    tdKg.className = "mono text-end";
    tdKg.textContent = formatNumber(g.kg);

    tr.appendChild(tdId);
    tr.appendChild(tdDest);
    tr.appendChild(tdData);
    tr.appendChild(tdServ);
    tr.appendChild(tdShc);
    tr.appendChild(tdReco);
    tr.appendChild(tdKg);

    tr.addEventListener("click", () => {
      if (g.count > 1) openGrupoModalByGroupKey(g.groupKey);
      else if (g.keys[0]) openCargaModalByKey(g.keys[0]);
    });

    frag.appendChild(tr);
  }

  tbody.appendChild(frag);
  const volumesTotal = rows.reduce((sum, r) => sum + (Math.max(1, Number.isFinite(r.pecas) ? r.pecas : 1)), 0);
  setText("tableStamp", `${rows.length} cargas | ${groups.length} linhas`);
  setText("tabBadgeCargas", `${rows.length} cargas, ${volumesTotal} volumes`);
}

function renderEnderecosTable() {
  const tbody = $("tblEnderecos")?.querySelector("tbody");
  if (!tbody) return;

  const items = collectEnderecoItems();

  const score = (it) => {
    let s = 0;
    if (it.anyMixed) s += 10;
    if (it.destMixed) s += 3;
    if (it.dataMixed) s += 3;
    if (it.servMixed) s += 3;
    if (it.shcMixed) s += 2;
    if (it.shcConflict) s += 4;
    return s;
  };

  items.sort((a, b) => {
    if (a.dataMixed !== b.dataMixed) return (b.dataMixed ? 1 : 0) - (a.dataMixed ? 1 : 0);
    if (a.servMixed !== b.servMixed) return (b.servMixed ? 1 : 0) - (a.servMixed ? 1 : 0);
    const shcCountA = a.shcTokens ? a.shcTokens.size : 0;
    const shcCountB = b.shcTokens ? b.shcTokens.size : 0;
    if (shcCountA !== shcCountB) return shcCountB - shcCountA;
    const scoreA = score(a);
    const scoreB = score(b);
    if (scoreA !== scoreB) return scoreB - scoreA;
    return a.endereco.localeCompare(b.endereco);
  });

  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();

  for (const it of items) {
    const tr = document.createElement("tr");
    if (it.shcConflict) tr.classList.add("ds-risk");
    else if (it.anyMixed) tr.classList.add("ds-mixed");

    const tdNome = document.createElement("td");
    tdNome.className = "mono";
    const link = document.createElement("a");
    link.href = "#";
    link.className = "text-reset ds-link";
    link.textContent = it.endereco;
    link.addEventListener("click", (e) => {
      e.preventDefault();
      openEnderecoModal(it.endereco);
    });
    tdNome.appendChild(link);

    const tdDest = document.createElement("td");
    tdDest.className = "mono";
    tdDest.textContent = it.destLabel || "";

    const tdData = document.createElement("td");
    tdData.className = "mono";
    tdData.textContent = it.dataLabel || "";

    const tdDataMix = document.createElement("td");
    tdDataMix.className = "mono";
    if (it.dataMixed) tdDataMix.appendChild(buildBadge("SIM", "text-bg-warning"));
    else tdDataMix.textContent = "-";

    const tdServMix = document.createElement("td");
    tdServMix.className = "mono";
    if (it.servMixed) tdServMix.appendChild(buildBadge("SIM", "text-bg-warning"));
    else tdServMix.textContent = "-";

    const tdShc = document.createElement("td");
    tdShc.className = "mono";
    const shcLabel = shcLabelFromSet(it.shcTokens);
    if (shcLabel && shcLabel !== "-") {
      const shcParts = shcLabel.split(", ");
      const shcFrag = document.createDocumentFragment();
      for (let i = 0; i < shcParts.length; i++) {
        if (i > 0) {
          const comma = document.createElement("span");
          comma.textContent = ", ";
          shcFrag.appendChild(comma);
        }
        const code = shcParts[i].trim();
        shcFrag.appendChild(buildShcElement(code));
      }
      tdShc.appendChild(shcFrag);
    } else {
      tdShc.textContent = "-";
    }

    const tdKg = document.createElement("td");
    tdKg.className = "mono text-end";
    tdKg.textContent = formatNumber(it.kg);

    tr.appendChild(tdNome);
    tr.appendChild(tdDest);
    tr.appendChild(tdData);
    tr.appendChild(tdDataMix);
    tr.appendChild(tdServMix);
    tr.appendChild(tdShc);
    tr.appendChild(tdKg);
    frag.appendChild(tr);
  }

  tbody.appendChild(frag);
  setText("enderecosStamp", `${items.length} endereços`);
  setText("tabBadgeEnderecos", String(items.length));
}

function renderLotesTable() {
  const tbody = $("tblLotes")?.querySelector("tbody");
  if (!tbody) return;

  const singleOrMixed = (set) => {
    const vals = Array.from(set).filter((x) => (x ?? "").toString().trim() !== "");
    if (vals.length === 0) return "-";
    if (vals.length === 1) return vals[0];
    return "MISTO";
  };

  const byId = new Map();
  for (const r of state.rows) {
    const idKey = normalizeId(r.id);
    if (!idKey) continue;
    let agg = byId.get(idKey);
    if (!agg) {
      agg = {
        id: r.id,
        pecasTotal: 0,
        destinos: new Set(),
        datas: new Set(),
        servicoGroups: new Set(),
        shcTokens: new Set(),
        localizacao: r.localOriginalRaw || "", // Usar valor bruto
      };
      byId.set(idKey, agg);
    }
    const pecas = Math.max(1, Number.isFinite(r.pecas) ? r.pecas : 1);
    agg.pecasTotal += pecas;
    if (r.destino) agg.destinos.add(r.destino);
    agg.datas.add((r.data || "").toString());
    const servKey = (r.servicoGroup || servicoGroupKey(r.servico || "") || "").toString();
    if (servKey) agg.servicoGroups.add(servKey);
    for (const t of r.shcTokens || []) agg.shcTokens.add(t);
  }

  const items = Array.from(byId.values())
    .filter((a) => a.pecasTotal > 1)
    .filter((a) => !/^1k/i.test((a.localizacao || "").trim()))
    .filter((a) => parseMultiplosEnderecos(a.localizacao).size > 1)
    .map((a) => {
      const destino = destinoLabelFromRawSet(a.destinos) || "-";
      const data = singleOrMixed(a.datas);
      const servKey = singleOrMixed(a.servicoGroups);
      const servico = servKey === "MISTO" ? "MISTO" : servicoLabelFromKey(servKey) || "-";
      return {
        id: a.id,
        pecasTotal: a.pecasTotal,
        destino,
        data,
        servico,
        shc: shcLabelFromSet(a.shcTokens),
        localizacao: a.localizacao,
        isSplit: true, // Simplificar: sempre mostrar como split (lote)
      };
    })
    .sort((a, b) => {
      if (a.isSplit !== b.isSplit) return a.isSplit ? -1 : 1;
      return b.pecasTotal - a.pecasTotal || normalizeId(a.id).localeCompare(normalizeId(b.id));
    });

  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();

  for (const it of items) {
    const tr = document.createElement("tr");
    if (it.isSplit) tr.classList.add("ds-mixed");

    const tdPecas = document.createElement("td");
    tdPecas.className = "mono";
    tdPecas.textContent = String(it.pecasTotal);

    const tdId = document.createElement("td");
    tdId.className = "mono";
    const idLink = document.createElement("a");
    idLink.href = "#";
    idLink.className = "text-reset ds-link";
    idLink.textContent = it.id;
    idLink.addEventListener("click", (e) => {
      e.preventDefault();
      openCargaModalById(it.id);
    });
    tdId.appendChild(idLink);

    const tdAddr = document.createElement("td");
    tdAddr.className = "mono";
    tdAddr.textContent = formatLocalizacao(it.localizacao) || "-";

    const tdDest = document.createElement("td");
    tdDest.className = "mono";
    tdDest.textContent = it.destino === "-" ? "" : it.destino;

    const tdData = document.createElement("td");
    tdData.className = "mono";
    tdData.textContent = it.data === "-" ? "" : it.data;

    const tdShc = document.createElement("td");
    tdShc.className = "mono";
    const shcLabel = it.shc || "-";
    if (shcLabel && shcLabel !== "-") {
      const shcParts = shcLabel.split(", ");
      const shcFrag = document.createDocumentFragment();
      for (let i = 0; i < shcParts.length; i++) {
        if (i > 0) {
          const comma = document.createElement("span");
          comma.textContent = ", ";
          shcFrag.appendChild(comma);
        }
        const code = shcParts[i].trim();
        shcFrag.appendChild(buildShcElement(code));
      }
      tdShc.appendChild(shcFrag);
    } else {
      tdShc.textContent = "-";
    }

    const tdServ = document.createElement("td");
    tdServ.className = "mono";
    tdServ.textContent = it.servico === "-" ? "" : it.servico;

    tr.appendChild(tdPecas);
    tr.appendChild(tdId);
    tr.appendChild(tdDest);
    tr.appendChild(tdAddr);
    tr.appendChild(tdData);
    tr.appendChild(tdShc);
    tr.appendChild(tdServ);
    frag.appendChild(tr);
  }

  tbody.appendChild(frag);
  setText("lotesStamp", `${items.length}`);
  setText("tabBadgeLotes", String(items.length));
}

function renderPereciveisTable() {
  const tbody = $("tblPerecivel")?.querySelector("tbody");
  if (!tbody) return;

  // Função para ordenar por prioridade de serviço
  const servicePriority = (servKey) => {
    const k = (servKey || "").toUpperCase();
    if (k === "SAUDE") return 0;
    if (k.includes("URGENTE")) return 1;
    if (k === "EXPRESSO") return 2;
    return 3; // ECO_MODA e outros
  };

  // Coletar todas as AWBs com SHC PER
  const pereceisMap = new Map();
  
  for (const r of state.rows) {
    // Verificar se possui SHC PER entre os tokens
    const temPER = r.shcTokens && Array.from(r.shcTokens).some(t => normalizeUpper(t) === "PER");
    if (!temPER) continue;
    
    const idKey = normalizeId(r.id);
    if (!idKey) continue;
    
    if (!pereceisMap.has(idKey)) {
      pereceisMap.set(idKey, {
        id: r.id,
        pecasTotal: 0,
        destino: r.destino || "",
        enderecos: new Set(),
        datas: new Set(),
        servicos: new Set(),
        shcTokens: new Set(),
      });
    }
    
    const agg = pereceisMap.get(idKey);
    const pecas = Math.max(1, Number.isFinite(r.pecas) ? r.pecas : 1);
    agg.pecasTotal += pecas;
    
    const addr = normalizeEndereco(r.localOriginal || "");
    if (REGEX_ENDERECO.test(addr)) {
      agg.enderecos.add(addr);
    }
    
    agg.datas.add((r.data || "").toString());
    
    const servKey = (r.servicoGroup || servicoGroupKey(r.servico || "") || "").toString();
    if (servKey) agg.servicos.add(servKey);
    
    // Coletar todos os SHCs (não apenas PER)
    for (const t of r.shcTokens || []) agg.shcTokens.add(t);
  }

  const items = Array.from(pereceisMap.values())
    .map((a) => {
      const endereco = Array.from(a.enderecos).length === 1 
        ? Array.from(a.enderecos)[0] 
        : (Array.from(a.enderecos).length > 1 ? "MÚLTIPLOS" : "-");
      
      const data = a.datas.size === 1 
        ? Array.from(a.datas)[0] 
        : (a.datas.size > 1 ? "MÚLTIPLOS" : "-");
      
      const servicos = Array.from(a.servicos);
      const servicoKey = servicos.length === 1 ? servicos[0] : (servicos.length > 1 ? "MÚLTIPLOS" : "");
      const servico = servicoLabelFromKey(servicoKey) || "-";
      const shc = shcLabelFromSet(a.shcTokens);
      
      return {
        id: a.id,
        pecasTotal: a.pecasTotal,
        destino: a.destino || "-",
        endereco,
        data,
        servico,
        shc,
        servicoKey,
      };
    })
    .sort((a, b) => {
      // Ordenar por prioridade de serviço primeiro
      const priA = servicePriority(a.servicoKey);
      const priB = servicePriority(b.servicoKey);
      if (priA !== priB) return priA - priB;
      
      // Depois por quantidade de peças (maior primeiro)
      if (a.pecasTotal !== b.pecasTotal) return b.pecasTotal - a.pecasTotal;
      
      // Por fim, por ID
      return normalizeId(a.id).localeCompare(normalizeId(b.id));
    });

  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();

  for (const it of items) {
    const tr = document.createElement("tr");

    const tdId = document.createElement("td");
    tdId.className = "mono";
    const idLink = document.createElement("a");
    idLink.href = "#";
    idLink.className = "text-reset ds-link";
    idLink.textContent = it.id;
    idLink.addEventListener("click", (e) => {
      e.preventDefault();
      openCargaModalById(it.id);
    });
    tdId.appendChild(idLink);

    const tdPecas = document.createElement("td");
    tdPecas.className = "mono";
    tdPecas.textContent = String(it.pecasTotal);

    const tdDest = document.createElement("td");
    tdDest.className = "mono";
    tdDest.textContent = it.destino === "-" ? "" : it.destino;

    const tdEndereco = document.createElement("td");
    tdEndereco.className = "mono";
    tdEndereco.textContent = it.endereco;

    const tdData = document.createElement("td");
    tdData.className = "mono";
    tdData.textContent = it.data === "-" ? "" : it.data;

    const tdShc = document.createElement("td");
    tdShc.className = "mono";
    const shcLabel = it.shc || "-";
    if (shcLabel && shcLabel !== "-") {
      const shcParts = shcLabel.split(", ");
      const shcFrag = document.createDocumentFragment();
      for (let i = 0; i < shcParts.length; i++) {
        if (i > 0) {
          const comma = document.createElement("span");
          comma.textContent = ", ";
          shcFrag.appendChild(comma);
        }
        const code = shcParts[i].trim();
        shcFrag.appendChild(buildShcElement(code));
      }
      tdShc.appendChild(shcFrag);
    } else {
      tdShc.textContent = "-";
    }

    const tdServ = document.createElement("td");
    tdServ.className = "mono";
    tdServ.textContent = it.servico === "-" ? "" : it.servico;

    // Calcular troca de gelo (48 horas após data da carga)
    const dataMs = parsePtDateToMs(it.data);
    let trocaGeloHoras = "-";
    let trocaGeloClass = "";
    if (dataMs && Number.isFinite(dataMs)) {
      const agora = Date.now();
      const diffMs = agora - dataMs;
      const diffHoras = Math.round(diffMs / (60 * 60 * 1000));
      const horasParaTroca = diffHoras - 48;
      if (horasParaTroca >= 0) {
        // Já passou do tempo de troca
        trocaGeloHoras = `-${horasParaTroca}h`;
        trocaGeloClass = "text-danger";
      } else {
        // Ainda faltam horas
        trocaGeloHoras = `${Math.abs(horasParaTroca)}h`;
      }
    }
    const tdTrocaGelo = document.createElement("td");
    tdTrocaGelo.className = `mono ${trocaGeloClass}`;
    tdTrocaGelo.textContent = trocaGeloHoras;

    tr.appendChild(tdId);
    tr.appendChild(tdPecas);
    tr.appendChild(tdDest);
    tr.appendChild(tdEndereco);
    tr.appendChild(tdData);
    tr.appendChild(tdTrocaGelo);
    tr.appendChild(tdShc);
    tr.appendChild(tdServ);
    frag.appendChild(tr);
  }

  tbody.appendChild(frag);
  setText("perecivelStamp", `${items.length}`);
  setText("tabBadgePerecivel", String(items.length));
}

function renderTop3Destinos() {
  const body = $("top3DestinosBody");
  if (!body) return;

  // Contar peso por destino (apenas cargas endereçadas em bolsões)
  const destinos = new Map();
  
  for (const r of state.rows) {
    const destNorm = normalizeUpper(r.destino || "");
    if (DESTINOS_LOJA.has(destNorm)) continue;
    if (isCarretaLocation(r.localOriginal || "")) continue;
    if (!isEnderecado(r)) continue;
    const addr = normalizeEndereco(r.localOriginal || "");
    if (!isValidBolsaoAddress(addr)) continue;
    
    const peso = Number.isFinite(r.peso) ? r.peso : 0;
    const destKey = r.destino || "-";
    const current = destinos.get(destKey) || { destino: destKey, peso: 0, count: 0 };
    current.peso += peso;
    current.count += 1;
    destinos.set(destKey, current);
  }

  const items = Array.from(destinos.values())
    .sort((a, b) => b.peso - a.peso);

  body.innerHTML = "";
  
  if (items.length === 0) {
    const div = document.createElement("div");
    div.className = "text-muted small";
    div.textContent = "Sem dados de destinos endereçados";
    body.appendChild(div);
    return;
  }

  // Cores para o pie chart
  const colors = [
    "#111111", "#5a5a5a", "#b9b9b9", "#d6d6d6", "#f0f0f0",
    "#0d6efd", "#6c757d", "#198754", "#ffc107", "#dc3545",
    "#20c997", "#0dcaf0", "#f8f9fa", "#e9ecef", "#adb5bd"
  ];

  const totalPeso = items.reduce((sum, i) => sum + i.peso, 0);
  const maxItems = Math.min(items.length, colors.length);

  const wrapper = document.createElement("div");
  wrapper.style.display = "grid";
  wrapper.style.gridTemplateColumns = "auto 1fr";
  wrapper.style.gap = "16px";
  wrapper.style.alignItems = "start";

  // Pie chart em SVG
  const svgWrapper = document.createElement("div");
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 100 100");
  svg.setAttribute("width", "140");
  svg.setAttribute("height", "140");

  let currentAngle = -90;
  for (let i = 0; i < maxItems; i++) {
    const item = items[i];
    const sliceAngle = (item.peso / totalPeso) * 360;
    const startAngle = currentAngle * (Math.PI / 180);
    const endAngle = (currentAngle + sliceAngle) * (Math.PI / 180);

    const x1 = 50 + 40 * Math.cos(startAngle);
    const y1 = 50 + 40 * Math.sin(startAngle);
    const x2 = 50 + 40 * Math.cos(endAngle);
    const y2 = 50 + 40 * Math.sin(endAngle);

    const largeArc = sliceAngle > 180 ? 1 : 0;

    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", `M 50 50 L ${x1} ${y1} A 40 40 0 ${largeArc} 1 ${x2} ${y2} Z`);
    path.setAttribute("fill", colors[i % colors.length]);
    path.setAttribute("stroke", "#fff");
    path.setAttribute("stroke-width", "1");
    
    svg.appendChild(path);
    currentAngle += sliceAngle;
  }

  // Adicionar círculo central para donut effect (opcional)
  const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
  circle.setAttribute("cx", "50");
  circle.setAttribute("cy", "50");
  circle.setAttribute("r", "20");
  circle.setAttribute("fill", "#fff");
  svg.appendChild(circle);

  svgWrapper.appendChild(svg);
  wrapper.appendChild(svgWrapper);

  // Legenda
  const legend = document.createElement("div");
  legend.style.display = "grid";
  legend.style.gap = "6px";
  legend.style.fontSize = "0.75rem";
  legend.style.maxHeight = "140px";
  legend.style.overflowY = "auto";

  for (let i = 0; i < maxItems; i++) {
    const item = items[i];
    const percentage = ((item.peso / totalPeso) * 100).toFixed(1);

    const row = document.createElement("div");
    row.style.display = "flex";
    row.style.gap = "6px";
    row.style.alignItems = "center";

    const colorBox = document.createElement("div");
    colorBox.style.width = "12px";
    colorBox.style.height = "12px";
    colorBox.style.backgroundColor = colors[i % colors.length];
    colorBox.style.flexShrink = "0";

    const label = document.createElement("div");
    label.style.display = "flex";
    label.style.flexDirection = "column";
    label.style.gap = "2px";
    label.style.flex = "1";

    const destName = document.createElement("span");
    destName.style.fontWeight = "700";
    destName.textContent = escapeHtml(item.destino);

    const meta = document.createElement("span");
    meta.className = "text-muted";
    meta.style.fontSize = "0.7rem";
    meta.innerHTML = `${formatNumber(item.peso)} kg • ${percentage}% • ${item.count} ${item.count === 1 ? 'carga' : 'cargas'}`;

    label.appendChild(destName);
    label.appendChild(meta);

    row.appendChild(colorBox);
    row.appendChild(label);
    legend.appendChild(row);
  }

  // Se houver mais itens que cores, adicionar aviso
  if (items.length > maxItems) {
    const moreRow = document.createElement("div");
    moreRow.className = "small text-muted";
    moreRow.textContent = `... e ${items.length - maxItems} mais`;
    moreRow.style.paddingTop = "6px";
    moreRow.style.borderTop = "1px solid #ddd";
    legend.appendChild(moreRow);
  }

  wrapper.appendChild(legend);
  body.appendChild(wrapper);
}

function renderAll() {
  updateAppStamp();
  updateServicoFilter();
  updateKpis();
  renderEstoquePanel();
  renderEnderecoStats();
  renderCargasPendentesTable();
  renderEnderecosTable();
  renderLotesTable();
  renderPereciveisTable();
  renderTop3Destinos();
  initTooltips();
}

function initTooltips() {
  if (typeof bootstrap === "undefined") return;
  const triggers = Array.from(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
  triggers.forEach((el) => bootstrap.Tooltip.getOrCreateInstance(el));
}

function clearAll() {
  state.fileName = null;
  state.delimiter = ";";
  state.hasHeader = true;
  state.colId = 0;
  state.colDest = 2;
  state.colLocal = 5;
  state.colPeso = 4;
  state.colPecas = 3;
  state.colData = 8;
  state.colServico = 9;
  state.colShc = 6;
  state.colPri = null;
  state.colSla = null;
  state.colPosse = null;
  state.colEmissao = null;
  state.rawRows = [];
  state.rows = [];

  const file = $("fileCargas");
  if (file) file.value = "";

  setError("");
  setActionMsg("");
  setText("tableStamp", "0");
  setText("lotesStamp", "0");
  setText("perecivelStamp", "0");
  setText("enderecosStamp", "0 endereços");
  setText("tabBadgeCargas", "0");
  setText("tabBadgeEnderecos", "0");
  setText("tabBadgeLotes", "0");
  setText("tabBadgePerecivel", "0");
  setText("appStamp", "Aguardando CSV");
  const filterServico = $("filterServico");
  if (filterServico) filterServico.value = "TODOS";

  renderAll();
}

async function loadCsvFromFile(file) {
  setError("");
  setActionMsg("");

  const buf = await readFileAsBuffer(file);
  const preferred = "utf-8";
  const alternate = "iso-8859-1";
  const textPref = decodeBuffer(buf, preferred);
  const textAlt = decodeBuffer(buf, alternate);
  const scorePref = mojibakeScore(textPref);
  const scoreAlt = mojibakeScore(textAlt);
  const picked =
    scoreAlt < scorePref && (scorePref - scoreAlt >= 2 || scoreAlt === 0) ? { text: textAlt, enc: alternate } : { text: textPref, enc: preferred };
  const text = picked.text;

  if (picked.enc !== preferred) setActionMsg(`Encoding ajustado automaticamente para ${picked.enc.toUpperCase()}.`);

  const firstLine = (text || "").split(/\r?\n/)[0] || "";
  state.delimiter = detectDelimiter(firstLine);
  state.rawRows = parseCsv(text, state.delimiter);
  state.fileName = file.name || "cargas.csv";

  state.hasHeader = guessHasHeader(state.rawRows);
  detectColumnsFromHeader();

  if (state.rawRows.length <= 1 && state.hasHeader) {
    setError("CSV parece vazio (ou só cabeçalho).");
  }

  mapRowsFromRaw();
  renderAll();
}

function enableModalStacking() {
  if (typeof bootstrap === "undefined") return;

  document.addEventListener("show.bs.modal", (event) => {
    const modal = event.target;
    if (!(modal instanceof HTMLElement)) return;
    const openModals = document.querySelectorAll(".modal.show").length;
    const zIndex = 1055 + openModals * 10;
    modal.style.zIndex = String(zIndex);

    window.setTimeout(() => {
      const backdrops = document.querySelectorAll(".modal-backdrop:not(.modal-stack)");
      const backdrop = backdrops[backdrops.length - 1];
      if (backdrop instanceof HTMLElement) {
        backdrop.classList.add("modal-stack");
        backdrop.style.zIndex = String(zIndex - 1);
      }
    }, 0);
  });

  document.addEventListener("hidden.bs.modal", () => {
    const openModals = document.querySelectorAll(".modal.show");
    if (openModals.length > 0) {
      document.body.classList.add("modal-open");
    } else {
      // Garantir limpeza completa quando todos os modais forem fechados
      document.body.classList.remove("modal-open");
      // Remover todos os backdrops órfãos
      const orphanBackdrops = document.querySelectorAll(".modal-backdrop");
      orphanBackdrops.forEach(b => b.remove());
      // Remover padding do body
      document.body.style.paddingRight = "";
    }
  });

  // Adicionar tratamento para cliques no backdrop
  document.addEventListener("click", (event) => {
    if (event.target.classList.contains("modal-backdrop")) {
      const modal = event.target.previousElementSibling;
      if (modal && modal.classList.contains("modal")) {
        const instance = bootstrap.Modal.getInstance(modal);
        if (instance) instance.hide();
      }
    }
  });
}


function pad2(n) {
  return String(n).padStart(2, "0");
}

function isPracaAddress(addr) {
  const m = (addr || "").match(/^([ABCD])(\d{2})([A-Z])$/);
  if (!m) return false;
  const bolsao = m[1];
  const num = parseInt(m[2], 10);
  const col = m[3];
  const code = col.charCodeAt(0);

  if (bolsao === "A" || bolsao === "B") return num >= 1 && num <= 22 && code >= 65 && code <= 72; // A-H
  if (bolsao === "C") return num >= 1 && num <= 3 && code >= 65 && code <= 69; // A-E
  if (bolsao === "D") return num >= 1 && num <= 12 && code >= 65 && code <= 74; // A-J
  return false;
}

function buildEnderecoSummaryMap() {
  const rows = collectPrintableRows();
  const byAddr = new Map();
  for (const r of rows) {
    const addr = normalizeEndereco(r.localOriginal || "");
    if (!REGEX_ENDERECO.test(addr)) continue;
    if (!isPracaAddress(addr)) continue;
    const list = byAddr.get(addr) || [];
    list.push(r);
    byAddr.set(addr, list);
  }

  const out = new Map();
  for (const [addr, rows] of byAddr.entries()) {
    const a = analyzeEnderecoGroup(rows);
    out.set(addr, {
      endereco: addr,
      destino: a.destLabel || "-",
      data: a.dataLabel || "-",
      mixed: a.anyMixed,
      destMixed: a.destMixed,
      dataMixed: a.dataMixed,
      servMixed: a.servMixed,
      shcTokens: a.shcTokens,
      count: rows.length,
      kg: a.kg,
    });
  }
  return out;
}

function buildCheckPracaHtml() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const stamp = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`;

  const map = buildEnderecoSummaryMap();
  const fileLabel = state.fileName || "-";

  const buildTable = (bolsao, maxNum, lastCol) => {
    const allCols = [];
    for (let c = 65; c <= lastCol.charCodeAt(0); c++) allCols.push(String.fromCharCode(c));
    const cols = bolsao === "A" ? [...allCols].reverse() : allCols;
    const fullColspan = cols.length;
    const rowDefs = buildRowDefinitions(bolsao, cols, maxNum);

    let html = `<section class="section" data-bolsao="${bolsao}"><div class="section-title">Bolsão ${bolsao}</div>`;
    html += `<table class="grid"><thead><tr>`;
    for (const col of cols) html += `<th>${col}</th>`;
    html += `</tr></thead><tbody>`;

    for (const row of rowDefs) {
      html += `<tr>`;
      const allowedLetters = row.allowedLetters || new Set(cols);
      for (const col of cols) {
        if (!allowedLetters.has(col)) {
          html += `<td class="cell cell-empty"></td>`;
          continue;
        }
        const addr = `${bolsao}${pad2(row.num)}${col}`;
        const item = map.get(addr) || null;
        let cls = "cell";
        if (item?.destMixed) cls += " cell-dest-mixed";
        else if (item?.dataMixed) cls += " cell-date-mixed";
        else if (item?.servMixed) cls += " cell-serv-mixed";
        else if (item?.mixed && (!item?.shcTokens || item.shcTokens.size === 0)) cls += " cell-mixed";
        const destino = (item?.destino ?? "").toString().trim();
        const data = (item?.data ?? "").toString().trim();
        const count = item?.count ?? 0;
        html += `<td class="${cls}">`;
        const hasDestino = destino && destino !== "-";
        const line1 = hasDestino ? `${addr} - ${destino}` : addr;
        html += `<div class="addr">${escapeHtml(line1)}</div>`;
        if (count > 0) {
          let metaHtml = "";
          if (item?.dataMixed && item?.destMixed) {
            metaHtml = `<div class="meta">DATA MIST.<br>DESTINO MIST.`;
            if (item?.servMixed) metaHtml += `<br>SERV MIST.`;
            metaHtml += ` <span class="x">x${count}</span></div>`;
          } else if (item?.dataMixed && item?.servMixed) {
            metaHtml = `<div class="meta">DATA MIST.<br>SERV MIST. <span class="x">x${count}</span></div>`;
          } else if (item?.destMixed && item?.servMixed) {
            metaHtml = `<div class="meta">DESTINO MIST.<br>SERV MIST. <span class="x">x${count}</span></div>`;
          } else {
            // Endereço OK - mostrar data e SHCs (se houver)
            let metaText = data || "-";
            if (item?.dataMixed) metaText = "DATA MIST.";
            else if (item?.destMixed) metaText = "DESTINO MIST.";
            else if (item?.servMixed) metaText = "SERV MIST.";
            
            // Para endereços OK (sem mistura), adicionar SHCs abaixo da data
            let shcText = "";
            if (!item?.dataMixed && !item?.destMixed && !item?.servMixed && item?.shcTokens && item.shcTokens.size > 0) {
              const shcArray = Array.from(item.shcTokens).sort();
              shcText = shcArray.join(", ");
            }
            
            if (shcText) {
              metaHtml = `<div class="meta">${escapeHtml(metaText)}<br>${escapeHtml(shcText)} <span class="x">x${count}</span></div>`;
            } else {
              metaHtml = `<div class="meta">${escapeHtml(metaText)} <span class="x">x${count}</span></div>`;
            }
          }
          html += metaHtml;
        } else html += `<div class="meta">&nbsp;</div>`;
        html += `</td>`;
      }
      html += `</tr>`;
      if (row.spacerAfter) html += `<tr class="spacer-row"><td colspan="${fullColspan}"></td></tr>`;
    }

    html += `</tbody></table></section>`;
    return html;
  };

  const body = `
    <div class="page page-a">${buildTable("A", 22, "H")}</div>
    <div class="page page-b">${buildTable("B", 22, "H")}</div>
    <div class="page page-cd">${buildTable("C", 3, "E")}${buildTable("D", 12, "J")}</div>
  `;

  return `<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Check de Praça</title>
  <style>
    :root { --fg:#111; --muted:#666; --border:#222; --mixed:#e6e6e6; }
    body { font-family: system-ui, -apple-system, "Segoe UI", Arial, sans-serif; color:var(--fg); margin: 0; }
    .toolbar { position: sticky; top: 0; background: #fff; border-bottom: 1px solid var(--border); padding: 10px 12px; display:flex; gap:12px; align-items:baseline; flex-wrap:wrap; }
    .title { font-weight: 800; letter-spacing: 0.2px; }
    .meta { color: var(--muted); font-size: 12px; }
    .btn { font: inherit; font-weight: 800; border: 1px solid var(--border); background:#fff; padding: 6px 10px; cursor:pointer; }
    .wrap { padding: 12px; }
    .section { margin-bottom: 16px; break-inside: avoid; }
    .section-title { font-weight: 800; text-transform: uppercase; letter-spacing: 0.3px; font-size: 12px; margin: 6px 0; }
    table.grid { width: 100%; border-collapse: collapse; table-layout: fixed; }
    table.grid th, table.grid td { border: 1px solid var(--border); padding: 4px; vertical-align: top; }
    table.grid thead th { background: #f3f3f3; font-size: 11px; }
    th.rowhead { width: 34px; text-align: center; font-size: 11px; background: #f3f3f3; }
    td.cell { height: 48px; }
    td.cell-mixed { background: var(--mixed); }
    td.cell-date-mixed { background: #e0e0e0; }
    td.cell-dest-mixed { background: #d5d5d5; }
    td.cell-serv-mixed { background: #e8e8e8; }
    .addr { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Courier New", monospace; font-weight: 800; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; line-height: 1.1; }
    .meta { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Courier New", monospace; font-size: 11px; line-height: 1.1; }
    .x { color: #000; font-weight: 800; }
    .cell-empty {
      background: transparent;
    }
    tr.spacer-row td {
      padding: 0;
      border: none;
      height: 4mm;
      background: transparent;
    }
    @media print {
      /* Força impressão de cores de fundo */
      * {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
      }

      .toolbar { display: none; }
      .wrap { padding: 8px 0 0 8px; }
      .page { break-after: page; break-inside: avoid; page-break-inside: avoid; }
      .page.page-cd { break-after: auto; }
      .page:last-child { break-after: auto; }
      .section { margin-bottom: 6mm; }

      /* Cabeça/espacamento compactos para caber em 1 folha */
      table.grid { break-inside: avoid; page-break-inside: avoid; }
      table.grid th, table.grid td { padding: 2px; }
      .section-title { margin: 2mm 0 1mm 0; }

      /* Bolsões A e B: 22 linhas precisam caber em 1 página */
      .page.page-a td.cell, .page.page-b td.cell { height: 8.8mm; max-height: 8.8mm; overflow: hidden; }
      .page.page-a .addr, .page.page-b .addr { font-size: 9.5px; }
      .page.page-a .meta, .page.page-b .meta { font-size: 8.5px; }
      .page.page-a th.rowhead, .page.page-b th.rowhead { font-size: 9px; }

      /* C e D na mesma página (C é pequeno) */
      .page.page-cd td.cell { height: 8.8mm; max-height: 8.8mm; overflow: hidden; }
      .page.page-cd .addr { font-size: 9.5px; }
      .page.page-cd .meta { font-size: 8.5px; }

      /* Mantém cores de fundo na impressão */
      table.grid thead th { background: #f3f3f3 !important; }
      th.rowhead { background: #f3f3f3 !important; }
      td.cell-mixed { background: #e6e6e6 !important; }
      td.cell-date-mixed { background: #e0e0e0 !important; }
      td.cell-dest-mixed { background: #d5d5d5 !important; }
      td.cell-serv-mixed { background: #e8e8e8 !important; }

      /* Espaçamento entre linhas */
      tr.spacer-row td {
        padding: 0 !important;
        border: none !important;
        height: 4mm !important;
        background: transparent !important;
      }

      @page { margin: 10mm; }
    }
  </style>
</head>
<body>
  <div class="toolbar">
    <div class="title">Check de Praça</div>
    <div class="meta">Arquivo: ${escapeHtml(fileLabel)} | Gerado: ${escapeHtml(stamp)}</div>
    <button class="btn" onclick="window.print()">Imprimir</button>
  </div>
  <div class="wrap">
    ${body}
  </div>
</body>
</html>`;
}

function buildAuditoriaHtml() {
  const stats = getEnderecoStats();
  
  const renderTable = (title, cargas, colspan = 6) => {
    let html = `<h3 style="margin-top: 20px; margin-bottom: 10px; border-bottom: 2px solid #ccc; padding-bottom: 8px;">${title} (${cargas.length} cargas, ${cargas.reduce((sum, c) => sum + c.pecas, 0)} volumes)</h3>`;
    html += `<table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">`;
    html += `<thead><tr style="background: #f3f3f3;">
      <th style="border: 1px solid #ccc; padding: 8px; text-align: left;">AWB</th>
      <th style="border: 1px solid #ccc; padding: 8px; text-align: center;">Volume Registrado</th>
      <th style="border: 1px solid #ccc; padding: 8px; text-align: left;">Destino</th>
      <th style="border: 1px solid #ccc; padding: 8px; text-align: left;">Endereço</th>
      <th style="border: 1px solid #ccc; padding: 8px; text-align: left;">Localização (Raw)</th>
    </tr></thead><tbody>`;
    
    for (const c of cargas) {
      html += `<tr style="border: 1px solid #ddd;">
        <td style="border: 1px solid #ddd; padding: 8px; font-family: monospace;">${escapeHtml(c.id)}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: center; font-family: monospace;">${c.pecas}</td>
        <td style="border: 1px solid #ddd; padding: 8px; font-family: monospace;">${escapeHtml(c.destino)}</td>
        <td style="border: 1px solid #ddd; padding: 8px; font-family: monospace;">${escapeHtml(c.endereco)}</td>
        <td style="border: 1px solid #ddd; padding: 8px; font-family: monospace; font-size: 11px;">${escapeHtml(c.localizacao)}</td>
      </tr>`;
    }
    html += `</tbody></table>`;
    return html;
  };
  
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const stamp = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
  const fileLabel = state.fileName || "-";
  
  return `<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Auditoria - Endereçamento</title>
  <style>
    body { font-family: system-ui, -apple-system, "Segoe UI", Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
    .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; }
    .header { border-bottom: 2px solid #333; padding-bottom: 15px; margin-bottom: 20px; }
    .title { font-size: 24px; font-weight: 800; margin: 0 0 10px 0; }
    .meta { color: #666; font-size: 14px; }
    .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 30px; }
    .stat-box { background: #f9f9f9; border: 1px solid #ddd; border-radius: 4px; padding: 15px; }
    .stat-label { font-size: 12px; color: #666; text-transform: uppercase; margin-bottom: 5px; }
    .stat-value { font-size: 32px; font-weight: 800; color: #333; }
    h3 { color: #333; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="title">Auditoria de Endereçamento</div>
      <div class="meta">Arquivo: ${escapeHtml(fileLabel)} | Gerado: ${escapeHtml(stamp)}</div>
    </div>
    
    <div class="stats">
      <div class="stat-box">
        <div class="stat-label">Volumes Endereçados</div>
        <div class="stat-value">${stats.enderecoados}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Volumes na Triagem</div>
        <div class="stat-value">${stats.triagem}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Peso Total (kg)</div>
        <div class="stat-value">${formatNumber(stats.pesoEnderecoado + stats.pesoTriagem)}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Peso Endereçado (kg)</div>
        <div class="stat-value">${formatNumber(stats.pesoEnderecoado)}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Peso na Triagem (kg)</div>
        <div class="stat-value">${formatNumber(stats.pesoTriagem)}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Taxa de Endereçamento</div>
        <div class="stat-value">${stats.total ? Math.round((stats.enderecoados / stats.total) * 100) : 0}%</div>
      </div>
    </div>
    
    ${renderTable("Cargas Endereçadas", stats.cargasEnderecoadas)}
    ${renderTable("Cargas na Triagem", stats.cargasTriagem)}
  </div>
</body>
</html>`;
}

function openAuditoria() {
  const html = buildAuditoriaHtml();
  const blob = new Blob([html], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  window.open(url, "_blank");
}

function buildOtimizacaoHtml() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const stamp = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
  const fileLabel = state.fileName || "-";

  const addrToRows = new Map();
  for (const r of state.rows) {
    const addr = normalizeEndereco(r.localOriginal || "");
    if (!REGEX_ENDERECO.test(addr)) continue;
    if (!isPracaAddress(addr)) continue;
    const list = addrToRows.get(addr) || [];
    list.push(r);
    addrToRows.set(addr, list);
  }

  const groups = new Map(); // key dest|data|servico -> group
  for (const [addr, rows] of addrToRows.entries()) {
    if (!rows.length) continue;
    const a = analyzeEnderecoGroup(rows);
    if (a.destMixed || a.dataMixed || a.servMixed) continue;
    if (!a.destLabel || !a.dataLabel || !a.servLabel) continue;
    if (a.destLabel === "MISTO" || a.dataLabel === "MISTO" || a.servLabel === "MISTO") continue;

    const servKey = a.servLabel;
    const key = `${a.destLabel}|${a.dataLabel}|${servKey}`;
    let g = groups.get(key);
    if (!g) {
      g = {
        key,
        destino: a.destLabel,
        data: a.dataLabel,
        servKey,
        enderecos: [],
        shcTokens: new Set(),
        kg: 0,
        countRows: 0,
      };
      groups.set(key, g);
    }

    const servLabel = servicoLabelFromKey(servKey) || servKey || "-";
    g.enderecos.push({
      endereco: addr,
      servico: servLabel,
      shc: shcLabelFromSet(a.shcTokens),
      shcConflict: a.shcConflict,
      kg: a.kg,
      count: rows.length,
    });
    for (const t of a.shcTokens || []) g.shcTokens.add(t);
    g.kg += a.kg;
    g.countRows += rows.length;
  }

  const list = Array.from(groups.values())
    .filter((g) => g.enderecos.length > 1)
    .map((g) => {
      g.enderecos.sort((a, b) => a.endereco.localeCompare(b.endereco));
      g.anyShcConflict = analyzeShcConflict(g.shcTokens);
      g.shc = shcLabelFromSet(g.shcTokens);
      g.serv = servicoLabelFromKey(g.servKey) || g.servKey || "-";
      
      // Calcular economia potencial
      const maxKgEndereço = Math.max(...g.enderecos.map(e => e.kg));
      g.economiaKg = g.kg - maxKgEndereço;
      g.economiaPercent = Math.round((g.economiaKg / g.kg) * 100);
      
      return g;
    })
    .sort((a, b) => {
      // Ordenação: 1º por risco SHC, 2º por economia potencial (kg), 3º por qtd endereços
      if (a.anyShcConflict !== b.anyShcConflict) return a.anyShcConflict ? -1 : 1;
      if (b.economiaKg !== a.economiaKg) return b.economiaKg - a.economiaKg;
      if (b.enderecos.length !== a.enderecos.length) return b.enderecos.length - a.enderecos.length;
      const d = a.data.localeCompare(b.data);
      if (d) return d;
      return a.destino.localeCompare(b.destino);
    });

  // Estatísticas gerais
  const stats = {
    totalGrupos: list.length,
    totalEnderecos: list.reduce((sum, g) => sum + g.enderecos.length, 0),
    pesoTotal: list.reduce((sum, g) => sum + g.kg, 0),
    economiaTotal: list.reduce((sum, g) => sum + g.economiaKg, 0),
    gruposComConflito: list.filter(g => g.anyShcConflict).length,
  };

  const emptyMsg =
    state.rows.length === 0
      ? "Carregue um CSV para gerar o relatório."
      : list.length === 0
        ? "Nenhum caso de mesmo destino, data e serviço em mais de um endereço (apenas bolsões A-D)."
        : "";

  const sections = list
    .map((g, idx) => {
      const endComMaiorKg = g.enderecos.reduce((max, e) => e.kg > max.kg ? e : max);

      let html = "";
      for (let i = 0; i < g.enderecos.length; i++) {
        const e = g.enderecos[i];
        const isFirst = i === 0;
        
        html += `<tr style="border-bottom: 1px solid #ddd;">`;
        
        if (isFirst) {
          html += `<td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-weight: 600; font-size: 13px;" rowspan="${g.enderecos.length}">${escapeHtml(g.destino)}</td>`;
          html += `<td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;" rowspan="${g.enderecos.length}">${escapeHtml(g.data)}</td>`;
          html += `<td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;" rowspan="${g.enderecos.length}">${escapeHtml(g.serv)}</td>`;
          html += `<td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; text-align: right; font-size: 13px; font-weight: 600;" rowspan="${g.enderecos.length}">${g.economiaPercent}% / ${formatNumber(g.economiaKg)}kg</td>`;
        }
        
        html += `<td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;">${escapeHtml(e.endereco)}${isFirst && e.endereco === endComMaiorKg.endereco ? ' *' : ''}</td>`;
        html += `<td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; text-align: right; font-size: 13px;">${e.count}</td>`;
        html += `<td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; text-align: right; font-size: 13px;">${formatNumber(e.kg)}</td>`;
        html += `<td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;">${escapeHtml(e.shc)}</td>`;
        html += `</tr>`;
      }
      
      return html;
    })
    .join("\n");

  return `<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Otimização</title>
  <style>
    body { font-family: system-ui, -apple-system, "Segoe UI", Arial, sans-serif; margin: 0; padding: 20px; background: #fff; font-size: 14px; }
    .container { max-width: 1400px; margin: 0 auto; }
    .header { border-bottom: 2px solid #333; padding-bottom: 15px; margin-bottom: 15px; }
    .title { font-size: 20px; font-weight: 800; margin: 0 0 8px 0; }
    .meta { color: #666; font-size: 13px; }
    .info { padding: 12px 0; font-size: 13px; margin-bottom: 15px; }
    table { width: 100%; border-collapse: collapse; }
    thead tr { }
    th { border: 1px solid #333; padding: 6px; text-align: left; font-size: 13px; font-weight: 700; }
    td { border: 1px solid #333; padding: 6px; font-family: monospace; font-size: 13px; }
    td.text-right { text-align: right; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="title">Otimização</div>
      <div class="meta">Arquivo: ${escapeHtml(fileLabel)} | Gerado: ${escapeHtml(stamp)}</div>
    </div>
    
    ${emptyMsg ? `<div class="info">${escapeHtml(emptyMsg)}</div>` : ""}
    
    ${stats.totalGrupos > 0 ? `
    <div class="info">${stats.totalGrupos} grupos • ${stats.totalEnderecos} endereços • ${formatNumber(stats.pesoTotal)}kg • ${formatNumber(stats.economiaTotal)}kg economia</div>
    <table>
      <thead>
        <tr>
          <th>Destino</th>
          <th>Data</th>
          <th>Serviço</th>
          <th>Economia</th>
          <th>Endereço</th>
          <th class="text-right">Volume Registrado</th>
          <th class="text-right">Peso (kg)</th>
          <th>SHC</th>
        </tr>
      </thead>
      <tbody>
        ${sections}
      </tbody>
    </table>
    ` : ""}
  </div>
</body>
</html>`;
}

function openCheckPraca() {
  const html = buildCheckPracaHtml();
  const win = window.open("", "_blank");
  if (!win) {
    setActionMsg("Pop-up bloqueado: permita abrir novas abas para gerar o Check de Praça.");
    return;
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
  win.focus();
}

function openOtimizacao() {
  const html = buildOtimizacaoHtml();
  const win = window.open("", "_blank");
  if (!win) {
    setActionMsg("Pop-up bloqueado: permita abrir novas abas para gerar a Otimização.");
    return;
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
  win.focus();
}

function listBolsaoAddresses(bolsao) {
  const b = (bolsao || "").toString().toUpperCase();
  const spec =
    b === "A" || b === "B"
      ? { maxNum: 22, lastCol: "H" }
      : b === "C"
        ? { maxNum: 3, lastCol: "E" }
        : b === "D"
          ? { maxNum: 12, lastCol: "J" }
          : null;
  if (!spec) return [];

  const cols = [];
  for (let c = 65; c <= spec.lastCol.charCodeAt(0); c++) cols.push(String.fromCharCode(c));

  const out = [];
  for (let n = 1; n <= spec.maxNum; n++) {
    for (const col of cols) out.push(`${b}${pad2(n)}${col}`);
  }
  return out;
}

function renderEstoquePanel() {
  setText("estoquePanelTitle", "Estoque");
  setText("estoquePanelSub", `${state.fileName || "-"} | ${state.rows.length} cargas`);

  const panelBody = $("estoquePanelBody");
  if (!panelBody) return;
  panelBody.innerHTML = "";

  if (!state.rows.length) {
    panelBody.innerHTML = '<div class="ds-estoque-empty text-muted">Carregue um CSV para ver o estoque.</div>';
    return;
  }

  const addrToRows = new Map();
  for (const r of state.rows) {
    const addr = normalizeEndereco(r.localOriginal || "");
    if (!REGEX_ENDERECO.test(addr)) continue;
    if (!isPracaAddress(addr)) continue;
    const list = addrToRows.get(addr) || [];
    list.push(r);
    addrToRows.set(addr, list);
  }

  const sum = { totalPos: 0, occupiedPos: 0, freePos: 0, kg: 0, cargas: 0, mixed: 0, shc: 0, shcConflict: 0 };

  const formatPct = (occ, total) => {
    if (!total) return "0%";
    const pct = Math.round((occ / total) * 100);
    return `${pct}%`;
  };

  const formatTop = (m) => {
    const arr = Array.from(m.entries()).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
    if (!arr.length) return "-";
    return arr.map(([k, v]) => `${k}(${v})`).join(", ");
  };

  const bolsaoStats = [];
  for (const bolsao of ["A", "B", "C", "D"]) {
    const allAddrs = listBolsaoAddresses(bolsao);
    const totalPos = allAddrs.length;
    const occupied = allAddrs.filter((a) => addrToRows.has(a));
    const occupiedPos = occupied.length;
    const freePos = totalPos - occupiedPos;

    const destinos = new Map();
    const destinoAddresses = new Map(); // Mapear destino → endereços
    const servicos = new Map();
    let kg = 0;
    let cargas = 0;
    let mixed = 0;
    let shc = 0;
    let shcConflict = 0;

    for (const addr of occupied) {
      const rows = addrToRows.get(addr) || [];
      if (!rows.length) continue;
      const a = analyzeEnderecoGroup(rows);
      cargas += rows.length;
      kg += a.kg;
      if (a.destLabel) {
        // Contar volumes (pecas) em vez de cargas
        const volumes = rows.reduce((sum, r) => sum + (r.pecas || 1), 0);
        destinos.set(a.destLabel, (destinos.get(a.destLabel) || 0) + volumes);
        // Guardar endereços por destino
        if (!destinoAddresses.has(a.destLabel)) {
          destinoAddresses.set(a.destLabel, new Set());
        }
        destinoAddresses.get(a.destLabel).add(addr);
      }
      if (a.servLabel) servicos.set(a.servLabel, (servicos.get(a.servLabel) || 0) + 1);
      if (a.anyMixed) mixed++;
      if (a.shcTokens && a.shcTokens.size) shc++;
      if (a.shcConflict) shcConflict++;
    }

    const pctOcc = Math.round((occupiedPos / totalPos) * 100);
    const pctClass = shcConflict > 0 ? "risco" : mixed > 0 ? "alerta" : "";

    bolsaoStats.push({
      bolsao,
      destinos: formatTop(destinos),
      destinoAddresses,
      servicos: formatTop(new Map(Array.from(servicos.entries()).map(([k, v]) => [servicoLabelFromKey(k) || k, v]))),
      totalPos,
      occupiedPos,
      freePos,
      pct: formatPct(occupiedPos, totalPos),
      pctNum: pctOcc,
      pctClass,
      kg,
      cargas,
      mixed,
      shc,
      shcConflict,
    });

    sum.totalPos += totalPos;
    sum.occupiedPos += occupiedPos;
    sum.freePos += freePos;
    sum.kg += kg;
    sum.cargas += cargas;
    sum.mixed += mixed;
    sum.shc += shc;
    sum.shcConflict += shcConflict;
  }

  const cargasPendentes = state.rows.filter((r) => !isEnderecado(r)).length;
  const cargasAlocadasPraca = state.rows.filter((r) => {
    const addr = normalizeEndereco(r.localOriginal || "");
    return REGEX_ENDERECO.test(addr) && isPracaAddress(addr);
  }).length;
  const cargasAlocadasForaPraca = state.rows.filter((r) => {
    const addr = normalizeEndereco(r.localOriginal || "");
    return REGEX_ENDERECO.test(addr) && !isPracaAddress(addr);
  }).length;

  const frag = document.createDocumentFragment();

  const summary = document.createElement("div");
  summary.className = "ds-estoque-summary-slim";
  summary.innerHTML = `
    <div class="ds-estoque-metric">
      <div class="ds-estoque-metric-label">Capacidade A-D</div>
      <div class="ds-estoque-metric-value">${sum.occupiedPos}/${sum.totalPos}</div>
      <div class="ds-estoque-metric-meta">${formatPct(sum.occupiedPos, sum.totalPos)} ocupados · Livre ${sum.freePos}</div>
    </div>
    <div class="ds-estoque-metric">
      <div class="ds-estoque-metric-label">Pendentes</div>
      <div class="ds-estoque-metric-value">${cargasPendentes}</div>
      <div class="ds-estoque-metric-meta">Em A-D ${cargasAlocadasPraca} · Fora ${cargasAlocadasForaPraca}</div>
    </div>
    <div class="ds-estoque-metric">
      <div class="ds-estoque-metric-label">Alertas</div>
      <div class="ds-estoque-metric-value">${sum.mixed}</div>
      <div class="ds-estoque-metric-meta">SHC ${sum.shc} · Risco ${sum.shcConflict}</div>
    </div>
  `;
  frag.appendChild(summary);

  const bolsaoList = document.createElement("div");
  bolsaoList.className = "ds-estoque-bolsao-list";
  for (const b of bolsaoStats) {
    const card = document.createElement("article");
    card.className = "ds-estoque-bolsao-card";
    const servicos = escapeHtml(b.servicos || "-");
    
    card.innerHTML = `
      <header class="ds-estoque-bolsao-head">
        <span class="ds-estoque-bolsao-title">Bolso ${escapeHtml(b.bolsao)}</span>
        <span class="ds-estoque-bolsao-count mono">${escapeHtml(`${b.occupiedPos}/${b.totalPos}`)}</span>
      </header>
    `;
    
    // Criar barra visual de ocupação
    const barraDom = document.createElement("div");
    barraDom.className = "ds-estoque-barra-ocupacao";
    const fill = document.createElement("div");
    fill.className = `ds-estoque-barra-fill ${b.pctClass}`;
    fill.style.width = `${b.pctNum}%`;
    fill.textContent = b.pctNum > 15 ? `${b.pctNum}%` : "";
    barraDom.appendChild(fill);
    card.appendChild(barraDom);
    
    // Corpo com informações
    const body = document.createElement("div");
    body.className = "ds-estoque-bolsao-body";
    
    // Criar linha de destinos com tooltips
    const destinosLine = document.createElement("div");
    destinosLine.className = "ds-estoque-bolsao-line";
    const destLabel = document.createElement("span");
    destLabel.textContent = "Destinos";
    destinosLine.appendChild(destLabel);
    
    const destContainer = document.createElement("span");
    destContainer.className = "mono";
    
    if (b.destinos === "-") {
      destContainer.textContent = "-";
    } else {
      // Parsear destinos e adicionar tooltips
      const destArray = b.destinos.split(", ");
      destArray.forEach((destItem, idx) => {
        // Extrair código e quantidade: "RBR(13)" -> "RBR", "13"
        const match = destItem.match(/^(.+?)\((\d+)\)$/);
        const destSpan = document.createElement("span");
        destSpan.className = "ds-estoque-destino";
        destSpan.textContent = destItem;
        
        if (match) {
          const destCode = match[1];
          const destAddrs = b.destinoAddresses.get(destCode);
          
          if (destAddrs && destAddrs.size > 0) {
            destSpan.classList.add("ds-estoque-destino-tooltip");
            destSpan.setAttribute("data-bs-toggle", "tooltip");
            destSpan.setAttribute("data-bs-placement", "top");
            destSpan.setAttribute("title", Array.from(destAddrs).sort().join(", "));
          }
        }
        
        destContainer.appendChild(destSpan);
        
        // Adicionar vírgula após cada destino, exceto o último
        if (idx < destArray.length - 1) {
          const sep = document.createTextNode(", ");
          destContainer.appendChild(sep);
        }
      });
    }
    
    destinosLine.appendChild(destContainer);
    body.appendChild(destinosLine);
    
    // Adicionar outras linhas
    const otherLines = document.createElement("div");
    otherLines.innerHTML = `
      <div class="ds-estoque-bolsao-line">
        <span>Serviços</span>
        <span class="mono">${servicos}</span>
      </div>
      <div class="ds-estoque-bolsao-line">
        <span>Cargas</span>
        <span class="mono">${escapeHtml(String(b.cargas))} · ${escapeHtml(formatNumber(b.kg))} kg</span>
      </div>
      <div class="ds-estoque-bolsao-line">
        <span>Mistos / SHC / Risco</span>
        <span class="mono">${escapeHtml(String(b.mixed))} · ${escapeHtml(String(b.shc))} · ${escapeHtml(String(b.shcConflict))}</span>
      </div>
    `;
    body.appendChild(otherLines);
    card.appendChild(body);
    
    // Meta
    const meta = document.createElement("div");
    meta.className = "ds-estoque-bolsao-meta";
    meta.textContent = `${b.pct} ocupados · ${b.freePos} livres`;
    card.appendChild(meta);
    
    bolsaoList.appendChild(card);
  }
  frag.appendChild(bolsaoList);

  panelBody.appendChild(frag);
  
  // Inicializar tooltips Bootstrap
  document.querySelectorAll(".ds-estoque-destino-tooltip").forEach(el => {
    new bootstrap.Tooltip(el);
  });
}

// ========== VOOS ==========
function parseVooRow(cols) {
  // Modal, Empresa, N.º de controle, Origem / Destino, Partida, Capacidade, Disponibilidade, Quantidade Pax, Equipamento, Tipo de equipamento
  // 0,      1,        2,                3,                4,        5,          6,               7,               8,          9
  if (cols.length < 10) return null;

  const controle = (cols[2] || "").trim();
  const origemDestino = (cols[3] || "").trim();
  const partida = (cols[4] || "").trim();
  const capacidade = (cols[5] || "").trim();
  const disponibilidade = (cols[6] || "").trim();
  const pax = (cols[7] || "").trim();
  const tipoEquipamento = (cols[9] || "").trim();

  if (!controle || !origemDestino) return null;

  // Extrair destino (pega tudo após " / ")
  const destino = origemDestino.split("/").pop()?.trim() || "";

  // Extrair data, hora corte, hora voo de partida (formato: "22/01/2026 06:00 07:50")
  const partidaParts = partida.split(/\s+/);
  const data = partidaParts[0] || "";
  const corte = partidaParts[1] || "";
  const voo = partidaParts[2] || "";

  // Extrair capacidade em kg
  const capKg = parsePtNumber(capacidade.split(/\s+/)[0] || "");

  // Extrair disponibilidade em kg
  const dispKg = parsePtNumber(disponibilidade.split(/\s+/)[0] || "");

  // Extrair pax
  const quantPax = parseIntSafe(pax);

  // Extrair tipo de equipamento (primeira palavra antes de "-")
  const equipCode = tipoEquipamento.split(/\s*-\s*/)[0]?.trim() || "";

  return {
    controle,
    destino,
    data,
    corte,
    voo,
    capKg,
    dispKg,
    pax: quantPax,
    equipamento: equipCode,
  };
}

async function loadVoosFromFile(file) {
  const buf = await readFileAsBuffer(file);
  const encoding = "UTF-8";
  const text = decodeBuffer(buf, encoding);
  
  // Detectar delimiter (TAB ou ;)
  const firstLine = text.split("\n")[0] || "";
  const delimiter = detectDelimiter(firstLine);
  
  const rows = parseCsv(text, delimiter);

  if (rows.length < 2) {
    throw new Error("Arquivo de voos vazio ou sem dados.");
  }

  stateVoos.fileName = file.name;
  stateVoos.rawRows = rows;
  stateVoos.voos = [];

  const startRow = 1; // Pular header
  for (let i = startRow; i < rows.length; i++) {
    const voo = parseVooRow(rows[i]);
    if (voo) stateVoos.voos.push(voo);
  }

  renderVoosTable();
  setText("programacaoMsg", `${stateVoos.voos.length} voos carregados com sucesso.`);
  setText("voosStamp", `${stateVoos.voos.length} voos`);
  setText("programacaoError", "");
}

function renderVoosTable() {
  const tbody = $("tblVoos")?.querySelector("tbody");
  if (!tbody) return;

  // Filtrar voos
  let voosExibir = stateVoos.voos;

  // Filtro de data
  const hoje = new Date();
  const amanha = new Date(hoje);
  amanha.setDate(amanha.getDate() + 1);

  const formatDataPt = (d) => {
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  };

  const hojeStr = formatDataPt(hoje);
  const amanhaStr = formatDataPt(amanha);

  if (stateVoos.filtroData === "hoje") {
    voosExibir = voosExibir.filter((v) => v.data === hojeStr);
  } else if (stateVoos.filtroData === "amanha") {
    voosExibir = voosExibir.filter((v) => v.data === amanhaStr);
  }

  // Filtro de terminal
  if (stateVoos.filtroTerminal === "T2") {
    // T2: horário de corte entre 16h e 22:50h
    voosExibir = voosExibir.filter((v) => {
      const corte = v.corte; // formato "HH:MM"
      if (!corte || !corte.includes(":")) return false;
      const [hh, mm] = corte.split(":").map(Number);
      const minutes = hh * 60 + mm;
      const inicio16h = 16 * 60; // 960 minutos
      const fim22h50 = 22 * 60 + 50; // 1370 minutos
      return minutes >= inicio16h && minutes <= fim22h50;
    });
  }
  // T1 e T3 não fazem filtro por enquanto

  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();

  for (const voo of voosExibir) {
    const tr = document.createElement("tr");

    const tdControle = document.createElement("td");
    tdControle.className = "mono";
    tdControle.textContent = voo.controle;

    const tdDestino = document.createElement("td");
    tdDestino.className = "mono";
    tdDestino.textContent = voo.destino;

    const tdData = document.createElement("td");
    tdData.className = "mono";
    tdData.textContent = voo.data;

    const tdCorte = document.createElement("td");
    tdCorte.className = "mono";
    tdCorte.textContent = voo.corte;

    const tdVoo = document.createElement("td");
    tdVoo.className = "mono";
    tdVoo.textContent = voo.voo;

    const tdDisp = document.createElement("td");
    tdDisp.className = "mono text-end";
    tdDisp.textContent = formatNumber(voo.dispKg);

    const tdPax = document.createElement("td");
    tdPax.className = "mono text-end";
    tdPax.textContent = voo.pax;

    const tdEquip = document.createElement("td");
    tdEquip.className = "mono";
    tdEquip.textContent = voo.equipamento;

    const tdProgram = document.createElement("td");
    tdProgram.className = "text-center";
    const btnProgram = document.createElement("button");
    btnProgram.className = "btn btn-sm btn-outline-dark";
    btnProgram.textContent = "Programar";
    btnProgram.addEventListener("click", (e) => {
      e.preventDefault();
      showFiltroServico(voo);
    });
    tdProgram.appendChild(btnProgram);

    tr.appendChild(tdControle);
    tr.appendChild(tdDestino);
    tr.appendChild(tdData);
    tr.appendChild(tdCorte);
    tr.appendChild(tdVoo);
    tr.appendChild(tdDisp);
    tr.appendChild(tdPax);
    tr.appendChild(tdEquip);
    tr.appendChild(tdProgram);

    frag.appendChild(tr);
  }

  tbody.appendChild(frag);
  setText("voosStamp", `${voosExibir.length} voos`);
}

function getPrioridadeSla(slaPctRemaining) {
  if (slaPctRemaining === null) return 3; // sem SLA fica com prioridade baixa
  if (slaPctRemaining < 30) return 0; // menos de 30%: máxima prioridade
  if (slaPctRemaining < 50) return 1; // menos de 50%: média prioridade
  return 2; // mais de 50%: baixa prioridade
}

function getPrioridadeServico(servicoGroup) {
  const s = normalizeUpper(servicoGroup || "");
  if (s.includes("SAUDE")) return 4;
  if (s.includes("URGENTE")) return 3;
  if (s === "EXPRESSO") return 2;
  if (s.includes("ECO") || s.includes("MODA")) return 1;
  return 0; // outros
}

function temApenasELI(shcTokens) {
  if (!shcTokens || shcTokens.size === 0) return false;
  const tokens = Array.from(shcTokens || []);
  const eli = new Set(["ELI", "ELM"]);
  return tokens.every((t) => eli.has(t));
}

function buildProgramacaoVooHtml(voo, filtro = null) {
  // Coletar cargas com destino = voo.destino
  let cargasVoo = state.rows.filter((r) => normalizeUpper(r.destino) === normalizeUpper(voo.destino));
  
  // Filtrar por tipo de serviço se especificado
  if (filtro === "SAUDE") {
    cargasVoo = cargasVoo.filter(r => {
      const serv = normalizeUpper(r.servicoGroup || servicoGroupKey(r.servico || ""));
      return serv.includes("SAUDE");
    });
  } else if (filtro === "URGENTE_EXPRESSO") {
    cargasVoo = cargasVoo.filter(r => {
      const serv = normalizeUpper(r.servicoGroup || servicoGroupKey(r.servico || ""));
      return serv.includes("URGENTE") || serv === "EXPRESSO";
    });
  } else if (filtro === "ECONOMICO") {
    cargasVoo = cargasVoo.filter(r => {
      const serv = normalizeUpper(r.servicoGroup || servicoGroupKey(r.servico || ""));
      return serv.includes("ECO") || serv.includes("MODA") || serv.includes("ECONOMICO");
    });
  }

  // Agrupar por endereço e ordenar
  const porEndereco = new Map();
  for (const carga of cargasVoo) {
    const addr = normalizeEndereco(carga.localOriginal || "");
    if (!REGEX_ENDERECO.test(addr)) continue;
    if (!isPracaAddress(addr)) continue;

    const list = porEndereco.get(addr) || [];
    list.push(carga);
    porEndereco.set(addr, list);
  }

  // Criar lista de cargas com ordenação por endereço e dentro dele
  const todasCargas = [];
  const enderecoInteiro = new Set(); // Endereços que têm TODAS as cargas do estoque
  
  // Verificar quais endereços têm todas as cargas
  for (const [addr, cargas] of porEndereco.entries()) {
    const todasCargasDoEndereco = state.rows.filter((r) => {
      const rAddr = normalizeEndereco(r.localOriginal || "");
      return rAddr === addr && REGEX_ENDERECO.test(rAddr) && isPracaAddress(rAddr);
    });
    
    if (cargas.length === todasCargasDoEndereco.length) {
      enderecoInteiro.add(addr);
    }
  }

  for (const [addr, cargas] of porEndereco.entries()) {
    const cargasOrdenadas = cargas.slice().sort((a, b) => {
      const slaPrioA = getPrioridadeSla(a.slaPctRemaining);
      const slaPrioB = getPrioridadeSla(b.slaPctRemaining);
      if (slaPrioA !== slaPrioB) return slaPrioA - slaPrioB;

      const temPER_A = a.shcTokens && a.shcTokens.has("PER");
      const temPER_B = b.shcTokens && b.shcTokens.has("PER");
      const eliA = temApenasELI(a.shcTokens);
      const eliB = temApenasELI(b.shcTokens);
      const servA = normalizeUpper(a.servicoGroup || servicoGroupKey(a.servico || ""));
      const servB = normalizeUpper(b.servicoGroup || servicoGroupKey(b.servico || ""));

      // PER e ELI têm prioridade máxima
      const temPER_ELI_A = (temPER_A || eliA);
      const temPER_ELI_B = (temPER_B || eliB);
      if (temPER_ELI_A && !temPER_ELI_B) return -1;
      if (!temPER_ELI_A && temPER_ELI_B) return 1;

      if (servA.includes("ECONOMICO") && temPER_A && !(servB.includes("ECONOMICO") && temPER_B)) return 1;
      if (servB.includes("ECONOMICO") && temPER_B && !(servA.includes("ECONOMICO") && temPER_A)) return -1;

      const servPrioA = getPrioridadeServico(servA);
      const servPrioB = getPrioridadeServico(servB);
      if (servPrioA !== servPrioB) return servPrioB - servPrioA;

      // Peso em ordem descendente (maior primeiro)
      if (a.peso !== b.peso) return b.peso - a.peso;

      return normalizeId(a.id).localeCompare(normalizeId(b.id));
    });

    for (const c of cargasOrdenadas) {
      todasCargas.push({ ...c, endereco: addr, enderecoInteiro: enderecoInteiro.has(addr) });
    }
  }

  // Renderizar tabela única - com limite de capacidade
  let html = `<table style="width: 100%; border-collapse: collapse;">`;
  html += `<thead><tr style="background: #f3f3f3;">
    <th style="border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 13px;">Endereço</th>
    <th style="border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 13px;">ID</th>
    <th style="border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 13px;">SLA</th>
    <th style="border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 13px;">Serviço</th>
    <th style="border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 13px;">SHC</th>
    <th style="border: 1px solid #ccc; padding: 6px; text-align: right; font-size: 13px;">Kg</th>
  </tr></thead><tbody>`;
  
  let ultimoEndereco = null;
  let pesoAcumulado = 0;
  let cargasIncluidas = 0;
  const cargasExcluidas = todasCargas.length;
  
  for (const c of todasCargas) {
    // Verificar se adicionar esta carga excederia a capacidade
    const pesoAtual = Number.isFinite(c.peso) ? c.peso : 0;
    if (pesoAcumulado + pesoAtual > voo.dispKg) {
      break; // Parou de caber, não adiciona mais
    }
    
    // Se é endereço inteiro, mostrar apenas a primeira linha com "INTEIRO"
    if (c.enderecoInteiro && c.endereco === ultimoEndereco) {
      continue; // Pular cargas duplicadas do mesmo endereço inteiro
    }
    
    let slaText = "-";
    if (c.slaPctRemaining !== null) {
      const pct = Math.round(c.slaPctRemaining);
      if (c.slaPctRemaining < 30) {
        slaText = `<30%`;
      } else if (c.slaPctRemaining < 50) {
        slaText = `<50%`;
      } else {
        slaText = `>${pct}%`;
      }
    }
    const servLabel = serviceLabelForRow(c);
    const shcLabel = shcLabelFromSet(c.shcTokens);
    const idDisplay = c.enderecoInteiro ? `(INTEIRO)` : escapeHtml(c.id);
    
    html += `<tr style="border: 1px solid #ddd;">
      <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-weight: 600; font-size: 13px;">${escapeHtml(c.endereco)}</td>
      <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;">${idDisplay}</td>
      <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;">${escapeHtml(slaText)}</td>
      <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;">${escapeHtml(servLabel)}</td>
      <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;">${escapeHtml(shcLabel)}</td>
      <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; text-align: right; font-size: 13px;">${formatNumber(c.peso)}</td>
    </tr>`;
    
    pesoAcumulado += pesoAtual;
    cargasIncluidas++;
    ultimoEndereco = c.endereco;
  }
  html += `</tbody></table>`;
  
  // Adicionar cargas em triagem
  const cargasTriagem = filteredRows();
  if (cargasTriagem.length > 0) {
    html += `<h3 style="margin-top: 30px; font-size: 16px; font-weight: 700; border-bottom: 2px solid #333; padding-bottom: 10px;">Cargas em Triagem</h3>`;
    html += `<table style="width: 100%; border-collapse: collapse;">`;
    html += `<thead><tr style="background: #f3f3f3;">
      <th style="border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 13px;">ID</th>
      <th style="border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 13px;">SLA</th>
      <th style="border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 13px;">Serviço</th>
      <th style="border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 13px;">SHC</th>
      <th style="border: 1px solid #ccc; padding: 6px; text-align: right; font-size: 13px;">Kg</th>
    </tr></thead><tbody>`;
    
    for (const c of cargasTriagem) {
      const destNorm = normalizeUpper(c.destino || "");
      const vooDestNorm = normalizeUpper(voo.destino);
      if (destNorm !== vooDestNorm) continue;
      
      // Aplicar filtro se necessário
      if (filtro) {
        const serv = normalizeUpper(c.servicoGroup || servicoGroupKey(c.servico || ""));
        let inclui = false;
        if (filtro === "SAUDE") {
          inclui = serv.includes("SAUDE");
        } else if (filtro === "URGENTE_EXPRESSO") {
          inclui = serv.includes("URGENTE") || serv === "EXPRESSO";
        } else if (filtro === "ECONOMICO") {
          inclui = serv.includes("ECO") || serv.includes("MODA") || serv.includes("ECONOMICO");
        }
        if (!inclui) continue;
      }
      
      let slaText = "-";
      if (c.slaPctRemaining !== null) {
        const pct = Math.round(c.slaPctRemaining);
        if (c.slaPctRemaining < 30) {
          slaText = `<30%`;
        } else if (c.slaPctRemaining < 50) {
          slaText = `<50%`;
        } else {
          slaText = `>${pct}%`;
        }
      }
      const servLabel = serviceLabelForRow(c);
      const shcLabel = shcLabelFromSet(c.shcTokens);
      
      html += `<tr style="border: 1px solid #ddd; background: #f9f9f9;">
        <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;">${escapeHtml(c.id)}</td>
        <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;">${escapeHtml(slaText)}</td>
        <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;">${escapeHtml(servLabel)}</td>
        <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; font-size: 13px;">${escapeHtml(shcLabel)}</td>
        <td style="border: 1px solid #ddd; padding: 6px; font-family: monospace; text-align: right; font-size: 13px;">${formatNumber(c.peso)}</td>
      </tr>`;
    }
    html += `</tbody></table>`;
  }

  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const stamp = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
  const cargasNaoIncluidas = cargasExcluidas - cargasIncluidas;
  const avisoCapacidade = cargasNaoIncluidas > 0 
    ? `Cargas incluídas: ${cargasIncluidas}/${cargasExcluidas} (${cargasNaoIncluidas} excluídas por capacidade)`
    : `Todas as ${cargasIncluidas} cargas estão incluídas`;

  return `<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Programação de Voo</title>
  <style>
    body { font-family: system-ui, -apple-system, "Segoe UI", Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; font-size: 14px; }
    .container { max-width: 1400px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; }
    .header { border-bottom: 2px solid #333; padding-bottom: 15px; margin-bottom: 15px; }
    .title { font-size: 20px; font-weight: 800; margin: 0 0 8px 0; }
    .meta { color: #666; font-size: 13px; }
    .info { background: #cfe2ff; border: 1px solid #0d6efd; color: #084298; padding: 12px; border-radius: 4px; margin-bottom: 15px; font-size: 13px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="title">Programação de Voo</div>
      <div class="meta">Voo: ${escapeHtml(voo.controle)} | ${escapeHtml(voo.destino)} | ${escapeHtml(voo.data)} ${escapeHtml(voo.voo)} | Gerado: ${escapeHtml(stamp)}</div>
    </div>
    
    <div class="info">Peso carregado: ${formatNumber(pesoAcumulado)}kg / ${formatNumber(voo.dispKg)}kg • ${avisoCapacidade}</div>
    
    ${html}
  </div>
</body>
</html>`;
}

function openProgramacaoVoo(voo, filtro = null) {
  const html = buildProgramacaoVooHtml(voo, filtro);
  const blob = new Blob([html], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  window.open(url, "_blank");
}

function showFiltroServico(voo) {
  const modal = document.createElement("div");
  modal.className = "modal fade";
  modal.id = "filtroServicoModal";
  modal.tabIndex = "-1";
  modal.innerHTML = `
    <div class="modal-dialog modal-dialog-centered">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title">Selecione o tipo de carga</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
        </div>
        <div class="modal-body" style="gap: 10px; display: flex; flex-direction: column;">
          <button class="btn btn-outline-primary" style="font-size: 16px; padding: 12px;">Apenas SAÚDE</button>
          <button class="btn btn-outline-primary" style="font-size: 16px; padding: 12px;">URGENTE e EXPRESSO</button>
          <button class="btn btn-outline-primary" style="font-size: 16px; padding: 12px;">Apenas ECONOMICO</button>
          <button class="btn btn-outline-primary" style="font-size: 16px; padding: 12px;">MISTO (Todos os Serviços)</button>
        </div>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
  const bsModal = new bootstrap.Modal(modal, { backdrop: "static", keyboard: false });
  
  const buttons = modal.querySelectorAll("button.btn-outline-primary");
  buttons[0].addEventListener("click", () => {
    bsModal.hide();
    setTimeout(() => modal.remove(), 300);
    openProgramacaoVoo(voo, "SAUDE");
  });
  buttons[1].addEventListener("click", () => {
    bsModal.hide();
    setTimeout(() => modal.remove(), 300);
    openProgramacaoVoo(voo, "URGENTE_EXPRESSO");
  });
  buttons[2].addEventListener("click", () => {
    bsModal.hide();
    setTimeout(() => modal.remove(), 300);
    openProgramacaoVoo(voo, "ECONOMICO");
  });
  buttons[3].addEventListener("click", () => {
    bsModal.hide();
    setTimeout(() => modal.remove(), 300);
    openProgramacaoVoo(voo, null);
  });
  
  bsModal.show();
}

function bindUi() {
  $("btnReset")?.addEventListener("click", clearAll);

  $("fileCargas")?.addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      await loadCsvFromFile(file);
    } catch (err) {
      setError(err?.message || "Falha ao processar o CSV.");
    }
  });

  $("fileVoos")?.addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      await loadVoosFromFile(file);
    } catch (err) {
      setText("programacaoError", err?.message || "Falha ao processar o arquivo de voos.");
    }
  });

  // Filtros de voos
  $("filtroDataSwitch")?.addEventListener("change", (e) => {
    stateVoos.filtroData = e.target.checked ? "amanha" : "hoje";
    setText("filtroDataLabel", stateVoos.filtroData === "hoje" ? "Hoje" : "Amanhã");
    renderVoosTable();
  });

  $("selectTerminal")?.addEventListener("change", (e) => {
    stateVoos.filtroTerminal = e.target.value;
    renderVoosTable();
  });

  $("filterServico")?.addEventListener("change", () => renderAll());
  $("headerCheck")?.addEventListener("click", (e) => {
    e.preventDefault();
    openCheckPraca();
  });
  $("headerOtimizacao")?.addEventListener("click", (e) => {
    e.preventDefault();
    openOtimizacao();
  });
  $("btnAuditoria")?.addEventListener("click", (e) => {
    e.preventDefault();
    openAuditoria();
  });
}

document.addEventListener("DOMContentLoaded", () => {
  enableModalStacking();
  bindUi();
  clearAll();
});
