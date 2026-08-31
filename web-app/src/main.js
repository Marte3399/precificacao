import './style.css';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';

const profiles = [
  { key: 'arquiteto', label: '5.2.1.3. Arquiteto de Sistema - Nível 2', shortLabel: 'Arquiteto', rate: 292.36 },
  { key: 'desenvolvedor', label: '5.2.1.4. Desenvolvedor - Nível 4', shortLabel: 'Desenvolvedor', rate: 204.24 },
  { key: 'requisito', label: '5.2.1.2. Analista de Requisito/Sistema - Nível 4', shortLabel: 'Requisito', rate: 243.56 },
  { key: 'coordenador', label: '5.2.1.8. Coordenador de Sistemas - Nível 2', shortLabel: 'Coordenador', rate: 246.47 }
];

const multilinePlaceholders = new Set([
  'SOLICITACAO2',
  'FUNCIONALIDADES3',
  'ATIVIDADES4',
  'OUTRAS5',
  'ATIVIDADES_ARQUITETO',
  'ATIVIDADES_REQUISITO',
  'ATIVIDADES_DESENVOLVEDOR',
  'ATIVIDADES_COORDENADOR'
]);

const paragraphFormattingRules = {
  SOLICITACAO2: { alignLeft: true, clearColor: true, normalizeTypography: true },
  FUNCIONALIDADES3: { alignLeft: true, clearColor: true, normalizeTypography: true },
  ATIVIDADES4: { alignLeft: true, clearColor: true, normalizeTypography: true },
  OUTRAS5: { alignLeft: true, clearColor: true, normalizeTypography: true }
};

const WORD_MAIN_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const TOTAL_ROW_SHADING = 'c9ab83';
const WORD_DEFAULT_FONT_FAMILY = 'Calibri';
const WORD_DEFAULT_FONT_SIZE = '22'; // half-points (22 = 11pt)
const REFERENCE_BODY_TEXT_FRAGMENT = 'o s-codes, é um sistema informatizado';

function ensureRunProperties(run, xmlDoc) {
  let rPr = Array.from(run.childNodes).find((child) => child.nodeName === 'w:rPr');
  if (!rPr) {
    rPr = xmlDoc.createElement('w:rPr');
    run.insertBefore(rPr, run.firstChild);
  }
  return rPr;
}

function setRunFontStyle(run, xmlDoc) {
  const rPr = ensureRunProperties(run, xmlDoc);

  let rFonts = Array.from(rPr.childNodes).find((child) => child.nodeName === 'w:rFonts');
  if (!rFonts) {
    rFonts = xmlDoc.createElement('w:rFonts');
    rPr.appendChild(rFonts);
  }
  rFonts.setAttribute('w:ascii', WORD_DEFAULT_FONT_FAMILY);
  rFonts.setAttribute('w:hAnsi', WORD_DEFAULT_FONT_FAMILY);
  rFonts.setAttribute('w:cs', WORD_DEFAULT_FONT_FAMILY);

  let sz = Array.from(rPr.childNodes).find((child) => child.nodeName === 'w:sz');
  if (!sz) {
    sz = xmlDoc.createElement('w:sz');
    rPr.appendChild(sz);
  }
  sz.setAttribute('w:val', WORD_DEFAULT_FONT_SIZE);

  let szCs = Array.from(rPr.childNodes).find((child) => child.nodeName === 'w:szCs');
  if (!szCs) {
    szCs = xmlDoc.createElement('w:szCs');
    rPr.appendChild(szCs);
  }
  szCs.setAttribute('w:val', WORD_DEFAULT_FONT_SIZE);

  return rPr;
}

function createWordParagraph(xmlDoc, text = '', { bold = false, align } = {}) {
  const paragraph = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:p');
  if (align) {
    let pPr = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:pPr');
    const jc = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:jc');
    jc.setAttribute('w:val', align);
    pPr.appendChild(jc);
    paragraph.appendChild(pPr);
  }

  const lines = (text ?? '').split('\n');
  if (text === '') {
    paragraph.appendChild(xmlDoc.createElementNS(WORD_MAIN_NS, 'w:r'));
    return paragraph;
  }

  lines.forEach((line, index) => {
    const run = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:r');
    setRunFontStyle(run, xmlDoc);
    if (bold) {
      const rPr = ensureRunProperties(run, xmlDoc);
      rPr.appendChild(xmlDoc.createElementNS(WORD_MAIN_NS, 'w:b'));
    }
    const textElement = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:t');
    if (/^\s|\s$/.test(line)) {
      textElement.setAttributeNS(XML_NS, 'xml:space', 'preserve');
    }
    textElement.textContent = line;
    run.appendChild(textElement);
    paragraph.appendChild(run);
    if (index < lines.length - 1) {
      run.appendChild(xmlDoc.createElementNS(WORD_MAIN_NS, 'w:br'));
    }
  });

  return paragraph;
}

function createTableCell(xmlDoc, { text = '', bold = false, align = 'center', shading, width }) {
  const cell = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tc');
  const tcPr = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tcPr');
  const tcW = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tcW');
  tcW.setAttribute('w:type', 'auto');
  tcW.setAttribute('w:w', width || '0');
  tcPr.appendChild(tcW);
  const tcBorders = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tcBorders');
  ['top', 'left', 'bottom', 'right'].forEach((side) => {
    const border = xmlDoc.createElementNS(WORD_MAIN_NS, `w:${side}`);
    border.setAttribute('w:val', 'single');
    border.setAttribute('w:sz', '8');
    border.setAttribute('w:space', '0');
    border.setAttribute('w:color', '000000');
    tcBorders.appendChild(border);
  });
  tcPr.appendChild(tcBorders);
  if (shading) {
    const shd = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:shd');
    shd.setAttribute('w:val', 'clear');
    shd.setAttribute('w:color', 'auto');
    shd.setAttribute('w:fill', shading);
    tcPr.appendChild(shd);
  }
  cell.appendChild(tcPr);
  cell.appendChild(createWordParagraph(xmlDoc, text, { bold, align }));
  return cell;
}

function createTableRow(xmlDoc, cells) {
  const row = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tr');
  cells.forEach((cell) => {
    row.appendChild(createTableCell(xmlDoc, cell));
  });
  return row;
}

function createTable(xmlDoc, columnWidths = []) {
  const table = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tbl');
  const tblPr = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tblPr');
  const tblStyle = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tblStyle');
  tblStyle.setAttribute('w:val', 'TableGrid');
  const tblW = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tblW');
  tblW.setAttribute('w:type', 'auto');
  tblW.setAttribute('w:w', '0');
  const tblBorders = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tblBorders');
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].forEach((side) => {
    const border = xmlDoc.createElementNS(WORD_MAIN_NS, `w:${side}`);
    border.setAttribute('w:val', 'single');
    border.setAttribute('w:sz', '8');
    border.setAttribute('w:space', '0');
    border.setAttribute('w:color', '000000');
    tblBorders.appendChild(border);
  });
  const tblLook = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tblLook');
  tblLook.setAttribute('w:val', '04A0');
  tblLook.setAttribute('w:firstRow', '1');
  tblLook.setAttribute('w:lastRow', '0');
  tblLook.setAttribute('w:firstColumn', '1');
  tblLook.setAttribute('w:lastColumn', '0');
  tblLook.setAttribute('w:noHBand', '0');
  tblLook.setAttribute('w:noVBand', '1');
  tblPr.appendChild(tblStyle);
  tblPr.appendChild(tblW);
  tblPr.appendChild(tblBorders);
  tblPr.appendChild(tblLook);
  table.appendChild(tblPr);

  const tblGrid = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:tblGrid');
  const widths = columnWidths.length ? columnWidths : new Array(5).fill('2400');
  widths.forEach((width) => {
    const gridCol = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:gridCol');
    gridCol.setAttribute('w:w', width);
    tblGrid.appendChild(gridCol);
  });
  table.appendChild(tblGrid);

  return table;
}

function buildActivitiesTable(xmlDoc, documentData) {
  const columnWidths = ['800', '4500', '1200', '1200', '1200', '1200'];
  const table = createTable(xmlDoc, columnWidths);
  const headerCells = [
    { text: 'Item da Atividade', bold: true, shading: 'e7e7e7' },
    { text: 'Requisitos / Atividades', bold: true, shading: 'e7e7e7' },
    { text: 'Arquiteto (H)', bold: true, shading: 'e7e7e7' },
    { text: 'Requisito (H)', bold: true, shading: 'e7e7e7' },
    { text: 'Desenvolvedor (H)', bold: true, shading: 'e7e7e7' },
    { text: 'Coordenador (H)', bold: true, shading: 'e7e7e7' }
  ];
  table.appendChild(createTableRow(xmlDoc, headerCells));

  documentData.enumeratedActivities.forEach((atividade, index) => {
    const hoursSnapshot = documentData.perActivityHours[index];
    const rowCells = [
      { text: `${index + 1}`, align: 'center' },
      { text: atividade, align: 'left' },
      { text: `${hoursSnapshot.arquiteto}`, align: 'center' },
      { text: `${hoursSnapshot.requisito}`, align: 'center' },
      { text: `${hoursSnapshot.desenvolvedor}`, align: 'center' },
      { text: `${hoursSnapshot.coordenador}`, align: 'center' }
    ];
    table.appendChild(createTableRow(xmlDoc, rowCells));
  });

  const totalRowCells = [
    { text: 'Total', bold: true, shading: TOTAL_ROW_SHADING, align: 'right' },
    { text: '', shading: TOTAL_ROW_SHADING },
    { text: `${documentData.totals.arquiteto || 0}`, shading: TOTAL_ROW_SHADING },
    { text: `${documentData.totals.requisito || 0}`, shading: TOTAL_ROW_SHADING },
    { text: `${documentData.totals.desenvolvedor || 0}`, shading: TOTAL_ROW_SHADING },
    { text: `${documentData.totals.coordenador || 0}`, shading: TOTAL_ROW_SHADING }
  ];
  table.appendChild(createTableRow(xmlDoc, totalRowCells));

  return table;
}

function buildProfileTable(xmlDoc, documentData) {
  const columnWidths = ['3000', '4500', '1400', '1100', '1600'];
  const table = createTable(xmlDoc, columnWidths);
  const headerCells = [
    { text: 'Perfil Profissional', bold: true, shading: 'e7e7e7' },
    { text: 'Atividades', bold: true, shading: 'e7e7e7' },
    { text: 'Valor-Hora', bold: true, shading: 'e7e7e7' },
    { text: 'QTD Horas', bold: true, shading: 'e7e7e7' },
    { text: 'Total', bold: true, shading: 'e7e7e7' }
  ];
  table.appendChild(createTableRow(xmlDoc, headerCells));

  profiles.forEach((profile) => {
    const horas = documentData.totals[profile.key] || 0;
    const totalPerfil = horas * profile.rate;
    const rowCells = [
      { text: profile.label, align: 'center', bold: true, shading: 'c3e9ec' },
      { text: documentData.activitiesByProfile[profile.key], align: 'left' },
      { text: formatCurrency(profile.rate) },
      { text: `${horas}` },
      { text: formatCurrency(totalPerfil) }
    ];
    table.appendChild(createTableRow(xmlDoc, rowCells));
  });

  const totalRowCells = [
    { text: '', shading: TOTAL_ROW_SHADING },
    { text: 'TOTAL:', bold: true, shading: TOTAL_ROW_SHADING, align: 'right' },
    { text: '', shading: TOTAL_ROW_SHADING },
    { text: `${documentData.totalHoras}`, bold: true, shading: TOTAL_ROW_SHADING },
    { text: formatCurrency(documentData.totalFinanceiro), bold: true, shading: TOTAL_ROW_SHADING }
  ];
  table.appendChild(createTableRow(xmlDoc, totalRowCells));

  return table;
}

function getParagraphText(paragraph) {
  const texts = paragraph.getElementsByTagName('w:t');
  let result = '';
  Array.from(texts).forEach((textNode) => {
    result += textNode.textContent || '';
  });
  return result;
}

function findParagraphByText(xmlDoc, targetText) {
  const paragraphs = xmlDoc.getElementsByTagName('w:p');
  const normalizedTarget = targetText.toLowerCase();
  for (let i = 0; i < paragraphs.length; i += 1) {
    const text = getParagraphText(paragraphs[i]).trim().toLowerCase();
    if (text === normalizedTarget) {
      return paragraphs[i];
    }
  }
  return null;
}

function getRecentMonthKeys(lastCount) {
  const now = new Date();
  const keys = [];
  for (let index = 0; index < lastCount; index += 1) {
    const date = new Date(now.getFullYear(), now.getMonth() - index, 1);
    keys.push(`${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`);
  }
  return new Set(keys);
}

function getSelectedDailyStatuses() {
  return new Set(dailyStatusCheckboxes().filter((cb) => cb.checked).map((cb) => cb.value));
}

function rowStatusFilterKey(statusValue) {
  const raw = String(statusValue || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '');
  if (raw.includes('finaliz')) return 'finalizado';
  if (raw.includes('andamento') || raw.includes('progresso')) return 'andamento';
  if (raw.includes('bloque') || raw.includes('imped')) return 'bloqueado';
  if (raw.includes('iniciado')) return 'naoiniciado';
  return raw || 'naoiniciado';
}

function rowMatchesSearch(row) {
  const query = (dailyGridSearch?.value || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (!query) return true;
  return dailyColumns.some((column) => {
    const value = String(row?.[column.key] || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    return value.includes(query);
  });
}

function shouldDisplayDailyRow(row) {
  if (!rowMatchesSearch(row)) return false;

  const selectedStatuses = getSelectedDailyStatuses();
  const statusKey = rowStatusFilterKey(row?.status);
  if (!selectedStatuses.has(statusKey)) return false;

  if (activeDailyFilterMode === 'all') return true;

  const rowMonth = parseMonthKeyFromDateValue(row?.entrada || row?.prazo || '');
  if (!rowMonth) return false;

  if (activeDailyFilterMode === 'month') {
    return rowMonth === dailyGridFilterMonth.value;
  }

  if (activeDailyFilterMode === 'last3') {
    return getRecentMonthKeys(3).has(rowMonth);
  }

  return true;
}

function findParagraphContainingText(xmlDoc, targetFragment) {
  const paragraphs = xmlDoc.getElementsByTagName('w:p');
  const normalizedTarget = targetFragment.toLowerCase();
  for (let i = 0; i < paragraphs.length; i += 1) {
    const text = getParagraphText(paragraphs[i]).trim().toLowerCase();
    if (text.includes(normalizedTarget)) {
      return paragraphs[i];
    }
  }
  return null;
}

function applyParagraphStyle(paragraph, xmlDoc, styleId) {
  if (!styleId) return;
  let pPr = Array.from(paragraph.childNodes).find((child) => child.nodeName === 'w:pPr');
  if (!pPr) {
    pPr = xmlDoc.createElement('w:pPr');
    paragraph.insertBefore(pPr, paragraph.firstChild);
  }

  let pStyle = Array.from(pPr.childNodes).find((child) => child.nodeName === 'w:pStyle');
  if (!pStyle) {
    pStyle = xmlDoc.createElement('w:pStyle');
    pPr.appendChild(pStyle);
  }

  pStyle.setAttribute('w:val', styleId);
}

function extractReferenceTypography(xmlDoc) {
  const referenceParagraph = findParagraphContainingText(xmlDoc, REFERENCE_BODY_TEXT_FRAGMENT);
  if (!referenceParagraph) {
    return {};
  }

  const pPr = Array.from(referenceParagraph.childNodes).find((child) => child.nodeName === 'w:pPr');
  const pStyle = pPr
    ? Array.from(pPr.childNodes).find((child) => child.nodeName === 'w:pStyle')
    : null;
  const paragraphStyleId = pStyle?.getAttribute('w:val') || pStyle?.getAttribute('val') || '';

  const runNodes = referenceParagraph.getElementsByTagName('w:r');
  let runProperties = null;
  for (let i = 0; i < runNodes.length; i += 1) {
    const rPr = Array.from(runNodes[i].childNodes).find((child) => child.nodeName === 'w:rPr');
    if (rPr) {
      runProperties = rPr.cloneNode(true);
      break;
    }
  }

  return {
    paragraphStyleId,
    runProperties
  };
}

function insertAfterReference(referenceNode, newNodes) {
  if (!newNodes.length) return;
  const parent = referenceNode?.parentNode || referenceNode;
  if (!parent) return;
  let cursor = referenceNode;
  newNodes.forEach((node) => {
    if (cursor && cursor.nextSibling) {
      parent.insertBefore(node, cursor.nextSibling);
    } else {
      parent.appendChild(node);
    }
    cursor = node;
  });
}

function insertEstimationTables(xmlDoc, documentData) {
  const body = xmlDoc.getElementsByTagName('w:body')[0];
  const anchor = findParagraphByText(xmlDoc, 'demanda:') || body.lastChild;
  const nodesToInsert = [];
  nodesToInsert.push(createWordParagraph(xmlDoc, '6.1.1. Estimativa com base nas Atividades', { bold: true }));
  nodesToInsert.push(buildActivitiesTable(xmlDoc, documentData));
  nodesToInsert.push(createWordParagraph(xmlDoc, ''));
  nodesToInsert.push(createWordParagraph(xmlDoc, '1.1.1. Estimativa com base nas atividades acima e listagem de Perfil e Valor', { bold: true }));
  nodesToInsert.push(buildProfileTable(xmlDoc, documentData));
  nodesToInsert.push(createWordParagraph(xmlDoc, ''));

  if (anchor && anchor.parentNode) {
    insertAfterReference(anchor, nodesToInsert);
  } else if (body) {
    nodesToInsert.forEach((node) => body.appendChild(node));
  }
}

const fixedActivities = [
  'Testes em ambiente de Desenvolvimento e Homologação.',
  'Implantação em Produção.'
];

const hourOptions = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];

const ICONS = {
  trash: '🗑',
  edit: '✎'
};

function createEmptyHoursMap() {
  const map = {};
  profiles.forEach(({ key }) => {
    map[key] = 0;
  });
  return map;
}

function normalizeDescription(text) {
  return text.replace(/\s+/g, ' ').trim().toLowerCase();
}

function parseHoursValue(rawValue) {
  if (rawValue == null || rawValue === '') return null;
  const normalized = String(rawValue).replace(',', '.').trim();
  const parsed = Number(normalized);
  if (!Number.isFinite(parsed) || parsed <= 0) return null;
  return parsed;
}

function formatHoursValue(value) {
  if (!Number.isFinite(value)) return '';
  return new Intl.NumberFormat('pt-BR', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  }).format(value);
}

const state = {
  manualActivities: [],
  activities: [],
  hours: {
    arquiteto: [],
    requisito: [],
    desenvolvedor: [],
    coordenador: []
  }
};

const DEFAULT_ACTIVITY_BUTTON_TEXT = 'Inserir atividade';
const EDIT_ACTIVITY_BUTTON_TEXT = 'Salvar alterações';

const builderState = {
  profile: null,
  hours: null,
  editingIndex: null
};

function rebuildActivitiesList() {
  const previousActivitiesLength = state.activities.length;
  const previousManualCount = Math.max(0, previousActivitiesLength - fixedActivities.length);
  const previousHours = {};
  profiles.forEach(({ key }) => {
    previousHours[key] = (state.hours[key] || []).slice();
  });

  const manualDescriptions = state.manualActivities.map((item) => item.description);
  state.activities = [...manualDescriptions, ...fixedActivities];

  profiles.forEach(({ key }) => {
    const manualValues = state.manualActivities.map((item) => item.hours[key] || 0);
    const existingFixed = previousHours[key]
      ? previousHours[key].slice(previousManualCount)
      : [];
    while (existingFixed.length < fixedActivities.length) {
      existingFixed.push(0);
    }
    state.hours[key] = [...manualValues, ...existingFixed.slice(0, fixedActivities.length)];
  });

  ensureHoursLength();
}

const templatePath = '/modelo_documento.docx';
let templateBuffer = null;
let templateLoadPromise = null;

function loadTemplateBuffer() {
  if (templateBuffer) {
    return Promise.resolve(templateBuffer);
  }

  if (!templateLoadPromise) {
    templateLoadPromise = fetch(templatePath)
      .then((response) => {
        if (!response.ok) {
          throw new Error(`Falha ao carregar o modelo oficial (status ${response.status})`);
        }
        return response.arrayBuffer();
      })
      .then((buffer) => {
        templateBuffer = buffer;
        return buffer;
      });
  }

  return templateLoadPromise;
}

const app = document.querySelector('#app');
app.innerHTML = `
  <div class="app-shell">
    <header class="header">
      <h1>Precificação e Geração de Estimativas</h1>
      <p>Preencha os dados do ticket, distribua as horas por perfil e gere automaticamente o documento Word pronto para envio.</p>
    </header>
    <div class="workspace-layout">
      <nav class="side-menu" aria-label="Módulos do sistema">
        <button type="button" class="side-menu__item side-menu__item--active" data-module="precificacao">Precificação</button>
        <button type="button" class="side-menu__item" data-module="daily">Planilha da Daily</button>
      </nav>
      <main class="content">
      <section class="module-view module-view--active" data-module-view="precificacao">
      <section class="section">
        <h2>Dados do Ticket</h2>
        <div class="field-grid">
          <div>
            <label>Número do Ticket</label>
            <input id="numeroTicket" type="text" placeholder="2026070210000971" />
          </div>
          <div>
            <label>Fator de Gordura</label>
            <input id="gordura" type="number" step="0.05" value="1.1" min="1" />
          </div>
        </div>
        <div class="field-grid field-grid-stack">
          <div>
            <label>Solicitação realizada pelo cliente</label>
            <textarea id="solicitacaoCliente" class="tall-textarea" placeholder="Descreva a demanda enviada..."></textarea>
          </div>
          <div>
            <label>Funcionalidades afetadas</label>
            <textarea id="funcionalidadesAfetadas" class="tall-textarea" placeholder="Liste os módulos, integrações ou impactos..."></textarea>
          </div>
        </div>
        <div>
          <label>Outras informações</label>
          <textarea id="outrasInformacoes" class="tall-textarea" placeholder="Observações adicionais"></textarea>
        </div>
      </section>

      <section class="section">
        <h2>Atividades</h2>
        <div class="activity-builder">
          <label>Descrição da atividade</label>
          <input id="atividadeDescricao" type="text" placeholder="Ex.: 4.1 Levantamento de dados" />
          <div class="activity-builder__group">
            <span>Selecione o perfil responsável</span>
            <div class="pill-group" id="perfilButtons"></div>
          </div>
          <div class="activity-builder__group">
            <span>Selecione as horas</span>
            <div class="pill-group" id="horaButtons"></div>
            <div class="manual-hours">
              <label for="horaManualInput">Ou digite manualmente</label>
              <input id="horaManualInput" type="text" inputmode="decimal" placeholder="Ex.: 5,5" />
            </div>
          </div>
          <div class="activity-builder__actions">
            <button class="primary" id="btnAdicionarAtividade">Inserir atividade</button>
            <p class="activity-hint">As atividades fixas serão acrescentadas automaticamente.</p>
          </div>
        </div>
        <div class="reference-panel">
          <h3>Referências rápidas</h3>
          <p>Será sempre acrescentado automaticamente:</p>
          <ul>
            ${fixedActivities.map((item) => `<li>${item}</li>`).join('')}
          </ul>
        </div>
        <div id="atividadesContainer" class="activities-list"></div>
      </section>

      <section class="section">
        <h2>Resumo de Horas</h2>
        <div class="summary-cards" id="summaryCards"></div>
        <div id="financeSection" style="display:none; flex-direction:column; gap:16px;">
          <table class="finance-table">
            <thead>
              <tr>
                <th>Perfil Profissional</th>
                <th>Atividades Selecionadas</th>
                <th>Valor-Hora</th>
                <th>Horas</th>
                <th>Total</th>
              </tr>
            </thead>
            <tbody id="financeTableBody"></tbody>
          </table>
        </div>
      </section>

      <section class="section">
        <h2>Quadros por Perfil</h2>
        <div id="profileBoards" class="summary-cards"></div>
      </section>

      <section class="section">
        <h2>Prévia textual</h2>
        <pre id="previewOutput" style="white-space:pre-wrap; font-family:'Space Grotesk', monospace;"></pre>
      </section>

      <div class="actions">
        <button class="primary" id="btnGerar">Gerar documento Word</button>
      </div>
      </section>

      <section class="module-view" data-module-view="daily">
        <section class="section">
          <h2>Planilha da Daily</h2>
          <p>Grid editável para acompanhamento diário. Clique em uma célula para editar.</p>
          <div id="dailySystemTabs" class="daily-system-tabs" role="tablist" aria-label="Sistemas da daily">
            <button type="button" class="daily-system-tab daily-system-tab--active" data-daily-system="scode">SCode</button>
            <button type="button" class="daily-system-tab" data-daily-system="siai">SIAI</button>
            <button type="button" class="daily-system-tab" data-daily-system="ctx">CTX</button>
            <button type="button" class="daily-system-tab" data-daily-system="sani">Sani</button>
            <button type="button" class="daily-system-tab" data-daily-system="sistrs">SISTRS</button>
            <button type="button" class="daily-system-tab" data-daily-system="opm">OPM</button>
            <button type="button" class="daily-system-tab" data-daily-system="outros">Outros</button>
          </div>
          <div class="daily-toolbar">
            <button type="button" class="secondary" id="btnDailyAddRow">Adicionar linha</button>
            <button type="button" class="secondary" id="btnDailySaveAll">Salvar planilha</button>
            <button type="button" class="secondary" id="btnDailyReset">Limpar planilha</button>
            <button type="button" class="secondary" id="btnDailyImportXlsx">Importar XLSX</button>
            <button type="button" class="secondary" id="btnDailyExportXlsx">Exportar XLSX</button>
            <input type="file" id="dailyXlsxInput" accept=".xlsx,.xls" hidden />
          </div>
          <div class="daily-report-controls">
            <label for="dailyReportMonth">Mês do relatório</label>
            <input type="month" id="dailyReportMonth" />
            <button type="button" class="secondary" id="btnDailyGenerateReport">Gerar relatório mensal</button>
            <button type="button" class="secondary" id="btnDailyPrintReport">Imprimir relatório</button>
          </div>
          <div class="daily-grid-filters">
            <label for="dailyGridSearch">Buscar</label>
            <input type="search" id="dailyGridSearch" placeholder="Buscar em todas as colunas" />
            <label for="dailyGridFilterMode">Filtro da grade</label>
            <select id="dailyGridFilterMode">
              <option value="month" selected>Mês selecionado</option>
              <option value="last3">Últimos 3 meses</option>
              <option value="all">Sem filtro</option>
            </select>
            <input type="month" id="dailyGridFilterMonth" />
            <fieldset class="daily-status-filter">
              <legend>Status</legend>
              <label><input type="checkbox" value="finalizado" checked> Finalizado</label>
              <label><input type="checkbox" value="andamento" checked> Em andamento</label>
              <label><input type="checkbox" value="bloqueado" checked> Bloqueado</label>
              <label><input type="checkbox" value="naoiniciado" checked> Não iniciado</label>
            </fieldset>
          </div>
          <div id="dailyColumnActions" class="daily-column-actions"></div>
          <div class="daily-grid-wrap">
            <table class="daily-grid" id="dailyGridTable">
              <thead id="dailyGridHead"></thead>
              <tbody id="dailyGridBody"></tbody>
            </table>
          </div>
          <div id="dailyMonthlyReport" class="daily-monthly-report"></div>
        </section>
      </section>
      </main>
    </div>
  </div>
`;

const atividadesContainer = document.querySelector('#atividadesContainer');
const summaryCards = document.querySelector('#summaryCards');
const financeTableBody = document.querySelector('#financeTableBody');
const financeSection = document.querySelector('#financeSection');
const profileBoards = document.querySelector('#profileBoards');
const previewOutput = document.querySelector('#previewOutput');
const atividadeDescricaoInput = document.querySelector('#atividadeDescricao');
const perfilButtonsContainer = document.querySelector('#perfilButtons');
const horaButtonsContainer = document.querySelector('#horaButtons');
const horaManualInput = document.querySelector('#horaManualInput');
const btnAdicionarAtividade = document.querySelector('#btnAdicionarAtividade');
const btnDailyAddRow = document.querySelector('#btnDailyAddRow');
const btnDailySaveAll = document.querySelector('#btnDailySaveAll');
const btnDailyReset = document.querySelector('#btnDailyReset');
const btnDailyImportXlsx = document.querySelector('#btnDailyImportXlsx');
const btnDailyExportXlsx = document.querySelector('#btnDailyExportXlsx');
const dailyXlsxInput = document.querySelector('#dailyXlsxInput');
const btnDailyGenerateReport = document.querySelector('#btnDailyGenerateReport');
const btnDailyPrintReport = document.querySelector('#btnDailyPrintReport');
const dailyColumnActions = document.querySelector('#dailyColumnActions');
const dailyGridHead = document.querySelector('#dailyGridHead');
const dailyGridBody = document.querySelector('#dailyGridBody');
const dailyGridFilterMode = document.querySelector('#dailyGridFilterMode');
const dailyGridFilterMonth = document.querySelector('#dailyGridFilterMonth');
const dailyGridSearch = document.querySelector('#dailyGridSearch');
const dailyStatusFilter = document.querySelector('.daily-status-filter');
const dailyStatusCheckboxes = () => Array.from(dailyStatusFilter.querySelectorAll('input[type="checkbox"]'));
const dailyReportMonthInput = document.querySelector('#dailyReportMonth');

const dailyAutocompleteContainer = document.createElement('div');
dailyAutocompleteContainer.id = 'dailyAutocomplete';
dailyAutocompleteContainer.className = 'daily-autocomplete';
dailyAutocompleteContainer.style.position = 'absolute';
dailyAutocompleteContainer.style.display = 'none';
dailyAutocompleteContainer.style.zIndex = '1000';
document.body.appendChild(dailyAutocompleteContainer);

let dailyAutocompleteActiveTd = null;
let dailyAutocompleteSuggestions = [];
let dailyAutocompleteIndex = -1;
const dailyMonthlyReport = document.querySelector('#dailyMonthlyReport');
const dailySystemButtons = Array.from(document.querySelectorAll('.daily-system-tab'));
const moduleButtons = Array.from(document.querySelectorAll('.side-menu__item'));
const moduleViews = Array.from(document.querySelectorAll('.module-view'));
let activeModule = 'precificacao';

const DAILY_ROWS_STORAGE_PREFIX = 'precificacao.daily.grid.rows.v2.';
const DAILY_MIN_ROWS = 20;
const DAILY_LEGACY_STORAGE_KEY = 'precificacao.daily.grid.rows.v1';
const DAILY_API_BASE_URL = '/api/daily-grid';
const dailySystems = [
  { key: 'scode', label: 'SCode' },
  { key: 'siai', label: 'SIAI' },
  { key: 'ctx', label: 'CTX' },
  { key: 'sani', label: 'Sani' },
  { key: 'sistrs', label: 'SISTRS' },
  { key: 'opm', label: 'OPM' },
  { key: 'outros', label: 'Outros' }
];
let activeDailySystem = 'scode';

const dailyColumns = [
  { key: 'prioridade', label: 'Prioridade' },
  { key: 'ticket', label: 'Nome da tarefa' },
  { key: 'descricao', label: 'Observação' },
  { key: 'status', label: 'Status' },
  { key: 'responsavel', label: 'Atribuído a' },
  { key: 'entrada', label: 'Data de início' },
  { key: 'prazo', label: 'Data do término' },
  { key: 'entrega', label: 'Duração' },
  { key: 'observacoes', label: 'Obs.' }
];

const monthNamesPtBr = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
const monthAbbreviationsPtBr = {
  jan: 1,
  fev: 2,
  mar: 3,
  abr: 4,
  mai: 5,
  jun: 6,
  jul: 7,
  ago: 8,
  set: 9,
  out: 10,
  nov: 11,
  dez: 12
};

const dailySheetBySystem = {
  scode: ['Scodes', 'SCode', 'S CODES'],
  siai: ['SIAI', 'SIA'],
  ctx: ['CTX'],
  sani: ['Sani', 'SANI'],
  sistrs: ['SISTRS', 'Sistrs'],
  opm: ['OPM'],
  outros: ['OUTROS', 'Outros']
};

function findSheetNameForSystem(workbook, systemKey) {
  const candidates = dailySheetBySystem[systemKey] || [systemKey];
  const normalizedMap = new Map(
    workbook.SheetNames.map((name) => [name.trim().toLowerCase(), name])
  );

  for (const candidate of candidates) {
    const found = normalizedMap.get(candidate.trim().toLowerCase());
    if (found) return found;
  }

  return null;
}

function mapSheetRowToDailyRow(sheetRow) {
  const normalizeHeader = (text) => String(text || '')
    .toLowerCase()
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');

  const headers = Object.keys(sheetRow).reduce((acc, key) => {
    acc[normalizeHeader(key)] = key;
    return acc;
  }, {});

  const valueByHeader = (candidates) => {
    for (const candidate of candidates) {
      const normalized = normalizeHeader(candidate);
      const originalKey = headers[normalized];
      if (originalKey != null && sheetRow[originalKey] != null && String(sheetRow[originalKey]).trim() !== '') {
        return sheetRow[originalKey];
      }
    }
    return '';
  };

  const rawTermino = valueByHeader(['Data do Término', 'Data do Termino', 'Data do término', 'Data do termino', 'Datadotermino', 'Termino', 'Término', 'Data Termino', 'Data Término']);
  const rawInicio = valueByHeader(['Data de Início', 'Data de Inicio', 'Data de início', 'Data de inicio', 'Datadeinicio', 'Inicio', 'Data Inicio', 'Data Início']) || rawTermino;

  return {
    prioridade: valueByHeader(['Prioridade']),
    ticket: valueByHeader(['Nome da tarefa', 'Ticket', 'Tarefa']),
    descricao: valueByHeader(['Observação', 'Observacao', 'Observao']),
    status: valueByHeader(['Status']),
    responsavel: valueByHeader(['Atribuído a', 'Atribuido a', 'Atribuidoa', 'Responsável', 'Responsavel']),
    entrada: normalizeDateCellValue(rawInicio),
    prazo: normalizeDateCellValue(rawTermino),
    entrega: valueByHeader(['Duração', 'Duracao']),
    observacoes: valueByHeader(['Obs.', 'Obs'])
  };
}

function isDailyRowCompletelyEmpty(row) {
  return dailyColumns.every((column) => String(row?.[column.key] || '').trim() === '');
}

function locateHeaderRow(rawRows, start = 0, end = 10) {
  const normalize = (text) => String(text || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '');
  const importantHeaders = new Set(['prioridade', 'nomedatarefa', 'observacao', 'status', 'atribuidoa', 'atribuido', 'responsavel', 'datadeinicio', 'datadotermino', 'duracao', 'obs']);

  for (let index = start; index < Math.min(end, rawRows.length); index += 1) {
    const row = rawRows[index] || [];
    const normalized = row.map((cell) => normalize(cell));
    const matches = normalized.filter((cell) => importantHeaders.has(cell)).length;
    if (matches >= 3) {
      return { rowIndex: index, headers: row };
    }
  }

  return { rowIndex: 0, headers: rawRows[0] || [] };
}

async function importDailyFromWorkbook(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });

  const importedSystems = [];

  for (const system of dailySystems) {
    const sheetName = findSheetNameForSystem(workbook, system.key);
    if (!sheetName) continue;

    const sheet = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const { rowIndex: headerIndex, headers } = locateHeaderRow(rawRows);

    const jsonRows = [];
    for (let index = headerIndex + 1; index < rawRows.length; index += 1) {
      const row = rawRows[index];
      const mapped = {};
      for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
        const header = headers[columnIndex] ?? '';
        mapped[header] = row?.[columnIndex] ?? '';
      }
      jsonRows.push(mapped);
    }

    const mappedRows = jsonRows
      .map((sheetRow) => mapSheetRowToDailyRow(sheetRow))
      .filter((row) => !isDailyRowCompletelyEmpty(row))
      .filter((row) => !isDailyIgnoredRow(row));

    const normalized = normalizeDailyRows(mappedRows);
    while (normalized.length < DAILY_MIN_ROWS) {
      normalized.push(createEmptyDailyRow());
    }

    localStorage.setItem(getDailyRowsStorageKey(system.key), JSON.stringify(normalized));
    await saveDailyRowsToApi(system.key, normalized);
    importedSystems.push(system.label);
  }

  if (importedSystems.length === 0) {
    alert('Nenhuma aba reconhecida foi encontrada na planilha.');
    return;
  }

  loadDailyRows();
  ensureEmptyTrailingRow(dailyRows.length - 1, false);
  renderDailyGrid();
  alert(`Importação concluída para: ${importedSystems.join(', ')}`);
}

async function exportDailyToWorkbook() {
  const workbook = XLSX.utils.book_new();

  for (const system of dailySystems) {
    let systemRows = [];
    try {
      systemRows = await fetchDailyRowsFromApi(system.key);
    } catch (_error) {
      const stored = localStorage.getItem(getDailyRowsStorageKey(system.key));
      if (stored) {
        try {
          systemRows = normalizeDailyRows(JSON.parse(stored));
        } catch (_parseError) {
          systemRows = [];
        }
      }
    }

    const outputRows = systemRows
      .filter((row) => !isDailyRowCompletelyEmpty(row))
      .map((row) => ({
        Prioridade: row.prioridade || '',
        'Nome da tarefa': row.ticket || '',
        Observação: row.descricao || '',
        Status: row.status || '',
        'Atribuído a': row.responsavel || '',
        'Data de início': row.entrada || '',
        'Data do término': row.prazo || '',
        Duração: row.entrega || '',
        'Obs.': row.observacoes || ''
      }));

    const worksheet = XLSX.utils.json_to_sheet(outputRows);
    XLSX.utils.book_append_sheet(workbook, worksheet, system.label);
  }

  XLSX.writeFile(workbook, `Daily Atualizada ${new Date().toISOString().slice(0, 10)}.xlsx`);
}

let dailyRows = [];
let activeDailyFilterMode = 'month';

function setActiveModule(moduleName) {
  activeModule = moduleName;
  moduleButtons.forEach((button) => {
    const isActive = button.dataset.module === moduleName;
    button.classList.toggle('side-menu__item--active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  });
  moduleViews.forEach((view) => {
    view.classList.toggle('module-view--active', view.dataset.moduleView === moduleName);
  });
}

function getDailyRowsStorageKey(systemKey = activeDailySystem) {
  return `${DAILY_ROWS_STORAGE_PREFIX}${systemKey}`;
}


function setActiveDailySystem(systemKey) {
  if (!dailySystems.some((item) => item.key === systemKey)) return;
  activeDailySystem = systemKey;
  dailySystemButtons.forEach((button) => {
    const isActive = button.dataset.dailySystem === systemKey;
    button.classList.toggle('daily-system-tab--active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  });
  loadDailyRows();
  ensureEmptyTrailingRow(dailyRows.length - 1, false);
  renderDailyGrid();
}

function createEmptyDailyRow() {
  const row = {};
  dailyColumns.forEach((column) => {
    row[column.key] = '';
  });
  return row;
}

function ensureDailyMinimumRows() {
  while (dailyRows.length < DAILY_MIN_ROWS) {
    dailyRows.push(createEmptyDailyRow());
  }
}

function normalizeDailyRows(rows) {
  if (!Array.isArray(rows)) return [];
  return rows.map((row) => {
    const normalized = createEmptyDailyRow();
    dailyColumns.forEach((column) => {
      let cellValue = row[column.key] || '';
      if (column.key === 'prioridade') cellValue = normalizePrioridade(cellValue);
      if (column.key === 'entrada' || column.key === 'prazo') cellValue = normalizeDateCellValue(cellValue);
      if (column.key === 'entrega') {
        const numeric = Number(cellValue);
        if (!Number.isNaN(numeric) && numeric >= 10000) {
          cellValue = normalizeDateCellValue(cellValue);
        }
      }
      normalized[column.key] = cellValue;
    });
    return normalized;
  });
}

function convertExcelSerialToDateString(value) {
  const parsed = XLSX.SSF.parse_date_code(Number(value));
  if (!parsed || !parsed.y || !parsed.m || !parsed.d) return '';
  const day = String(parsed.d).padStart(2, '0');
  const month = String(parsed.m).padStart(2, '0');
  return `${day}/${month}/${parsed.y}`;
}

function normalizePrioridade(value) {
  const raw = String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (raw.includes('alt')) return 'Alto';
  if (raw.includes('medi') || raw.includes('medio')) return 'Médio';
  if (raw.includes('bai')) return 'Baixo';
  return String(value || '').trim();
}

function normalizeDateCellValue(value) {
  if (value == null) return '';
  if (typeof value === 'number') {
    const converted = convertExcelSerialToDateString(value);
    return converted || String(value);
  }

  const raw = String(value).trim();
  if (!raw) return '';

  if (/^\d+(?:\.\d+)?$/.test(raw)) {
    const numeric = Number(raw);
    if (numeric >= 20000 && numeric <= 80000) {
      const converted = convertExcelSerialToDateString(numeric);
      if (converted) return converted;
    }
  }

  return raw;
}

async function fetchDailyRowsFromApi(systemKey = activeDailySystem) {
  const response = await fetch(`${DAILY_API_BASE_URL}/${encodeURIComponent(systemKey)}`);
  if (!response.ok) {
    throw new Error(`Falha ao carregar daily (${response.status})`);
  }
  const payload = await response.json();
  return normalizeDailyRows(payload?.rows || []);
}

async function saveDailyRowsToApi(systemKey = activeDailySystem, rowsPayload = dailyRows) {
  try {
    await fetch(`${DAILY_API_BASE_URL}/${encodeURIComponent(systemKey)}`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ rows: rowsPayload })
    });
  } catch (error) {
    console.warn('API Daily indisponível. Mantendo dados no localStorage.', error);
  }
}

async function syncDailyRowsFromApi(systemKey = activeDailySystem) {
  try {
    const remoteRows = await fetchDailyRowsFromApi(systemKey);
    if (!remoteRows.length) {
      const hasLocalData = dailyRows.some((row) => dailyColumns.some((column) => String(row?.[column.key] || '').trim() !== ''));
      if (hasLocalData) {
        const snapshot = dailyRows.map((row) => ({ ...row }));
        await saveDailyRowsToApi(systemKey, snapshot);
      }
      return;
    }
    if (systemKey !== activeDailySystem) return;

    dailyRows = remoteRows;
    ensureDailyMinimumRows();
    localStorage.setItem(getDailyRowsStorageKey(systemKey), JSON.stringify(dailyRows));
    renderDailyGrid();
  } catch (error) {
    console.warn('Não foi possível sincronizar Daily com API. Usando dados locais.', error);
  }
}

function saveDailyRows() {
  const snapshot = dailyRows.map((row) => ({ ...row }));
  localStorage.setItem(getDailyRowsStorageKey(), JSON.stringify(dailyRows));
  void saveDailyRowsToApi(activeDailySystem, snapshot);
}

function escapeHtml(text) {
  return String(text ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function parseMonthKeyFromDateValue(value) {
  const raw = normalizeDateCellValue(value).trim().toLowerCase();
  if (!raw) return null;

  let match = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (match) {
    const month = Number(match[2]);
    let year = Number(match[3]);
    if (Number.isNaN(month) || month < 1 || month > 12) return null;
    if (year < 100) year += 2000;
    return `${year}-${String(month).padStart(2, '0')}`;
  }

  match = raw.match(/^(\d{1,2})\/([a-zç]{3})\/(\d{2,4})$/);
  if (match) {
    const month = monthAbbreviationsPtBr[match[2]];
    if (!month) return null;
    let year = Number(match[3]);
    if (year < 100) year += 2000;
    return `${year}-${String(month).padStart(2, '0')}`;
  }

  match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
    return `${match[1]}-${match[2]}`;
  }

  return null;
}

function formatMonthLabel(monthKey) {
  if (!monthKey || !/^\d{4}-\d{2}$/.test(monthKey)) return '';
  const [yearRaw, monthRaw] = monthKey.split('-');
  const monthIndex = Number(monthRaw) - 1;
  if (monthIndex < 0 || monthIndex > 11) return monthKey;
  return `${monthNamesPtBr[monthIndex]} de ${yearRaw}`;
}

function getStatusReportClass(statusValue) {
  const normalized = String(statusValue || '').trim().toLowerCase();
  if (normalized.includes('finaliz')) return 'is-done';
  if (normalized.includes('andamento') || normalized.includes('progresso')) return 'is-progress';
  if (normalized.includes('bloque')) return 'is-blocked';
  if (normalized.includes('iniciado')) return 'is-not-started';
  return '';
}

function getDailyReportTheme(systemKey) {
  const themes = {
    scode: { titleBg: '#2d8ec3' },
    siai: { titleBg: '#6f569d' },
    sani: { titleBg: '#1fae61' },
    opm: { titleBg: '#ea7f1b' },
    sistrs: { titleBg: '#22a7dc' }
  };
  return themes[systemKey] || { titleBg: '#2d8ec3' };
}

function isDailyIgnoredRow(row) {
  return dailyColumns.some((column) => String(row?.[column.key] || '').toUpperCase().includes('#VALOR!'));
}

function getMonthlyReportRows(monthKey) {
  return dailyRows
    .filter((row) => {
      if (isDailyIgnoredRow(row)) return false;
      if (!String(row.ticket || '').trim()) return false;
      const rowMonth = parseMonthKeyFromDateValue(row.entrada);
      return rowMonth === monthKey;
    })
    .map((row) => ({
      prioridade: String(row.prioridade || '').trim(),
      ticket: String(row.ticket || '').trim(),
      observacao: String(row.descricao || '').trim(),
      status: String(row.status || '').trim(),
      atribuida: String(row.responsavel || '').trim(),
      inicio: String(row.entrada || row.prazo || '').trim(),
      termino: String(row.prazo || '').trim(),
      duracao: String(row.entrega || '').trim(),
      obs: String(row.observacoes || '').trim(),
      statusClass: getStatusReportClass(row.status)
    }));
}

function renderDailyMonthlyReport() {
  const selectedMonth = dailyReportMonthInput.value;
  if (!selectedMonth) {
    alert('Selecione o mês para gerar o relatório.');
    return;
  }

  const currentSystem = dailySystems.find((item) => item.key === activeDailySystem);
  const systemLabel = currentSystem?.label || activeDailySystem.toUpperCase();
  const reportTheme = getDailyReportTheme(activeDailySystem);
  const monthLabel = formatMonthLabel(selectedMonth);
  const reportRows = getMonthlyReportRows(selectedMonth);

  const rowsHtml = reportRows.length
    ? reportRows
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.prioridade)}</td>
              <td>${escapeHtml(row.ticket)}</td>
              <td>${escapeHtml(row.observacao)}</td>
              <td class="${row.statusClass}">${escapeHtml(row.status)}</td>
              <td>${escapeHtml(row.atribuida)}</td>
              <td>${escapeHtml(row.inicio)}</td>
              <td>${escapeHtml(row.termino)}</td>
              <td>${escapeHtml(row.duracao)}</td>
              <td>${escapeHtml(row.obs)}</td>
            </tr>
          `
        )
        .join('')
    : `
      <tr>
        <td colspan="9" class="daily-report-empty">Nenhum registro encontrado para ${escapeHtml(monthLabel)}.</td>
      </tr>
    `;

  dailyMonthlyReport.innerHTML = `
    <section class="daily-report-card" style="--daily-report-title-bg: ${reportTheme.titleBg};">
      <div class="daily-report-card__title">● ${escapeHtml(systemLabel)}</div>
      <table class="daily-report-table">
        <thead>
          <tr>
            <th>Prioridade</th>
            <th>Nome da tarefa</th>
            <th>Observação</th>
            <th>Status</th>
            <th>Atribuído a</th>
            <th>Data de início</th>
            <th>Data do término</th>
            <th>Duração</th>
            <th>Obs.</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </section>
  `;
}

function printDailyMonthlyReport() {
  if (!dailyMonthlyReport.innerHTML.trim()) {
    alert('Gere o relatório mensal antes de imprimir.');
    return;
  }

  const printWindow = window.open('', '_blank', 'width=1200,height=800');
  if (!printWindow) {
    alert('Não foi possível abrir a janela de impressão. Verifique se o navegador bloqueou pop-up.');
    return;
  }

  const reportTheme = getDailyReportTheme(activeDailySystem);

  printWindow.document.write(`
    <!doctype html>
    <html lang="pt-BR">
      <head>
        <meta charset="UTF-8" />
        <title>Relatório Mensal Daily</title>
        <style>
          body { margin: 0; padding: 24px; background: #fff; font-family: Arial, Helvetica, sans-serif; }
          .daily-report-card { border: 1px solid #1c2d4f; }
          .daily-report-card__title { background: ${reportTheme.titleBg}; color: #fff; font-weight: 700; padding: 6px 10px; font-size: 18px; }
          .daily-report-table { width: 100%; border-collapse: collapse; table-layout: fixed; }
          .daily-report-table th, .daily-report-table td { border: 1px solid #1c2d4f; padding: 6px 8px; font-size: 14px; }
          .daily-report-table th { background: #1f3e67; color: #fff; text-align: center; }
          .daily-report-table td { background: #c9d5e4; }
          .daily-report-table tbody tr:nth-child(even) td { background: #ffffff; }
          .daily-report-table td:nth-child(1) { width: 7%; }
          .daily-report-table td:nth-child(2) { width: 22%; }
          .daily-report-table td:nth-child(3) { width: 22%; }
          .daily-report-table td:nth-child(4) { width: 11%; font-weight: 700; }
          .daily-report-table td:nth-child(5) { width: 9%; }
          .daily-report-table td:nth-child(6), .daily-report-table td:nth-child(7) { width: 10%; text-align: center; }
          .daily-report-table td:nth-child(8) { width: 7%; text-align: center; }
          .daily-report-table td:nth-child(9) { width: 12%; }
          .daily-report-table td.is-done { background: inherit !important; color: #000; }
          .daily-report-table td.is-progress { background: inherit !important; color: #000; }
          .daily-report-table td.is-blocked { background: inherit !important; color: #000; }
          .daily-report-table td.is-not-started { background: inherit !important; color: #000; }
          .daily-report-empty { text-align: center; font-style: italic; }
        </style>
      </head>
      <body>
        ${dailyMonthlyReport.innerHTML}
      </body>
    </html>
  `);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
}

function loadDailyRows() {
  dailyRows = [];
  try {
    const storedRowsRaw = localStorage.getItem(getDailyRowsStorageKey());
    if (storedRowsRaw) {
      const parsed = JSON.parse(storedRowsRaw);
      if (Array.isArray(parsed)) {
        dailyRows = normalizeDailyRows(parsed);
      }
    }

    if (!storedRowsRaw && activeDailySystem === 'scode') {
      const legacyRaw = localStorage.getItem(DAILY_LEGACY_STORAGE_KEY);
      if (legacyRaw) {
        const parsedLegacy = JSON.parse(legacyRaw);
        if (Array.isArray(parsedLegacy)) {
          dailyRows = normalizeDailyRows(parsedLegacy);
        }
      }
    }

    if (!dailyRows.length) {
      dailyRows = new Array(DAILY_MIN_ROWS).fill(null).map(() => createEmptyDailyRow());
    }

  } catch (error) {
    console.error('Falha ao carregar grade da daily:', error);
    dailyRows = new Array(DAILY_MIN_ROWS).fill(null).map(() => createEmptyDailyRow());
  }

  ensureDailyMinimumRows();
  void syncDailyRowsFromApi(activeDailySystem);
}

function getStatusClass(statusValue) {
  const normalized = String(statusValue || '').trim().toLowerCase();
  if (normalized.includes('finaliz')) return 'daily-status--done';
  if (normalized.includes('andamento') || normalized.includes('progresso')) return 'daily-status--progress';
  if (normalized.includes('bloque')) return 'daily-status--blocked';
  if (normalized.includes('iniciado')) return 'daily-status--not-started';
  return '';
}

function updateDailyCell(rowIndex, columnKey, value) {
  if (!dailyRows[rowIndex]) return;
  dailyRows[rowIndex][columnKey] = value;
  saveDailyRows();
}

function getDailyColumnSuggestions(columnKey, text) {
  const raw = String(text || '').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (!raw) return [];
  const seen = new Set();
  const matches = [];
  dailyRows.forEach((row) => {
    const value = String(row[columnKey] || '').trim();
    if (!value) return;
    const normalized = value.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (normalized.startsWith(raw) && !seen.has(value)) {
      seen.add(value);
      matches.push(value);
    }
  });
  const predefined = {
    status: ['Finalizado', 'Em andamento', 'Bloqueado', 'Não Iniciado'],
    prioridade: ['Alto', 'Médio', 'Baixo'],
  };
  (predefined[columnKey] || []).forEach((value) => {
    const normalized = value.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (normalized.startsWith(raw) && !seen.has(value)) {
      seen.add(value);
      matches.push(value);
    }
  });
  return matches.slice(0, 10);
}

function positionDailyAutocomplete(td) {
  const rect = td.getBoundingClientRect();
  dailyAutocompleteContainer.style.left = `${rect.left + window.scrollX}px`;
  dailyAutocompleteContainer.style.top = `${rect.bottom + window.scrollY}px`;
  dailyAutocompleteContainer.style.minWidth = `${rect.width}px`;
}

function showDailyAutocomplete(td, suggestions) {
  dailyAutocompleteActiveTd = td;
  dailyAutocompleteSuggestions = suggestions;
  dailyAutocompleteIndex = -1;
  dailyAutocompleteContainer.innerHTML = '';
  if (suggestions.length === 0) {
    dailyAutocompleteContainer.style.display = 'none';
    return;
  }
  suggestions.forEach((value, index) => {
    const item = document.createElement('div');
    item.className = 'daily-autocomplete__item';
    item.textContent = value;
    item.dataset.index = String(index);
    item.addEventListener('mousedown', (event) => {
      event.preventDefault();
      applyDailyAutocomplete(index);
    });
    dailyAutocompleteContainer.appendChild(item);
  });
  positionDailyAutocomplete(td);
  dailyAutocompleteContainer.style.display = 'block';
}

function hideDailyAutocomplete() {
  dailyAutocompleteContainer.style.display = 'none';
  dailyAutocompleteActiveTd = null;
  dailyAutocompleteSuggestions = [];
  dailyAutocompleteIndex = -1;
}

function applyDailyAutocomplete(index) {
  if (!dailyAutocompleteActiveTd || index < 0 || index >= dailyAutocompleteSuggestions.length) return;
  const value = dailyAutocompleteSuggestions[index];
  const rowIndex = Number(dailyAutocompleteActiveTd.dataset.row);
  const columnKey = dailyAutocompleteActiveTd.dataset.column;
  dailyAutocompleteActiveTd.textContent = value;
  updateDailyCell(rowIndex, columnKey, value);
  hideDailyAutocomplete();
}

function highlightDailyAutocompleteItem() {
  Array.from(dailyAutocompleteContainer.children).forEach((child, index) => {
    child.classList.toggle('daily-autocomplete__item--active', index === dailyAutocompleteIndex);
  });
}

function handleDailyAutocompleteKey(event, td) {
  if (dailyAutocompleteContainer.style.display === 'none' || !dailyAutocompleteActiveTd) return;
  if (event.key === 'ArrowDown') {
    event.preventDefault();
    dailyAutocompleteIndex = (dailyAutocompleteIndex + 1) % dailyAutocompleteSuggestions.length;
    highlightDailyAutocompleteItem();
  } else if (event.key === 'ArrowUp') {
    event.preventDefault();
    dailyAutocompleteIndex = (dailyAutocompleteIndex - 1 + dailyAutocompleteSuggestions.length) % dailyAutocompleteSuggestions.length;
    highlightDailyAutocompleteItem();
  } else if (event.key === 'Enter' || event.key === 'Tab') {
    event.preventDefault();
    if (dailyAutocompleteIndex >= 0) {
      applyDailyAutocomplete(dailyAutocompleteIndex);
    }
  } else if (event.key === 'Escape') {
    hideDailyAutocomplete();
  }
}

function renderDailyGrid() {
  dailyGridHead.innerHTML = '';
  dailyGridBody.innerHTML = '';

  const headRow = document.createElement('tr');
  const indexHeader = document.createElement('th');
  indexHeader.textContent = '#';
  headRow.appendChild(indexHeader);
  dailyColumns.forEach((column) => {
    const th = document.createElement('th');
    th.textContent = column.label;
    headRow.appendChild(th);
  });
  dailyGridHead.appendChild(headRow);

  const lastRowIndex = dailyRows.length - 1;
  const rowsToRender = dailyRows
    .map((row, rowIndex) => ({ row, rowIndex }))
    .filter(({ row, rowIndex }) => rowIndex === lastRowIndex || shouldDisplayDailyRow(row));

  rowsToRender.forEach(({ row, rowIndex }) => {
    const tr = document.createElement('tr');
    const indexCell = document.createElement('td');
    indexCell.className = 'daily-row-index';
    indexCell.textContent = String(rowIndex + 1);

    const rowActions = document.createElement('div');
    rowActions.className = 'daily-row-actions';

    const insertButton = document.createElement('button');
    insertButton.type = 'button';
    insertButton.className = 'daily-row-action daily-row-action--insert';
    insertButton.textContent = '+';
    insertButton.title = 'Inserir linha abaixo';
    insertButton.setAttribute('aria-label', 'Inserir linha abaixo');
    insertButton.addEventListener('click', () => insertDailyRowAfter(rowIndex));
    rowActions.appendChild(insertButton);

    const removeButton = document.createElement('button');
    removeButton.type = 'button';
    removeButton.className = 'daily-row-action daily-row-action--remove';
    removeButton.textContent = '×';
    removeButton.title = 'Excluir linha';
    removeButton.setAttribute('aria-label', 'Excluir linha');
    removeButton.addEventListener('click', () => removeDailyRow(rowIndex));
    rowActions.appendChild(removeButton);

    indexCell.appendChild(rowActions);
    tr.appendChild(indexCell);

    dailyColumns.forEach((column) => {
      const td = document.createElement('td');
      td.contentEditable = 'true';
      td.spellcheck = false;
      let cellValue = row[column.key] || '';
      if (column.key === 'prioridade') cellValue = normalizePrioridade(cellValue);
      if (column.key === 'entrada') {
        cellValue = normalizeDateCellValue(cellValue || row.prazo);
      } else if (column.key === 'prazo') {
        cellValue = normalizeDateCellValue(cellValue);
      }
      td.textContent = cellValue;
      td.dataset.row = String(rowIndex);
      td.dataset.column = column.key;

      if (column.key === 'status') {
        const statusClass = getStatusClass(row[column.key]);
        if (statusClass) {
          td.classList.add(statusClass);
        }
      }

      td.addEventListener('input', () => {
        const text = td.textContent || '';
        updateDailyCell(rowIndex, column.key, text);
        const suggestions = getDailyColumnSuggestions(column.key, text);
        showDailyAutocomplete(td, suggestions);
        if (column.key === 'status') {
          td.classList.remove('daily-status--done', 'daily-status--progress', 'daily-status--blocked', 'daily-status--not-started');
          const nextClass = getStatusClass(text);
          if (nextClass) td.classList.add(nextClass);
        }
      });
      td.addEventListener('keydown', (event) => {
        handleDailyAutocompleteKey(event, td);
      });
      td.addEventListener('focus', () => {
        const text = td.textContent || '';
        const suggestions = getDailyColumnSuggestions(column.key, text);
        showDailyAutocomplete(td, suggestions);
      });
      td.addEventListener('blur', () => {
        hideDailyAutocomplete();
        saveDailyRows();
        ensureEmptyTrailingRow(rowIndex);
      });

      tr.appendChild(td);
    });

    dailyGridBody.appendChild(tr);
  });
}

function ensureEmptyTrailingRow(rowIndex, shouldRender = true) {
  if (rowIndex !== dailyRows.length - 1) return;
  if (isDailyRowCompletelyEmpty(dailyRows[rowIndex])) return;
  const newRow = createEmptyDailyRow();
  const monthValue = dailyGridFilterMonth.value;
  if (monthValue) {
    const [year, month] = monthValue.split('-');
    newRow.entrada = `01/${month}/${year}`;
  }
  dailyRows.push(newRow);
  saveDailyRows();
  if (shouldRender) renderDailyGrid();
}

function addDailyRow() {
  const newRow = createEmptyDailyRow();
  const monthValue = dailyGridFilterMonth.value;
  if (monthValue) {
    const [year, month] = monthValue.split('-');
    newRow.entrada = `01/${month}/${year}`;
  }
  dailyRows.push(newRow);
  saveDailyRows();
  renderDailyGrid();
}

function removeDailyRow(rowIndex) {
  const currentSystem = dailySystems.find((item) => item.key === activeDailySystem);
  const confirmed = window.confirm(`Deseja excluir a linha ${rowIndex + 1} do sistema ${currentSystem?.label || activeDailySystem}?`);
  if (!confirmed) return;
  dailyRows.splice(rowIndex, 1);
  ensureDailyMinimumRows();
  saveDailyRows();
  renderDailyGrid();
}

function insertDailyRowAfter(rowIndex) {
  const newRow = createEmptyDailyRow();
  const monthValue = dailyGridFilterMonth.value;
  if (monthValue) {
    const [year, month] = monthValue.split('-');
    newRow.entrada = `01/${month}/${year}`;
  }
  dailyRows.splice(rowIndex + 1, 0, newRow);
  saveDailyRows();
  renderDailyGrid();
}

function resetDailyGrid() {
  const currentSystem = dailySystems.find((item) => item.key === activeDailySystem);
  const confirmed = window.confirm(`Deseja limpar toda a planilha da daily (${currentSystem?.label || activeDailySystem})?`);
  if (!confirmed) return;

  dailyRows = new Array(DAILY_MIN_ROWS).fill(null).map(() => createEmptyDailyRow());
  saveDailyRows();
  renderDailyGrid();
}

function setDailyGridFilterMode(mode) {
  activeDailyFilterMode = mode;
  renderDailyGrid();
}

function ensureHoursLength() {
  profiles.forEach(({ key }) => {
    const current = state.hours[key] || [];
    const next = new Array(state.activities.length).fill(0);
    for (let i = 0; i < Math.min(current.length, next.length); i += 1) {
      next[i] = current[i];
    }
    state.hours[key] = next;
  });
}

function renderPerfilButtons() {
  perfilButtonsContainer.innerHTML = '';
  profiles.forEach((profile) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `pill ${builderState.profile === profile.key ? 'pill--active' : ''}`;
    button.textContent = profile.shortLabel;
    button.addEventListener('click', () => {
      builderState.profile = builderState.profile === profile.key ? null : profile.key;
      renderPerfilButtons();
    });
    perfilButtonsContainer.appendChild(button);
  });
}

function renderHoraButtons() {
  horaButtonsContainer.innerHTML = '';
  hourOptions.forEach((option) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `pill ${builderState.hours === option ? 'pill--active' : ''}`;
    button.textContent = `${option}h`;
    button.addEventListener('click', () => {
      builderState.hours = builderState.hours === option ? null : option;
      horaManualInput.value = builderState.hours === null ? '' : formatHoursValue(builderState.hours);
      renderHoraButtons();
    });
    horaButtonsContainer.appendChild(button);
  });
}

function resetBuilderSelections(options = {}) {
  const { preserveEditing = false } = options;
  builderState.profile = null;
  builderState.hours = null;
  horaManualInput.value = '';
  renderPerfilButtons();
  renderHoraButtons();
  if (!preserveEditing) {
    builderState.editingIndex = null;
    btnAdicionarAtividade.textContent = DEFAULT_ACTIVITY_BUTTON_TEXT;
  }
}

function enterEditingMode(index) {
  const target = state.manualActivities[index];
  if (!target) return;
  builderState.editingIndex = index;
  atividadeDescricaoInput.value = target.description;
  btnAdicionarAtividade.textContent = EDIT_ACTIVITY_BUTTON_TEXT;
  resetBuilderSelections({ preserveEditing: true });
  atividadeDescricaoInput.focus();
}

function handleAddActivity() {
  const description = atividadeDescricaoInput.value.trim();
  if (!description) {
    alert('Digite a descrição da atividade.');
    return;
  }

  const isEditing = builderState.editingIndex !== null;

  if (!isEditing && !builderState.profile) {
    alert('Selecione o perfil responsável.');
    return;
  }
  if (!isEditing && !builderState.hours) {
    alert('Selecione a quantidade de horas.');
    return;
  }

  let target;
  if (isEditing) {
    target = state.manualActivities[builderState.editingIndex];
    if (!target) {
      resetBuilderSelections();
      atividadeDescricaoInput.value = '';
      return;
    }
    target.description = description;
    target.normalized = normalizeDescription(description);
    if (builderState.profile && builderState.hours) {
      target.hours[builderState.profile] += builderState.hours;
    }
  } else {
    const normalized = normalizeDescription(description);
    target = state.manualActivities.find((item) => item.normalized === normalized);
    if (!target) {
      target = {
        description,
        normalized,
        hours: createEmptyHoursMap()
      };
      state.manualActivities.push(target);
    } else {
      target.description = description;
      target.normalized = normalized;
    }

    target.hours[builderState.profile] += builderState.hours;
  }

  rebuildActivitiesList();
  renderActivities();
  refreshSummaries();

  atividadeDescricaoInput.value = '';
  resetBuilderSelections();
}

function removeManualActivity(index) {
  state.manualActivities.splice(index, 1);
  rebuildActivitiesList();
  renderActivities();
  refreshSummaries();

  if (builderState.editingIndex !== null) {
    atividadeDescricaoInput.value = '';
    resetBuilderSelections();
  }
}

function renderActivities() {
  atividadesContainer.innerHTML = '';
  const fragment = document.createDocumentFragment();

  const manualCount = state.manualActivities.length;

  state.manualActivities.forEach((item, manualIndex) => {
    const card = document.createElement('div');
    card.className = 'activity-card activity-card--manual';

    const header = document.createElement('div');
    header.className = 'activity-card__header';

    const title = document.createElement('h3');
    title.textContent = `4.${manualIndex + 1} ${item.description}`;
    header.appendChild(title);

    const actions = document.createElement('div');
    actions.className = 'activity-card__actions';

    const editButton = document.createElement('button');
    editButton.type = 'button';
    editButton.className = 'icon-button icon-button--edit';
    editButton.textContent = ICONS.edit;
    editButton.title = 'Editar atividade';
    editButton.setAttribute('aria-label', 'Editar atividade');
    editButton.addEventListener('click', () => enterEditingMode(manualIndex));
    actions.appendChild(editButton);

    const removeButton = document.createElement('button');
    removeButton.type = 'button';
    removeButton.className = 'icon-button icon-button--delete';
    removeButton.textContent = ICONS.trash;
    removeButton.title = 'Remover atividade';
    removeButton.setAttribute('aria-label', 'Remover atividade');
    removeButton.addEventListener('click', () => removeManualActivity(manualIndex));
    actions.appendChild(removeButton);

    header.appendChild(actions);
    card.appendChild(header);

    const grid = document.createElement('div');
    grid.className = 'hours-grid';
    profiles.forEach((profile) => {
      const field = document.createElement('div');
      field.className = 'hour-field';

      const label = document.createElement('span');
      label.textContent = profile.shortLabel;
      field.appendChild(label);

      const input = document.createElement('input');
      input.type = 'number';
      input.min = '0';
      input.step = '0.5';
      input.value = item.hours[profile.key] || 0;
      input.addEventListener('input', () => {
        const value = Number(input.value) || 0;
        item.hours[profile.key] = value;
        state.hours[profile.key][manualIndex] = value;
        refreshSummaries();
      });

      field.appendChild(input);
      grid.appendChild(field);
    });

    card.appendChild(grid);
    fragment.appendChild(card);
  });

  fixedActivities.forEach((atividade, fixedIndex) => {
    const globalIndex = manualCount + fixedIndex;
    const displayIndex = globalIndex + 1;

    const card = document.createElement('div');
    card.className = 'activity-card activity-card--fixed';

    const title = document.createElement('h3');
    title.textContent = `4.${displayIndex} ${atividade}`;
    card.appendChild(title);

    const grid = document.createElement('div');
    grid.className = 'hours-grid';

    profiles.forEach((profile) => {
      const field = document.createElement('div');
      field.className = 'hour-field';

      const label = document.createElement('span');
      label.textContent = profile.shortLabel;
      field.appendChild(label);

      const input = document.createElement('input');
      input.type = 'number';
      input.min = '0';
      input.step = '0.5';
      input.value = state.hours[profile.key][globalIndex] ?? 0;
      input.addEventListener('input', () => {
        const value = Number(input.value) || 0;
        state.hours[profile.key][globalIndex] = value;
        refreshSummaries();
      });

      field.appendChild(input);
      grid.appendChild(field);
    });

    card.appendChild(grid);
    fragment.appendChild(card);
  });

  atividadesContainer.appendChild(fragment);
}

function getFactor() {
  const raw = Number(document.querySelector('#gordura').value);
  return Number.isFinite(raw) && raw > 0 ? raw : 1.1;
}

function getTotals() {
  const fator = getFactor();
  const totals = {};

  profiles.forEach(({ key }) => {
    totals[key] = state.hours[key]
      .map((hora) => Math.ceil((Number(hora) || 0) * fator))
      .reduce((acc, value) => acc + value, 0);
  });

  const totalGeral = Object.values(totals).reduce((acc, value) => acc + value, 0);
  return { totals, totalGeral };
}

function formatCurrency(value) {
  return value.toLocaleString('pt-BR', {
    style: 'currency',
    currency: 'BRL',
    minimumFractionDigits: 2
  });
}

function refreshSummaries() {
  const { totals, totalGeral } = getTotals();

  summaryCards.innerHTML = profiles
    .map((profile) => {
      return `
        <div class="summary-card">
          <span>${profile.shortLabel}</span>
          <strong>${totals[profile.key]}h</strong>
          <small>${formatCurrency(totals[profile.key] * profile.rate)}</small>
        </div>
      `;
    })
    .concat([
      `
        <div class="summary-card" style="background:var(--primary); color:#fff;">
          <span>Total Geral</span>
          <strong>${totalGeral}h</strong>
          <small>${formatCurrency(totalFromRates(totals))}</small>
        </div>
      `
    ])
    .join('');

  renderFinanceTable(totals);
  renderProfileBoards(totals);
  renderPreview(totals, totalGeral);
}

function totalFromRates(totals) {
  return profiles.reduce((acc, profile) => acc + totals[profile.key] * profile.rate, 0);
}

function renderFinanceTable(totals) {
  const enumeratedActivities = state.activities.map((activity, index) => ({
    label: `4.${index + 1} ${activity}`,
    index
  }));

  financeTableBody.innerHTML = '';
  let hasHours = false;

  profiles.forEach((profile) => {
    const matchingActivities = enumeratedActivities.filter(({ index }) =>
      (state.hours[profile.key][index] || 0) > 0
    );

    const row = document.createElement('tr');

    const atividadesCell = matchingActivities.length
      ? matchingActivities.map((item) => item.label).join('\n')
      : 'Nenhuma atividade alocada.';

    row.innerHTML = `
      <td>${profile.label}</td>
      <td style="white-space:pre-wrap">${atividadesCell}</td>
      <td>${formatCurrency(profile.rate)}</td>
      <td>${totals[profile.key]}h</td>
      <td>${formatCurrency(totals[profile.key] * profile.rate)}</td>
    `;

    financeTableBody.appendChild(row);

    if (totals[profile.key] > 0) {
      hasHours = true;
    }
  });

  if (hasHours) {
    const footer = document.createElement('tr');
    footer.innerHTML = `
      <td colspan="3" style="text-align:right; font-weight:600;">TOTAL</td>
      <td>${Object.values(totals).reduce((acc, value) => acc + value, 0)}h</td>
      <td>${formatCurrency(totalFromRates(totals))}</td>
    `;
    financeTableBody.appendChild(footer);
  }

  financeSection.style.display = hasHours ? 'flex' : 'none';
}

function renderProfileBoards(totals) {
  const enumeratedActivities = state.activities.map((activity, index) => ({
    label: `4.${index + 1} ${activity}`,
    index
  }));

  profileBoards.innerHTML = profiles
    .map((profile) => {
      const entries = enumeratedActivities
        .filter(({ index }) => (state.hours[profile.key][index] || 0) > 0)
        .map(({ label }) => `<li>${label}</li>`)
        .join('');

      return `
        <div class="summary-card" style="align-items:flex-start;">
          <span>${profile.label}</span>
          <strong>${totals[profile.key]}h</strong>
          <ul style="padding-left:20px; margin:10px 0 0; color:var(--text-light);">
            ${entries || '<li>Nenhuma atividade</li>'}
          </ul>
        </div>
      `;
    })
    .join('');
}

function renderPreview(totals, totalGeral) {
  const fator = getFactor();
  const numeroTicket = document.querySelector('#numeroTicket').value || '000000';
  const solicitacao = document.querySelector('#solicitacaoCliente').value || '—';
  const funcionalidades = document.querySelector('#funcionalidadesAfetadas').value || '—';
  const outras = document.querySelector('#outrasInformacoes').value || 'Não se aplica.';

  const atividadesTexto = state.activities
    .map((atividade, index) => `4.${index + 1} ${atividade}`)
    .join('\n');

  const resumoPerfil = profiles
    .map((profile) => `• ${profile.shortLabel}: ${totals[profile.key]}h (${formatCurrency(totals[profile.key] * profile.rate)})`)
    .join('\n');

  previewOutput.textContent = `LEVANTAMENTO DE ESTIMATIVA - TICKET ${numeroTicket}

1. Solicitação
${solicitacao}

2. Funcionalidades afetadas
${funcionalidades}

3. Outras informações
${outras}

4. Atividades
${atividadesTexto}

5. Totais por perfil
${resumoPerfil}

TOTAL GERAL: ${totalGeral}h (${formatCurrency(totalFromRates(totals))})
Fator de gordura aplicado: ${fator}`;
}

function validateBeforeGenerate() {
  const numeroTicket = document.querySelector('#numeroTicket').value.trim();
  if (!numeroTicket) {
    alert('Informe o número do ticket antes de gerar.');
    return false;
  }

  if (!state.activities.length) {
    alert('Adicione pelo menos uma atividade manual ou utilize as fixas.');
    return false;
  }

  return true;
}


function collectDocumentData() {
  const fator = getFactor();
  const numeroTicket = document.querySelector('#numeroTicket').value.trim() || '000000';
  const solicitacao = document.querySelector('#solicitacaoCliente').value.trim() || '—';
  const funcionalidades = document.querySelector('#funcionalidadesAfetadas').value.trim() || '—';
  const outras = document.querySelector('#outrasInformacoes').value.trim() || 'Não se aplica.';
  const enumeratedActivities = state.activities.map((atividade, index) => `4.${index + 1} ${atividade}`);
  const { totals } = getTotals();
  const totalHoras = Object.values(totals).reduce((acc, value) => acc + value, 0);
  const totalFinanceiro = totalFromRates(totals);

  const perActivityHours = state.activities.map((_, index) => {
    const snapshot = {};
    profiles.forEach(({ key }) => {
      snapshot[key] = Math.ceil((state.hours[key][index] || 0) * fator);
    });
    return snapshot;
  });

  const activitiesByProfile = {};
  profiles.forEach(({ key }) => {
    const filtered = enumeratedActivities.filter((_, index) => perActivityHours[index][key] > 0);
    activitiesByProfile[key] = filtered.join('\n') || 'Nenhuma atividade alocada.';
  });

  return {
    numeroTicket,
    solicitacao,
    funcionalidades,
    outras,
    enumeratedActivities,
    totals,
    totalHoras,
    totalFinanceiro,
    fator,
    activitiesByProfile,
    perActivityHours
  };
}

function buildPlaceholderMap(documentData) {
  const replacements = {
    TICKET1: `TICKET ${documentData.numeroTicket}`,
    SOLICITACAO2: documentData.solicitacao,
    FUNCIONALIDADES3: documentData.funcionalidades,
    ATIVIDADES4: documentData.enumeratedActivities.join('\n'),
    OUTRAS5: documentData.outras,
    TOTAL_GERAL: documentData.totalHoras.toString(),
    VALOR_TOTAL: formatCurrency(documentData.totalFinanceiro),
    FATOR_GORDURA: documentData.fator.toString()
  };

  profiles.forEach((profile) => {
    const upperKey = profile.key.toUpperCase();
    const horasTotais = documentData.totals[profile.key] || 0;
    replacements[`TOTAL_${upperKey}`] = horasTotais.toString();
    replacements[`VALOR_${upperKey}`] = formatCurrency(horasTotais * profile.rate);
    replacements[`VALOR_HORA_${upperKey}`] = formatCurrency(profile.rate);
    replacements[`ATIVIDADES_${upperKey}`] = documentData.activitiesByProfile[profile.key];
  });

  return replacements;
}

function normalizePlaceholderValue(value, placeholder) {
  const multiline = multilinePlaceholders.has(placeholder);
  const base = value == null ? '' : String(value);
  if (!multiline) {
    return base.replace(/\r\n/g, ' ').replace(/\r/g, ' ').replace(/\n/g, ' ');
  }
  return base.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function collectTextSegments(xmlDoc) {
  const walker = xmlDoc.createTreeWalker(xmlDoc.documentElement, NodeFilter.SHOW_TEXT, null);
  const segments = [];
  let combined = '';
  let node = walker.nextNode();
  while (node) {
    const text = node.nodeValue || '';
    const start = combined.length;
    combined += text;
    segments.push({ node, start, end: start + text.length });
    node = walker.nextNode();
  }
  return { segments, combined };
}

function replacePlaceholderInXmlDoc(xmlDoc, placeholder, value, referenceTypography) {
  if (!placeholder) return;
  const affectedParagraphs = new Set();

  while (true) {
    const { segments, combined } = collectTextSegments(xmlDoc);
    const index = combined.indexOf(placeholder);
    if (index === -1) {
      break;
    }

    const endIndex = index + placeholder.length;
    const targetSegments = segments.filter(({ start, end }) => end > index && start < endIndex);
    if (!targetSegments.length) {
      break;
    }

    const first = targetSegments[0];
    const last = targetSegments[targetSegments.length - 1];
    const prefix = first.node.nodeValue.slice(0, index - first.start);
    const suffix = last.node.nodeValue.slice(endIndex - last.start);
    first.node.nodeValue = `${prefix}${value}${suffix}`;

    targetSegments.slice(1).forEach((segment) => {
      segment.node.nodeValue = '';
    });

    const paragraphNode = first.node?.parentNode?.parentNode?.parentNode;
    if (paragraphNode && paragraphNode.nodeName === 'w:p') {
      affectedParagraphs.add(paragraphNode);
    }
  }

  const formatting = paragraphFormattingRules[placeholder];
  if (formatting && affectedParagraphs.size) {
    affectedParagraphs.forEach((paragraph) =>
      applyParagraphFormatting(paragraph, xmlDoc, formatting, referenceTypography)
    );
  }
}

function applyParagraphFormatting(paragraph, xmlDoc, { alignLeft, clearColor, normalizeTypography }, referenceTypography = {}) {
  if (!paragraph || paragraph.nodeName !== 'w:p') return;

  if (alignLeft) {
    let pPr = Array.from(paragraph.childNodes).find((child) => child.nodeName === 'w:pPr');
    if (!pPr) {
      pPr = xmlDoc.createElement('w:pPr');
      paragraph.insertBefore(pPr, paragraph.firstChild);
    }
    let jc = Array.from(pPr.childNodes).find((child) => child.nodeName === 'w:jc');
    if (!jc) {
      jc = xmlDoc.createElement('w:jc');
      pPr.appendChild(jc);
    }
    jc.setAttribute('w:val', 'left');
  }

  if (normalizeTypography && referenceTypography.paragraphStyleId) {
    applyParagraphStyle(paragraph, xmlDoc, referenceTypography.paragraphStyleId);
  }

  if (clearColor || normalizeTypography) {
    const runs = paragraph.getElementsByTagName('w:r');
    Array.from(runs).forEach((run) => {
      if (normalizeTypography) {
        const runStyleTemplate = referenceTypography.runProperties;
        const rPr = ensureRunProperties(run, xmlDoc);
        if (runStyleTemplate) {
          while (rPr.firstChild) {
            rPr.removeChild(rPr.firstChild);
          }
          Array.from(runStyleTemplate.childNodes).forEach((child) => {
            rPr.appendChild(child.cloneNode(true));
          });
        } else {
          setRunFontStyle(run, xmlDoc);
        }
      }

      if (clearColor) {
        const rPr = Array.from(run.childNodes).find((child) => child.nodeName === 'w:rPr');
        if (rPr) {
          const colorNodes = Array.from(rPr.childNodes).filter((child) => child.nodeName === 'w:color');
          colorNodes.forEach((colorNode) => rPr.removeChild(colorNode));
        }
      }
    });
  }
}

function applyReplacementsWithDom(xml, replacements, documentData) {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xml, 'application/xml');
  if (xmlDoc.getElementsByTagName('parsererror').length) {
    throw new Error('Falha ao interpretar o modelo Word.');
  }

  const referenceTypography = extractReferenceTypography(xmlDoc);

  Object.entries(replacements).forEach(([placeholder, value]) => {
    const normalized = normalizePlaceholderValue(value, placeholder);
    replacePlaceholderInXmlDoc(xmlDoc, placeholder, normalized, referenceTypography);
  });

  convertNewlinesToWordBreaks(xmlDoc);
  insertEstimationTables(xmlDoc, documentData);

  const serializer = new XMLSerializer();
  return serializer.serializeToString(xmlDoc);
}

function convertNewlinesToWordBreaks(xmlDoc) {
  const walker = xmlDoc.createTreeWalker(xmlDoc.documentElement, NodeFilter.SHOW_TEXT, null);
  const targets = [];
  let node;
  while ((node = walker.nextNode())) {
    if (node.nodeValue && node.nodeValue.includes('\n')) {
      targets.push(node);
    }
  }

  targets.forEach((textNode) => {
    const textParent = textNode.parentNode; // w:t
    if (!textParent || textParent.nodeName !== 'w:t') {
      textNode.nodeValue = textNode.nodeValue.replace(/\n/g, ' ');
      return;
    }
    const runNode = textParent.parentNode; // w:r
    if (!runNode || runNode.nodeName !== 'w:r') {
      textNode.nodeValue = textNode.nodeValue.replace(/\n/g, ' ');
      return;
    }

    const lines = textNode.nodeValue.split('\n');
    textNode.nodeValue = lines.shift();
    const referenceNode = textParent.nextSibling;

    lines.forEach((line) => {
      const br = xmlDoc.createElement('w:br');
      runNode.insertBefore(br, referenceNode);
      const newTextElement = textParent.cloneNode(false);
      newTextElement.textContent = line;
      runNode.insertBefore(newTextElement, referenceNode);
    });
  });
}

async function generateWordDocument() {
  const data = collectDocumentData();
  const buffer = await loadTemplateBuffer();
  const workingBuffer = buffer.slice(0);
  const zip = await JSZip.loadAsync(workingBuffer);
  let documentXml = await zip.file('word/document.xml').async('string');
  const replacements = buildPlaceholderMap(data);
  documentXml = applyReplacementsWithDom(documentXml, replacements, data);
  zip.file('word/document.xml', documentXml);
  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });
  saveAs(blob, `Levantamento - Ticket ${data.numeroTicket}.docx`);
}

async function handleGenerate() {
  if (!validateBeforeGenerate()) return;

  try {
    await generateWordDocument();
  } catch (error) {
    console.error('Erro ao gerar documento Word:', error);
    alert('Não foi possível gerar o documento Word. Verifique o console para mais detalhes.');
  }
}

// Event bindings
btnAdicionarAtividade.addEventListener('click', handleAddActivity);
document.querySelector('#gordura').addEventListener('input', refreshSummaries);
document.querySelector('#btnGerar').addEventListener('click', handleGenerate);
horaManualInput.addEventListener('input', () => {
  builderState.hours = parseHoursValue(horaManualInput.value);
  renderHoraButtons();
});
horaManualInput.addEventListener('blur', () => {
  const parsed = parseHoursValue(horaManualInput.value);
  builderState.hours = parsed;
  horaManualInput.value = parsed === null ? '' : formatHoursValue(parsed);
  renderHoraButtons();
});
moduleButtons.forEach((button) => {
  button.addEventListener('click', () => {
    setActiveModule(button.dataset.module);
  });
});
dailySystemButtons.forEach((button) => {
  button.addEventListener('click', () => {
    setActiveDailySystem(button.dataset.dailySystem);
  });
});
btnDailyAddRow.addEventListener('click', addDailyRow);
btnDailySaveAll.addEventListener('click', () => {
  saveDailyRows();
});
btnDailyReset.addEventListener('click', resetDailyGrid);
btnDailyImportXlsx.addEventListener('click', () => {
  dailyXlsxInput.click();
});
dailyXlsxInput.addEventListener('change', async () => {
  const file = dailyXlsxInput.files?.[0];
  if (!file) return;

  try {
    await importDailyFromWorkbook(file);
  } catch (error) {
    console.error('Falha ao importar XLSX da Daily:', error);
    alert('Não foi possível importar a planilha. Verifique o formato do arquivo.');
  } finally {
    dailyXlsxInput.value = '';
  }
});
btnDailyExportXlsx.addEventListener('click', async () => {
  try {
    await exportDailyToWorkbook();
  } catch (error) {
    console.error('Falha ao exportar XLSX da Daily:', error);
    alert('Não foi possível exportar a planilha.');
  }
});
dailyGridFilterMode.addEventListener('change', () => {
  setDailyGridFilterMode(dailyGridFilterMode.value);
});
dailyGridFilterMonth.addEventListener('change', () => {
  if (activeDailyFilterMode === 'month') {
    renderDailyGrid();
  }
});
dailyStatusFilter.addEventListener('change', () => {
  renderDailyGrid();
});
dailyGridSearch.addEventListener('input', () => {
  renderDailyGrid();
});
btnDailyGenerateReport.addEventListener('click', renderDailyMonthlyReport);
btnDailyPrintReport.addEventListener('click', printDailyMonthlyReport);
['solicitacaoCliente', 'funcionalidadesAfetadas', 'outrasInformacoes'].forEach((id) => {
  document.querySelector(`#${id}`).addEventListener('input', () => refreshSummaries());
});

// Inicialização
setActiveModule(activeModule);
dailyGridFilterMonth.value = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`;
dailyReportMonthInput.value = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`;
setActiveDailySystem(activeDailySystem);
renderPerfilButtons();
renderHoraButtons();
rebuildActivitiesList();
renderActivities();
refreshSummaries();
loadTemplateBuffer().catch((error) => {
  console.error('Não foi possível pré-carregar o modelo oficial:', error);
});
