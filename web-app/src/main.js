import './style.css';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';

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
  SOLICITACAO2: { alignLeft: true, clearColor: true },
  FUNCIONALIDADES3: { alignLeft: true, clearColor: true },
  ATIVIDADES4: { alignLeft: true, clearColor: true },
  OUTRAS5: { alignLeft: true, clearColor: true }
};

const WORD_MAIN_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const TOTAL_ROW_SHADING = 'c9ab83';

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
    if (bold) {
      const rPr = xmlDoc.createElementNS(WORD_MAIN_NS, 'w:rPr');
      rPr.appendChild(xmlDoc.createElementNS(WORD_MAIN_NS, 'w:b'));
      run.appendChild(rPr);
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

const hourOptions = [1, 2, 3, 4];

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
    <main class="content">
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
    </main>
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
const btnAdicionarAtividade = document.querySelector('#btnAdicionarAtividade');

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
      renderHoraButtons();
    });
    horaButtonsContainer.appendChild(button);
  });
}

function resetBuilderSelections(options = {}) {
  const { preserveEditing = false } = options;
  builderState.profile = null;
  builderState.hours = null;
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

    const chips = document.createElement('div');
    chips.className = 'activity-card__chips';
    let hasHours = false;
    profiles.forEach((profile) => {
      const hours = item.hours[profile.key] || 0;
      if (hours > 0) {
        hasHours = true;
        const chip = document.createElement('span');
        chip.className = 'pill pill--tag';
        chip.textContent = `${profile.shortLabel}: ${hours}h`;
        chips.appendChild(chip);
      }
    });

    if (hasHours) {
      card.appendChild(chips);
    } else {
      const empty = document.createElement('p');
      empty.className = 'activity-card__empty';
      empty.textContent = 'Nenhuma hora atribuída ainda. Use o formulário acima para inserir.';
      card.appendChild(empty);
    }

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

function replacePlaceholderInXmlDoc(xmlDoc, placeholder, value) {
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
    affectedParagraphs.forEach((paragraph) => applyParagraphFormatting(paragraph, xmlDoc, formatting));
  }
}

function applyParagraphFormatting(paragraph, xmlDoc, { alignLeft, clearColor }) {
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

  if (clearColor) {
    const runs = paragraph.getElementsByTagName('w:r');
    Array.from(runs).forEach((run) => {
      const rPr = Array.from(run.childNodes).find((child) => child.nodeName === 'w:rPr');
      if (rPr) {
        const colorNodes = Array.from(rPr.childNodes).filter((child) => child.nodeName === 'w:color');
        colorNodes.forEach((colorNode) => rPr.removeChild(colorNode));
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

  Object.entries(replacements).forEach(([placeholder, value]) => {
    const normalized = normalizePlaceholderValue(value, placeholder);
    replacePlaceholderInXmlDoc(xmlDoc, placeholder, normalized);
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
['solicitacaoCliente', 'funcionalidadesAfetadas', 'outrasInformacoes'].forEach((id) => {
  document.querySelector(`#${id}`).addEventListener('input', () => refreshSummaries());
});

// Inicialização
renderPerfilButtons();
renderHoraButtons();
rebuildActivitiesList();
renderActivities();
refreshSummaries();
loadTemplateBuffer().catch((error) => {
  console.error('Não foi possível pré-carregar o modelo oficial:', error);
});
