import {
  initializeConnectionModal,
  requireRole,
  getCurrentUser,
  setCurrentUser,
  getSupabaseClient,
  onSupabaseReady
} from './supabaseClient.js';

const PLANNING_TOURS = [
  { id: 1 },
  { id: 2 },
  { id: 3 },
  { id: 4 },
  { id: 5 },
  { id: 6 }
];

const PLANNING_COLUMNS_TABLE = 'planning_columns';

const STATUS_TABS = [
  { id: 'pending', label: 'En attente', statuses: ['en attente'] },
  { id: 'accepted', label: 'Acceptées', statuses: ['validé'] },
  { id: 'refused', label: 'Refusées', statuses: ['refusé'] }
];

const STATUS_LABELS = new Map([
  [
    'en attente',
    {
      label: 'En attente',
      className: 'badge badge-warning'
    }
  ],
  [
    'validé',
    {
      label: 'Acceptée',
      className: 'badge badge-success'
    }
  ],
  [
    'refusé',
    {
      label: 'Refusée',
      className: 'badge badge-danger'
    }
  ]
]);

const QUALITY_LABELS = new Map([
  ['normale', 'Normale'],
  ['bonne', 'Bonne']
]);

const TYPE_LABELS = new Map([
  ['visite', 'Visite'],
  ['consultation', 'Consultation'],
  ['téléconsultation', 'Téléconsultation']
]);

const ADMIN_SETTINGS_TABLE = 'parametres_administratifs';
const AUDIT_TABLE = 'planning_choice_audit';
const CHOICES_TABLE = 'planning_choices';
const AUTO_ASSIGNMENT_WORK_TABLE = 'auto_assignment_work_queue';
const AUTO_ASSIGNMENT_RUNS_TABLE = 'auto_assignment_runs';
const AUTO_ASSIGNMENT_RUN_ENTRIES_TABLE = 'auto_assignment_run_entries';

const AUTO_ASSIGNMENT_ALGORITHMS = new Map([
  [
    'simple',
    {
      id: 'simple',
      label: 'Algorithme simple'
    }
  ]
]);

const AUTO_ASSIGNMENT_DEFAULTS = {
  populations: ['medecin', 'remplacant'],
  order: 'asc',
  startTrigram: '',
  algorithm: 'simple',
  rotations: 1
};

const DATE_FORMAT = new Intl.DateTimeFormat('fr-FR', {
  year: 'numeric',
  month: '2-digit',
  day: '2-digit'
});

const DATETIME_FORMAT = new Intl.DateTimeFormat('fr-FR', {
  year: 'numeric',
  month: '2-digit',
  day: '2-digit',
  hour: '2-digit',
  minute: '2-digit'
});

const USER_TYPE_FILTERS = new Set(['medecin', 'remplacant']);

const state = {
  supabase: null,
  actor: null,
  planningReference: null,
  activeTourId: PLANNING_TOURS[0].id,
  planningYear: null,
  planningMonthOne: null,
  planningMonthTwo: null,
  columns: new Map(),
  requests: [],
  activeTab: STATUS_TABS[0].id,
  isLoading: false,
  filters: {
    date: '',
    type: '',
    doctor: '',
    column: '',
    status: '',
    userType: ''
  },
  viewMode: 'requests',
  autoAssignment: {
    isRunning: false,
    directory: [],
    lastResult: null,
    lastRunId: null
  }
};

const elements = {
  logoutBtn: document.querySelector('#logout'),
  disconnectBtn: document.querySelector('#disconnect'),
  backBtn: document.querySelector('#back-to-admin'),
  modeTabs: document.querySelector('#request-mode-tabs'),
  userTypeTabsNav: document.querySelector('#request-user-tabs'),
  tabsNav: document.querySelector('#request-tabs'),
  filtersForm: document.querySelector('#request-filters'),
  feedback: document.querySelector('#request-feedback'),
  tableBody: document.querySelector('#requests-body'),
  userTypeTabButtons: Array.from(document.querySelectorAll('[data-user-type-tab]')),
  requestPanels: Array.from(document.querySelectorAll('[data-mode-panel="requests"]')),
  autoPanel: document.querySelector('#auto-assignment-panel'),
  autoForm: document.querySelector('#auto-assignment-form'),
  autoFeedback: document.querySelector('#auto-assignment-feedback'),
  autoActiveTour: document.querySelector('#auto-assignment-active-tour'),
  autoResultsSection: document.querySelector('#auto-assignment-results'),
  autoResultsBody: document.querySelector('#auto-assignment-results-body'),
  autoSummary: document.querySelector('#auto-assignment-summary'),
  autoTrigramOptions: document.querySelector('#auto-assignment-trigram-options'),
  autoStartInput: document.querySelector('#auto-assignment-start'),
  autoUndoButton: document.querySelector('[data-auto-action="undo"]')
};

const normalizeUserTypeFilter = (value) => {
  if (typeof value !== 'string') {
    return '';
  }
  const lowered = value
    .toString()
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '');
  return USER_TYPE_FILTERS.has(lowered) ? lowered : '';
};

const normalizeTrigram = (value) => {
  if (typeof value !== 'string') {
    return '';
  }
  return value
    .toString()
    .trim()
    .slice(0, 3)
    .toUpperCase();
};

const sanitizeActiveTour = (value) => {
  if (value == null || value === '') {
    return null;
  }
  const numeric = Number(value);
  if (!Number.isInteger(numeric)) {
    return null;
  }
  if (numeric < 1 || numeric > PLANNING_TOURS.length) {
    return null;
  }
  return numeric;
};

const getTourConfig = (tourId = state.activeTourId) =>
  PLANNING_TOURS.find((tour) => tour.id === tourId) ?? PLANNING_TOURS[0];

const parseNumeric = (value) => {
  if (value == null || value === '') {
    return null;
  }
  const parsed = Number.parseFloat(value);
  return Number.isNaN(parsed) ? null : parsed;
};

const haveMatchingChoiceIndex = (candidate, selectedIndex) => {
  if (selectedIndex == null) {
    return false;
  }
  const candidateIndex = parseNumeric(candidate.choice_index ?? candidate.choiceIndex);
  return candidateIndex != null && candidateIndex === selectedIndex;
};

const shouldTreatAsAlternative = (candidate, selectedRank, selectedIndex) => {
  if (haveMatchingChoiceIndex(candidate, selectedIndex)) {
    return true;
  }
  const candidateRank = parseNumeric(candidate.choice_rank ?? candidate.choiceRank);
  if (candidateRank == null) {
    return true;
  }
  const threshold = selectedRank ?? 1;
  return candidateRank > threshold;
};

const collectAlternativeIds = (competing, request) => {
  const selectedRank = parseNumeric(request.choiceRank ?? request.choice_rank);
  const selectedIndex = parseNumeric(request.choiceIndex ?? request.choice_index);
  const alternatives = competing
    .filter((item) => item.trigram === request.trigram && item.id !== request.id)
    .filter((item) => shouldTreatAsAlternative(item, selectedRank, selectedIndex))
    .map((item) => item.id);
  return Array.from(new Set(alternatives));
};

const toMonthPart = (month) => String(month + 1).padStart(2, '0');

const getPlanningReference = ({ tourId, year, monthOne, monthTwo }) =>
  `tour${tourId}-${year}-${toMonthPart(monthOne)}-${toMonthPart(monthTwo)}`;

const ensureSupabase = async () => {
  if (state.supabase) {
    return state.supabase;
  }
  await onSupabaseReady();
  const client = getSupabaseClient();
  state.supabase = client;
  return client;
};

const setFeedback = (message) => {
  if (elements.feedback) {
    elements.feedback.textContent = message ?? '';
  }
};

const setAutoAssignmentFeedback = (message) => {
  if (elements.autoFeedback) {
    elements.autoFeedback.textContent = message ?? '';
  }
};

const setAutoAssignmentRunning = (isRunning) => {
  state.autoAssignment.isRunning = isRunning;
  const form = elements.autoForm;
  if (form) {
    form.classList.toggle('is-loading', isRunning);
    form.querySelectorAll('input, select, button').forEach((control) => {
      if (!control) {
        return;
      }
      if (control.dataset?.autoAction === 'undo' && state.autoAssignment.lastRunId == null) {
        control.disabled = true;
        return;
      }
      control.disabled = isRunning;
    });
  }
  if (elements.autoUndoButton) {
    if (!isRunning) {
      elements.autoUndoButton.disabled = state.autoAssignment.lastRunId == null;
    } else {
      elements.autoUndoButton.disabled = true;
    }
  }
};

const updateModeTabs = () => {
  if (!elements.modeTabs) {
    return;
  }
  const active = state.viewMode;
  elements.modeTabs.querySelectorAll('[data-mode]').forEach((button) => {
    if (!button) {
      return;
    }
    const mode = button.dataset.mode === 'auto' ? 'auto' : 'requests';
    const isActive = mode === active;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-selected', isActive ? 'true' : 'false');
    button.setAttribute('tabindex', isActive ? '0' : '-1');
  });
};

const updateModePanels = () => {
  const isAuto = state.viewMode === 'auto';
  elements.requestPanels.forEach((panel) => {
    if (!panel) {
      return;
    }
    panel.classList.toggle('hidden', isAuto);
  });
  if (elements.autoPanel) {
    elements.autoPanel.classList.toggle('hidden', !isAuto);
  }
};

const parseTimeToMinutes = (value) => {
  if (!value || typeof value !== 'string') {
    return null;
  }
  const [hours, minutes] = value.split(':');
  const h = Number.parseInt(hours, 10);
  const m = Number.parseInt(minutes, 10);
  if (Number.isNaN(h) || Number.isNaN(m)) {
    return null;
  }
  return h * 60 + m;
};

const rangesOverlap = (a, b) => {
  if (!a || !b) {
    return false;
  }
  const startA = parseTimeToMinutes(a.start);
  const endA = parseTimeToMinutes(a.end);
  const startB = parseTimeToMinutes(b.start);
  const endB = parseTimeToMinutes(b.end);
  if (
    startA == null ||
    endA == null ||
    startB == null ||
    endB == null
  ) {
    return false;
  }
  return startA < endB && startB < endA;
};

const getColumnInfo = (columnNumber) => state.columns.get(columnNumber) ?? null;

const getSlotKey = (request) => {
  if (!request) {
    return null;
  }
  const dayKey = getRequestDayKey(request);
  if (!dayKey || request.columnNumber == null) {
    return null;
  }
  return `${dayKey}#${request.columnNumber}`;
};

const getColumnRange = (columnNumber) => {
  const column = getColumnInfo(columnNumber);
  if (!column) {
    return null;
  }
  const start = column.start_time ?? column.startTime ?? null;
  const end = column.end_time ?? column.endTime ?? null;
  if (!start && !end) {
    return null;
  }
  return { start, end };
};

const hasConflictWithAssignedSlots = (trigram, request, assignedMap) => {
  if (!trigram || !request) {
    return false;
  }
  const dayKey = getRequestDayKey(request);
  if (!dayKey) {
    return false;
  }
  const existing = assignedMap.get(trigram) ?? [];
  if (!existing.length) {
    return false;
  }
  const targetRange = getColumnRange(request.columnNumber);
  return existing.some((slot) => {
    if (slot.dayKey !== dayKey) {
      return false;
    }
    if (slot.columnNumber === request.columnNumber) {
      return true;
    }
    if (!targetRange || !slot.range) {
      return false;
    }
    return rangesOverlap(targetRange, slot.range);
  });
};

const formatHours = (columnNumber) => {
  const column = getColumnInfo(columnNumber);
  if (!column) {
    return '—';
  }
  const start = column.start_time ?? column.startTime;
  const end = column.end_time ?? column.endTime;
  if (!start && !end) {
    return '—';
  }
  const formatPart = (value) => {
    if (!value) {
      return null;
    }
    const [hour, minute] = value.split(':');
    if (hour == null || minute == null) {
      return value;
    }
    return `${hour}h${minute}`;
  };
  const formattedStart = formatPart(start);
  const formattedEnd = formatPart(end);
  if (formattedStart && formattedEnd) {
    return `${formattedStart} – ${formattedEnd}`;
  }
  return formattedStart ?? formattedEnd ?? '—';
};

const formatDate = (value) => {
  if (!value) {
    return '—';
  }
  const parsed = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return '—';
  }
  return DATE_FORMAT.format(parsed);
};

const formatDateTime = (value) => {
  if (!value) {
    return '—';
  }
  const parsed = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return '—';
  }
  return DATETIME_FORMAT.format(parsed);
};

const formatPriority = (request) => {
  const index = Number.parseInt(request.choiceIndex, 10);
  const rank = Number.parseInt(request.choiceRank, 10);
  if (!Number.isFinite(index) || index <= 0) {
    return '—';
  }
  if (!Number.isFinite(rank) || rank <= 1) {
    return String(index);
  }
  return `${index}.${rank}`;
};

const computeConsolidatedIndex = (choiceIndex, choiceRank) => {
  const numericIndex = Number.parseInt(choiceIndex, 10);
  const numericRank = Number.parseInt(choiceRank, 10);
  if (!Number.isFinite(numericIndex)) {
    return null;
  }
  if (!Number.isFinite(numericRank) || numericRank <= 1) {
    return String(numericIndex);
  }
  return `${numericIndex}.${numericRank}`;
};

const getRootChoiceIndex = (choiceIndex) => {
  const numericIndex = Number.parseInt(choiceIndex, 10);
  return Number.isFinite(numericIndex) ? numericIndex : null;
};

const getStatusConfig = (status) => STATUS_LABELS.get(status) ?? { label: status ?? '—', className: 'badge' };

const buildTabs = () => {
  if (!elements.tabsNav) {
    return;
  }
  elements.tabsNav.innerHTML = '';
  STATUS_TABS.forEach((tab) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'request-tab';
    button.id = `requests-tab-${tab.id}`;
    button.dataset.tab = tab.id;
    button.setAttribute('role', 'tab');
    button.setAttribute('aria-controls', 'requests-table');
    button.setAttribute('aria-selected', String(state.activeTab === tab.id));
    button.textContent = `${tab.label}`;
    elements.tabsNav.appendChild(button);
  });
};

const updateTabState = () => {
  if (!elements.tabsNav) {
    return;
  }
  const counts = new Map();
  STATUS_TABS.forEach((tab) => {
    const count = getFilteredRequests(tab.statuses).length;
    counts.set(tab.id, count);
  });
  elements.tabsNav.querySelectorAll('[data-tab]').forEach((button) => {
    const tabId = button.dataset.tab;
    const count = counts.get(tabId) ?? 0;
    button.classList.toggle('is-active', state.activeTab === tabId);
    button.setAttribute('aria-selected', String(state.activeTab === tabId));
    button.textContent = `${STATUS_TABS.find((tab) => tab.id === tabId)?.label ?? ''} (${count})`;
  });
};

const getRequestDayKey = (request) => {
  if (!request?.day) {
    return null;
  }
  const parsed = new Date(request.day);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }
  return parsed.toISOString().split('T')[0];
};

const getFilteredRequests = (statuses) => {
  const activeStatuses = Array.isArray(statuses) && statuses.length ? statuses : null;
  const filters = state.filters;
  return state.requests.filter((request) => {
    if (activeStatuses && !activeStatuses.includes(request.status)) {
      return false;
    }
    if (filters.status && filters.status !== request.status) {
      return false;
    }
    if (filters.userType) {
      if (normalizeUserTypeFilter(request.userType) !== filters.userType) {
        return false;
      }
    }
    if (filters.date) {
      const key = getRequestDayKey(request);
      if (!key || key !== filters.date) {
        return false;
      }
    }
    if (filters.type) {
      if ((request.activityType ?? '').toLowerCase() !== filters.type) {
        return false;
      }
    }
    if (filters.doctor) {
      const trigram = (request.trigram ?? '').toLowerCase();
      if (!trigram.includes(filters.doctor.toLowerCase())) {
        return false;
      }
    }
    if (filters.column) {
      const columnLabel = `${request.planningDayLabel ?? ''} ${request.columnLabel ?? ''} ${request.slotTypeCode ?? ''} ${
        request.columnNumber ?? ''
      }`
        .trim()
        .toLowerCase();
      if (!columnLabel.includes(filters.column.toLowerCase())) {
        return false;
      }
    }
    return true;
  });
};

const compareRequestsByPriority = (a, b) => {
  const indexA = Number.isFinite(a?.choiceIndex) ? a.choiceIndex : Number.NEGATIVE_INFINITY;
  const indexB = Number.isFinite(b?.choiceIndex) ? b.choiceIndex : Number.NEGATIVE_INFINITY;
  if (indexA !== indexB) {
    return indexB - indexA;
  }
  const rankA = Number.isFinite(a?.choiceRank) ? a.choiceRank : Number.MAX_SAFE_INTEGER;
  const rankB = Number.isFinite(b?.choiceRank) ? b.choiceRank : Number.MAX_SAFE_INTEGER;
  if (rankA !== rankB) {
    return rankA - rankB;
  }
  const timeA = a?.createdAt?.getTime?.() ?? new Date(a?.createdAt ?? 0).getTime();
  const timeB = b?.createdAt?.getTime?.() ?? new Date(b?.createdAt ?? 0).getTime();
  return timeA - timeB;
};

const sortRequests = (list) => [...list].sort(compareRequestsByPriority);

const sortRequestsByPreference = (list) => [...list].sort(compareRequestsByPriority);

const renderRequests = () => {
  if (!elements.tableBody) {
    return;
  }
  const tab = STATUS_TABS.find((item) => item.id === state.activeTab) ?? STATUS_TABS[0];
  const requests = sortRequests(getFilteredRequests(tab.statuses));
  elements.tableBody.innerHTML = '';
  if (!requests.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 11;
    cell.className = 'requests-empty-cell';
    cell.textContent = 'Aucune demande à afficher.';
    row.appendChild(cell);
    elements.tableBody.appendChild(row);
    updateTabState();
    return;
  }

  requests.forEach((request) => {
    const row = document.createElement('tr');

    const addCell = (content, column) => {
      const cell = document.createElement('td');
      if (column) {
        cell.dataset.column = column;
      }
      if (content instanceof HTMLElement) {
        cell.appendChild(content);
      } else {
        cell.textContent = content ?? '—';
      }
      row.appendChild(cell);
    };

    addCell(formatDate(request.day), 'date');
    const columnLabel =
      request.columnLabel ||
      request.slotTypeCode ||
      (request.columnNumber != null ? `Colonne ${request.columnNumber}` : '—');
    addCell(columnLabel, 'columnLabel');
    addCell(Number.isFinite(request.columnNumber) ? String(request.columnNumber) : '—', 'columnNumber');
    addCell(
      TYPE_LABELS.get(request.activityType) ?? TYPE_LABELS.get((request.activityType ?? '').toLowerCase()) ?? 'Visite',
      'type'
    );
    addCell(formatHours(request.columnNumber), 'hours');
    addCell((request.trigram ?? '').toUpperCase(), 'user');
    addCell(formatPriority(request), 'priority');
    addCell(QUALITY_LABELS.get(request.guardNature) ?? '—', 'quality');

    const statusConfig = getStatusConfig(request.status);
    const statusBadge = document.createElement('span');
    statusBadge.className = statusConfig.className;
    statusBadge.textContent = statusConfig.label;
    addCell(statusBadge, 'status');

    addCell(formatDateTime(request.createdAt), 'created');

    const actionsCell = document.createElement('td');
    actionsCell.dataset.column = 'actions';
    actionsCell.className = 'requests-actions-cell';

    const acceptBtn = document.createElement('button');
    acceptBtn.type = 'button';
    acceptBtn.className = 'subtle-button';
    acceptBtn.dataset.action = 'accept';
    acceptBtn.dataset.requestId = String(request.id);
    acceptBtn.textContent = 'Accepter';
    if (request.status === 'validé') {
      acceptBtn.setAttribute('disabled', 'true');
    }

    const refuseBtn = document.createElement('button');
    refuseBtn.type = 'button';
    refuseBtn.className = 'danger-link';
    refuseBtn.dataset.action = 'refuse';
    refuseBtn.dataset.requestId = String(request.id);
    refuseBtn.textContent = 'Refuser';
    if (request.status === 'refusé') {
      refuseBtn.setAttribute('disabled', 'true');
    }

    actionsCell.appendChild(acceptBtn);
    actionsCell.appendChild(refuseBtn);
    row.appendChild(actionsCell);

    elements.tableBody.appendChild(row);
  });

  updateTabState();
};

const ensureActor = async () => {
  if (state.actor) {
    return state.actor;
  }
  const supabase = await ensureSupabase();
  if (!supabase) {
    return null;
  }
  const currentUser = getCurrentUser();
  if (!currentUser?.id) {
    return null;
  }
  const { data, error } = await supabase
    .from('users')
    .select('id, username, trigram')
    .eq('id', currentUser.id)
    .maybeSingle();
  if (error) {
    console.error(error);
  }
  const actor = {
    id: data?.id ?? currentUser.id ?? null,
    username: data?.username ?? currentUser.username ?? '',
    trigram: (data?.trigram ?? currentUser.trigram ?? currentUser.username ?? '')
      .toString()
      .slice(0, 3)
      .toUpperCase()
  };
  state.actor = actor;
  return actor;
};

const loadAdministrativeSettings = async () => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return;
  }
  const { data, error } = await supabase
    .from(ADMIN_SETTINGS_TABLE)
    .select('active_tour, planning_year, planning_month_one, planning_month_two')
    .order('id', { ascending: true })
    .limit(1);
  if (error) {
    console.error(error);
    return;
  }
  const record = data?.[0] ?? null;
  const sanitizedActiveTour = sanitizeActiveTour(record?.active_tour);
  const now = new Date();
  state.activeTourId = sanitizedActiveTour ?? state.activeTourId ?? PLANNING_TOURS[0].id;
  state.planningYear = Number.isInteger(record?.planning_year) ? record.planning_year : now.getFullYear();
  state.planningMonthOne = Number.isInteger(record?.planning_month_one) ? record.planning_month_one : now.getMonth();
  state.planningMonthTwo = Number.isInteger(record?.planning_month_two)
    ? record.planning_month_two
    : (state.planningMonthOne + 1) % 12;
  state.planningReference = getPlanningReference({
    tourId: state.activeTourId,
    year: state.planningYear,
    monthOne: state.planningMonthOne,
    monthTwo: state.planningMonthTwo
  });
  updateAutoAssignmentActiveTour();
};

const loadPlanningColumns = async () => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return;
  }
  const { data, error } = await supabase
    .from(PLANNING_COLUMNS_TABLE)
    .select('tour_number, position, label, type_category, start_time, end_time')
    .eq('tour_number', state.activeTourId);
  if (error) {
    console.error(error);
    return;
  }
  state.columns = new Map((data ?? []).map((column) => [column.position, column]));
};

const updateAutoAssignmentActiveTour = () => {
  if (!elements.autoActiveTour) {
    return;
  }
  if (!state.activeTourId) {
    elements.autoActiveTour.textContent = '—';
    return;
  }
  const parts = [`Tour ${state.activeTourId}`];
  if (Number.isInteger(state.planningYear)) {
    parts.push(String(state.planningYear));
  }
  elements.autoActiveTour.textContent = parts.join(' • ');
};

const getSelectedAutoPopulations = () => {
  const form = elements.autoForm;
  if (!form) {
    return new Set();
  }
  const inputs = Array.from(form.querySelectorAll('input[name="population"]'));
  const selected = inputs
    .filter((input) => input.checked)
    .map((input) => normalizeUserTypeFilter(input.value));
  return new Set(selected.filter(Boolean));
};

const ensureAutoAssignmentStartTrigram = (options) => {
  const input = elements.autoStartInput;
  if (!input) {
    return;
  }
  const available = new Set(options.map((item) => item.trigram));
  const current = normalizeTrigram(input.value);
  if (available.size === 0) {
    input.value = '';
    return;
  }
  if (!available.has(current)) {
    input.value = options[0].trigram;
  }
};

const updateAutoAssignmentTrigramOptions = () => {
  const datalist = elements.autoTrigramOptions;
  if (!datalist) {
    return;
  }
  datalist.innerHTML = '';
  const populations = getSelectedAutoPopulations();
  const filtered = state.autoAssignment.directory.filter((entry) => {
    if (!entry.trigram) {
      return false;
    }
    if (!populations.size) {
      return true;
    }
    return populations.has(normalizeUserTypeFilter(entry.role));
  });
  filtered.sort((a, b) => a.trigram.localeCompare(b.trigram, 'fr', { sensitivity: 'base' }));
  filtered.forEach((entry) => {
    const option = document.createElement('option');
    option.value = entry.trigram;
    if (entry.username) {
      option.label = `${entry.trigram} — ${entry.username}`;
    }
    datalist.appendChild(option);
  });
  ensureAutoAssignmentStartTrigram(filtered);
};

const loadAutoAssignmentDirectory = async () => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return;
  }
  const { data, error } = await supabase
    .from('users')
    .select('id, username, trigram, role')
    .in('role', ['medecin', 'remplacant'])
    .order('trigram', { ascending: true });
  if (error) {
    console.error(error);
    return;
  }
  state.autoAssignment.directory = (data ?? []).map((record) => ({
    id: record.id ?? null,
    username: record.username ?? '',
    trigram: normalizeTrigram(record.trigram ?? record.username ?? ''),
    role: normalizeUserTypeFilter(record.role ?? '')
  }));
  updateAutoAssignmentTrigramOptions();
};

const mapRequestRecord = (record) => {
  const createdAt = record.created_at ? new Date(record.created_at) : null;
  const columnNumberValue = Number.parseInt(record.column_number ?? record.columnNumber ?? '', 10);
  const choiceIndexValue = Number.parseFloat(record.choice_index ?? record.choiceIndex ?? '');
  const choiceRankValue = Number.parseFloat(record.choice_rank ?? record.choiceRank ?? '');
  const choiceOrderValue = Number.parseInt(record.choice_order ?? record.choiceOrder ?? '', 10);
  return {
    id: record.id,
    trigram: (record.trigram ?? '').toUpperCase(),
    userType: normalizeUserTypeFilter(record.user_type ?? ''),
    day: record.day ?? null,
    columnNumber: Number.isNaN(columnNumberValue) ? null : columnNumberValue,
    columnLabel: record.column_label ?? null,
    planningDayLabel: record.planning_day_label ?? null,
    slotTypeCode: record.slot_type_code ?? null,
    guardNature: record.guard_nature ?? 'normale',
    activityType: (record.activity_type ?? 'visite').toLowerCase(),
    choiceIndex: Number.isNaN(choiceIndexValue) ? null : choiceIndexValue,
    choiceRank: Number.isNaN(choiceRankValue) ? null : choiceRankValue,
    choiceOrder: Number.isNaN(choiceOrderValue) ? null : choiceOrderValue,
    createdAt,
    status: (record.etat ?? 'en attente').toLowerCase(),
    isActive: record.is_active,
    planningReference: record.planning_reference ?? state.planningReference,
    tourNumber: record.tour_number ?? state.activeTourId
  };
};

const loadRequests = async () => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    setFeedback('Connexion à Supabase requise.');
    return;
  }
  state.isLoading = true;
  setFeedback('Chargement des demandes…');
  const query = supabase
    .from(CHOICES_TABLE)
    .select(
      'id, trigram, user_type, day, column_number, column_label, planning_day_label, slot_type_code, guard_nature, activity_type, choice_index, choice_rank, choice_order, created_at, etat, is_active, planning_reference, tour_number'
    );
  if (state.planningReference) {
    query.eq('planning_reference', state.planningReference);
  }
  if (state.activeTourId) {
    query.eq('tour_number', state.activeTourId);
  }
  query
    .order('choice_index', { ascending: false })
    .order('choice_rank', { ascending: true })
    .order('created_at', { ascending: true });
  const { data, error } = await query;
  state.isLoading = false;
  if (error) {
    console.error(error);
    setFeedback("Impossible de charger les demandes.");
    state.requests = [];
    renderRequests();
    return;
  }
  state.requests = (data ?? [])
    .map(mapRequestRecord)
    .filter((record) => record.isActive !== false);
  if (!state.requests.length) {
    setFeedback('Aucune demande enregistrée.');
  } else {
    setFeedback(`${state.requests.length} demande${state.requests.length > 1 ? 's' : ''} chargée${state.requests.length > 1 ? 's' : ''}.`);
  }
  renderRequests();
};

const recordAudit = async ({
  action,
  choiceId,
  targetTrigram,
  targetDay,
  targetColumnNumber,
  reason = null,
  metadata = {}
}) => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return;
  }
  const actor = await ensureActor();
  const payload = {
    action,
    choice_id: choiceId,
    target_trigram: targetTrigram,
    target_day: targetDay,
    target_column_number: targetColumnNumber,
    planning_reference: state.planningReference,
    tour_number: state.activeTourId,
    reason,
    metadata,
    actor_id: actor?.id ?? null,
    actor_trigram: actor?.trigram ?? null,
    actor_username: actor?.username ?? null
  };
  const { error } = await supabase.from(AUDIT_TABLE).insert(payload);
  if (error) {
    console.error('Audit log failed', error);
  }
};

const promoteAlternative = async (trigram, choiceIndex) => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return null;
  }
  const { data: alternative, error } = await supabase
    .from(CHOICES_TABLE)
    .select('id, choice_rank, day, column_number')
    .eq('planning_reference', state.planningReference)
    .eq('tour_number', state.activeTourId)
    .eq('trigram', trigram)
    .eq('choice_index', choiceIndex)
    .eq('etat', 'en attente')
    .gt('choice_rank', 1)
    .order('choice_rank', { ascending: true })
    .limit(1)
    .maybeSingle();
  if (error) {
    console.error(error);
    return;
  }
  if (!alternative) {
    return null;
  }
  const previousRank = alternative.choice_rank ?? null;
  const { error: updateError } = await supabase
    .from(CHOICES_TABLE)
    .update({ choice_rank: 1 })
    .eq('id', alternative.id);
  if (updateError) {
    console.error(updateError);
    return null;
  }
  await recordAudit({
    action: 'auto_promote',
    choiceId: alternative.id,
    targetTrigram: trigram,
    targetDay: alternative.day,
    targetColumnNumber: alternative.column_number,
    reason: "Promotion automatique de l'alternative"
  });
  return {
    choiceId: alternative.id,
    previous: { choice_rank: previousRank },
    next: { choice_rank: 1 },
    action: 'promote'
  };
};

const findAlternativesToDeactivate = async ({
  supabase,
  trigram,
  choiceIndex,
  selectedId,
  selectedRank,
  excludedIds = []
}) => {
  if (!supabase || !state.planningReference || !state.activeTourId) {
    return [];
  }
  const normalizedTrigram = typeof trigram === 'string' ? trigram.trim().toUpperCase() : '';
  const parsedChoiceIndex = Number.parseFloat(choiceIndex);
  const numericChoiceIndex = Number.isNaN(parsedChoiceIndex) ? Number.NaN : parsedChoiceIndex;
  if (!normalizedTrigram || Number.isNaN(numericChoiceIndex)) {
    return [];
  }

  const { data, error } = await supabase
    .from(CHOICES_TABLE)
    .select('id, etat, is_active, choice_rank, day, column_number')
    .eq('planning_reference', state.planningReference)
    .eq('tour_number', state.activeTourId)
    .eq('trigram', normalizedTrigram)
    .eq('choice_index', numericChoiceIndex)
    .neq('id', selectedId);
  if (error) {
    console.error(error);
    return [];
  }

  const excluded = new Set((excludedIds ?? []).map((value) => Number(value)));
  const parsedRank = Number.parseFloat(selectedRank);
  const rankThreshold = Number.isNaN(parsedRank) ? 1 : parsedRank;
  return (data ?? []).filter((record) => {
    if (!record || excluded.has(Number(record.id))) {
      return false;
    }
    const rankValue = Number.parseFloat(record.choice_rank);
    if (!Number.isNaN(rankValue) && rankValue <= rankThreshold) {
      return false;
    }
    if (record.etat === 'refusé' && record.is_active === false) {
      return false;
    }
    return true;
  });
};

const findChoiceOrderAlternatives = async ({
  supabase,
  trigram,
  choiceOrder,
  selectedId,
  excludedIds = []
}) => {
  if (!supabase || !state.planningReference || !state.activeTourId) {
    return [];
  }
  const normalizedTrigram = typeof trigram === 'string' ? trigram.trim().toUpperCase() : '';
  const parsedChoiceOrder = Number.parseInt(choiceOrder, 10);
  if (!normalizedTrigram || !Number.isFinite(parsedChoiceOrder)) {
    return [];
  }

  const excluded = new Set((excludedIds ?? []).map((value) => Number(value)));
  const { data, error } = await supabase
    .from(CHOICES_TABLE)
    .select('id, etat, is_active, choice_rank, day, column_number, choice_order')
    .eq('planning_reference', state.planningReference)
    .eq('tour_number', state.activeTourId)
    .eq('trigram', normalizedTrigram)
    .eq('choice_order', parsedChoiceOrder)
    .neq('id', selectedId);
  if (error) {
    console.error(error);
    return [];
  }

  return (data ?? []).filter((record) => record && !excluded.has(Number(record.id)));
};

const fetchCompetingRequests = async (request) => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return [];
  }
  const { data, error } = await supabase
    .from(CHOICES_TABLE)
    .select('id, trigram, choice_index, choice_rank, etat, day, column_number, is_active')
    .eq('planning_reference', request.planningReference)
    .eq('tour_number', request.tourNumber)
    .eq('day', request.day)
    .eq('column_number', request.columnNumber);
  if (error) {
    console.error(error);
    return [];
  }
  return data ?? [];
};

const hasConflictForDoctor = (request) => {
  const targetDay = getRequestDayKey(request);
  const targetColumn = getColumnInfo(request.columnNumber);
  const targetRange = targetColumn
    ? { start: targetColumn.start_time ?? null, end: targetColumn.end_time ?? null }
    : null;
  return state.requests.some((candidate) => {
    if (candidate.id === request.id) {
      return false;
    }
    if (candidate.trigram !== request.trigram) {
      return false;
    }
    if (candidate.status !== 'validé') {
      return false;
    }
    const candidateDay = getRequestDayKey(candidate);
    if (!candidateDay || !targetDay || candidateDay !== targetDay) {
      return false;
    }
    if (candidate.columnNumber === request.columnNumber) {
      return true;
    }
    const candidateColumn = getColumnInfo(candidate.columnNumber);
    if (!candidateColumn || !targetRange) {
      return false;
    }
    const candidateRange = {
      start: candidateColumn.start_time ?? null,
      end: candidateColumn.end_time ?? null
    };
    return rangesOverlap(targetRange, candidateRange);
  });
};

const acceptRequest = async (requestId) => {
  const request = state.requests.find((item) => item.id === Number(requestId));
  if (!request) {
    return;
  }
  if (request.status === 'validé') {
    setFeedback('La demande est déjà acceptée.');
    return;
  }
  if (hasConflictForDoctor(request)) {
    setFeedback('Conflit horaire détecté pour ce médecin.');
    return;
  }

  setFeedback('Validation de la demande…');
  const supabase = await ensureSupabase();
  if (!supabase) {
    setFeedback('Connexion à Supabase requise.');
    return;
  }

  const competing = await fetchCompetingRequests(request);

  const alternativeIds = collectAlternativeIds(competing, request);

  const parsedRequestRank = parseNumeric(request.choiceRank ?? request.choice_rank);
  const choiceRankThreshold = parsedRequestRank ?? 1;
  const additionalAlternatives = await findAlternativesToDeactivate({
    supabase,
    trigram: request.trigram,
    choiceIndex: request.choiceIndex,
    selectedId: request.id,
    selectedRank: choiceRankThreshold,
    excludedIds: alternativeIds
  });
  const alternativeIdSet = new Set(alternativeIds);
  additionalAlternatives.forEach((record) => {
    if (record?.id) {
      alternativeIdSet.add(record.id);
    }
  });
  const orderAlternatives = await findChoiceOrderAlternatives({
    supabase,
    trigram: request.trigram,
    choiceOrder: request.choiceOrder,
    selectedId: request.id,
    excludedIds: Array.from(alternativeIdSet)
  });
  orderAlternatives.forEach((record) => {
    if (record?.id && !alternativeIdSet.has(record.id)) {
      alternativeIdSet.add(record.id);
      additionalAlternatives.push(record);
    }
  });
  alternativeIds.length = 0;
  alternativeIdSet.forEach((id) => {
    alternativeIds.push(id);
  });
  const competingIds = competing
    .filter((item) => item.trigram !== request.trigram)
    .map((item) => item.id);
  const primaryRefused = competing.filter(
    (item) => item.trigram !== request.trigram && item.choice_rank === 1 && item.id !== request.id
  );

  const updates = [];

  updates.push(
    supabase
      .from(CHOICES_TABLE)
      .update({ etat: 'validé', is_active: true })
      .eq('id', request.id)
  );

  if (alternativeIds.length) {
    updates.push(
      supabase
        .from(CHOICES_TABLE)
        .update({ etat: 'refusé', is_active: false })
        .in('id', alternativeIds)
    );
  }

  if (competingIds.length) {
    updates.push(
      supabase
        .from(CHOICES_TABLE)
        .update({ etat: 'refusé' })
        .in('id', competingIds)
        .eq('etat', 'en attente')
    );
  }

  for (const operation of updates) {
    const { error } = await operation;
    if (error) {
      console.error(error);
      setFeedback("Impossible d'appliquer la décision d'acceptation.");
      return;
    }
  }

  await recordAudit({
    action: 'accept',
    choiceId: request.id,
    targetTrigram: request.trigram,
    targetDay: request.day,
    targetColumnNumber: request.columnNumber,
    reason: 'Demande acceptée'
  });

  for (const competitor of competingIds) {
    const refused = competing.find((item) => item.id === competitor);
    if (refused) {
      await recordAudit({
        action: 'refuse',
        choiceId: refused.id,
        targetTrigram: refused.trigram,
        targetDay: refused.day,
        targetColumnNumber: refused.column_number,
        reason: `Attribué à ${request.trigram}`
      });
    }
  }

  for (const competitor of primaryRefused) {
    await promoteAlternative(competitor.trigram, competitor.choice_index);
  }

  await loadRequests();
  setFeedback('Demande acceptée et planning mis à jour.');
};

const refuseRequest = async (requestId) => {
  const request = state.requests.find((item) => item.id === Number(requestId));
  if (!request) {
    return;
  }
  if (request.status === 'refusé') {
    setFeedback('La demande est déjà refusée.');
    return;
  }
  const reason = window.prompt('Motif du refus (optionnel) :', '');
  setFeedback('Refus de la demande…');

  const supabase = await ensureSupabase();
  if (!supabase) {
    setFeedback('Connexion à Supabase requise.');
    return;
  }

  const { error } = await supabase
    .from(CHOICES_TABLE)
    .update({ etat: 'refusé' })
    .eq('id', request.id);
  if (error) {
    console.error(error);
    setFeedback("Impossible de refuser cette demande.");
    return;
  }

  await recordAudit({
    action: 'refuse',
    choiceId: request.id,
    targetTrigram: request.trigram,
    targetDay: request.day,
    targetColumnNumber: request.columnNumber,
    reason: reason && reason.trim() ? reason.trim() : null
  });

  if (request.choiceRank === 1) {
    await promoteAlternative(request.trigram, request.choiceIndex);
  }

  await loadRequests();
  setFeedback('Demande refusée.');
};

const handleActionClick = (event) => {
  const button = event.target.closest('button[data-action]');
  if (!button) {
    return;
  }
  const requestId = button.dataset.requestId;
  const action = button.dataset.action;
  if (!requestId || !action) {
    return;
  }
  if (action === 'accept') {
    acceptRequest(requestId);
  } else if (action === 'refuse') {
    refuseRequest(requestId);
  }
};

const handleTabClick = (event) => {
  const button = event.target.closest('[data-tab]');
  if (!button) {
    return;
  }
  const tabId = button.dataset.tab;
  if (!tabId || state.activeTab === tabId) {
    return;
  }
  state.activeTab = tabId;
  renderRequests();
};

const updateUserTypeTabs = () => {
  if (!Array.isArray(elements.userTypeTabButtons)) {
    return;
  }
  const active = normalizeUserTypeFilter(state.filters.userType);
  elements.userTypeTabButtons.forEach((button) => {
    if (!button) {
      return;
    }
    const value = normalizeUserTypeFilter(button.dataset.userTypeTab ?? '');
    const isActive = !!value && value === active;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-selected', isActive ? 'true' : 'false');
    const shouldBeTabbable = isActive || !active;
    button.setAttribute('tabindex', shouldBeTabbable ? '0' : '-1');
  });
};

const setUserTypeFilter = (value) => {
  const normalized = normalizeUserTypeFilter(value);
  const current = normalizeUserTypeFilter(state.filters.userType);
  const next = current === normalized ? '' : normalized;
  state.filters = {
    ...state.filters,
    userType: next
  };
  updateUserTypeTabs();
  renderRequests();
};

const handleUserTypeTabClick = (event) => {
  const button = event.target.closest('[data-user-type-tab]');
  if (!button) {
    return;
  }
  const value = button.dataset.userTypeTab;
  if (!value) {
    return;
  }
  setUserTypeFilter(value);
};

const handleFiltersChange = () => {
  const form = elements.filtersForm;
  if (!form) {
    return;
  }
  const formData = new FormData(form);
  const previousUserType = normalizeUserTypeFilter(state.filters.userType);
  state.filters = {
    date: formData.get('date') ? formData.get('date').toString() : '',
    type: formData.get('type') ? formData.get('type').toString() : '',
    doctor: formData.get('doctor') ? formData.get('doctor').toString().trim() : '',
    column: formData.get('column') ? formData.get('column').toString().trim() : '',
    status: formData.get('status') ? formData.get('status').toString() : '',
    userType: previousUserType
  };
  updateUserTypeTabs();
  renderRequests();
};

const setViewMode = (mode) => {
  const normalized = mode === 'auto' ? 'auto' : 'requests';
  state.viewMode = normalized;
  updateModeTabs();
  updateModePanels();
  if (normalized === 'auto') {
    updateAutoAssignmentActiveTour();
    updateAutoAssignmentTrigramOptions();
  }
};

const handleModeTabClick = (event) => {
  const button = event.target.closest('[data-mode]');
  if (!button) {
    return;
  }
  const mode = button.dataset.mode === 'auto' ? 'auto' : 'requests';
  setViewMode(mode);
};

const handleAutoAssignmentFormChange = (event) => {
  const target = event.target;
  if (!target) {
    return;
  }
  if (target.name === 'population') {
    updateAutoAssignmentTrigramOptions();
  }
};

const getAutoAssignmentParameters = () => {
  const defaults = AUTO_ASSIGNMENT_DEFAULTS;
  const form = elements.autoForm;
  if (!form) {
    return { ...defaults };
  }
  const formData = new FormData(form);
  const populations = Array.from(
    new Set(
      formData
        .getAll('population')
        .map((value) => normalizeUserTypeFilter(value?.toString?.() ?? ''))
        .filter(Boolean)
    )
  );
  const orderValue = formData.get('order');
  const order = orderValue === 'desc' ? 'desc' : 'asc';
  const startTrigram = normalizeTrigram(formData.get('startTrigram'));
  const algorithmValue = formData.get('algorithm');
  const algorithm = AUTO_ASSIGNMENT_ALGORITHMS.has(algorithmValue)
    ? algorithmValue
    : defaults.algorithm;
  let rotations = Number.parseInt(formData.get('rotations'), 10);
  if (!Number.isInteger(rotations) || rotations < 1) {
    rotations = defaults.rotations;
  }
  return {
    populations,
    order,
    startTrigram,
    algorithm,
    rotations
  };
};

const validateAutoAssignmentParameters = (params) => {
  const errors = [];
  if (!params.populations || !params.populations.length) {
    errors.push('Sélectionnez au moins une population.');
  }
  if (!AUTO_ASSIGNMENT_ALGORITHMS.has(params.algorithm)) {
    errors.push('Algorithme inconnu.');
  }
  if (!Number.isInteger(params.rotations) || params.rotations < 1) {
    errors.push('Le nombre de rotations doit être un entier positif.');
  }
  const availableTrigrams = new Set(
    state.autoAssignment.directory
      .filter((entry) => params.populations.includes(entry.role))
      .map((entry) => entry.trigram)
      .filter(Boolean)
  );
  if (!params.startTrigram) {
    errors.push('Sélectionnez un trigramme de départ.');
  } else if (!availableTrigrams.has(params.startTrigram)) {
    errors.push('Le trigramme de départ doit appartenir à la population sélectionnée.');
  }
  return {
    isValid: errors.length === 0,
    errors
  };
};

const getUserTypeForTrigram = (trigram) => {
  const normalized = normalizeTrigram(trigram);
  if (!normalized) {
    return '';
  }
  const directoryEntry = state.autoAssignment.directory.find((entry) => entry.trigram === normalized);
  return directoryEntry?.role ?? '';
};

const getOrderedTrigrams = (params) => {
  const candidates = state.autoAssignment.directory
    .filter((entry) => params.populations.includes(entry.role))
    .map((entry) => entry.trigram)
    .filter(Boolean);
  const unique = Array.from(new Set(candidates));
  unique.sort((a, b) => a.localeCompare(b, 'fr', { sensitivity: 'base' }));
  if (params.order === 'desc') {
    unique.reverse();
  }
  const startIndex = unique.findIndex((value) => value === params.startTrigram);
  if (startIndex > 0) {
    return unique.slice(startIndex).concat(unique.slice(0, startIndex));
  }
  return unique;
};

const buildPendingRequestsMap = (params) => {
  const populations = new Set(params.populations);
  const map = new Map();
  state.requests.forEach((request) => {
    if (request.status !== 'en attente') {
      return;
    }
    const trigram = normalizeTrigram(request.trigram);
    if (!trigram) {
      return;
    }
    const userType = normalizeUserTypeFilter(request.userType);
    if (!populations.has(userType)) {
      return;
    }
    if (request.guardNature !== 'normale' && request.guardNature !== 'bonne') {
      return;
    }
    if (!map.has(trigram)) {
      map.set(trigram, {
        userType,
        normale: [],
        bonne: []
      });
    }
    const bucket = map.get(trigram);
    const listKey = request.guardNature === 'bonne' ? 'bonne' : 'normale';
    bucket[listKey].push(request);
  });
  map.forEach((bucket) => {
    bucket.normale = sortRequestsByPreference(bucket.normale);
    bucket.bonne = sortRequestsByPreference(bucket.bonne);
  });
  return map;
};

const buildAssignedSlotsMap = () => {
  const map = new Map();
  state.requests.forEach((request) => {
    if (request.status !== 'validé') {
      return;
    }
    const trigram = normalizeTrigram(request.trigram);
    if (!trigram) {
      return;
    }
    const dayKey = getRequestDayKey(request);
    if (!dayKey) {
      return;
    }
    if (!map.has(trigram)) {
      map.set(trigram, []);
    }
    map.get(trigram).push({
      dayKey,
      columnNumber: request.columnNumber,
      range: getColumnRange(request.columnNumber)
    });
  });
  return map;
};

const buildOccupiedSlotsSet = () => {
  const set = new Set();
  state.requests.forEach((request) => {
    if (request.status !== 'validé') {
      return;
    }
    const key = getSlotKey(request);
    if (key) {
      set.add(key);
    }
  });
  return set;
};

const findNextAvailableRequest = (list, trigram, assignedSlots, occupiedSlots, extraOccupied = new Set()) => {
  if (!Array.isArray(list) || !list.length) {
    return null;
  }
  for (let index = 0; index < list.length; index += 1) {
    const candidate = list[index];
    if (!candidate) {
      continue;
    }
    if (candidate.status && candidate.status !== 'en attente') {
      continue;
    }
    const slotKey = getSlotKey(candidate);
    if (!slotKey) {
      continue;
    }
    if (occupiedSlots.has(slotKey) || extraOccupied.has(slotKey)) {
      continue;
    }
    if (hasConflictWithAssignedSlots(trigram, candidate, assignedSlots)) {
      continue;
    }
    if (hasConflictForDoctor(candidate)) {
      continue;
    }
    return {
      request: candidate,
      index,
      slotKey
    };
  }
  return null;
};

const runSimpleAutoAssignment = (params) => {
  const orderedTrigrams = getOrderedTrigrams(params);
  const pendingMap = buildPendingRequestsMap(params);
  const assignedSlots = buildAssignedSlotsMap();
  const occupiedSlots = buildOccupiedSlotsSet();
  const assignments = [];
  let analysed = 0;
  let normals = 0;
  let good = 0;
  let skips = 0;
  let rotationsUsed = 0;

  if (!orderedTrigrams.length) {
    return {
      assignments: [],
      summary: {
        analysed: 0,
        normals: 0,
        good: 0,
        skips: 0,
        rotationsUsed: 0,
        maxRotations: params.rotations
      }
    };
  }

  let passNumber = 0;
  while (rotationsUsed < params.rotations) {
    passNumber += 1;
    let passAssignments = 0;
    for (const trigram of orderedTrigrams) {
      if (rotationsUsed >= params.rotations) {
        break;
      }
      analysed += 1;
      const entry = pendingMap.get(trigram);
      const result = {
        trigram,
        userType: entry?.userType ?? getUserTypeForTrigram(trigram),
        rotation: passNumber,
        assigned: false,
        normal: null,
        good: null,
        reason: ''
      };

      if (!entry || (!entry.normale.length && !entry.bonne.length)) {
        result.reason = 'Aucune demande éligible.';
        skips += 1;
        assignments.push(result);
        continue;
      }

      const tempOccupied = new Set();
      const normalCandidate = findNextAvailableRequest(entry.normale, trigram, assignedSlots, occupiedSlots, tempOccupied);
      if (normalCandidate?.slotKey) {
        tempOccupied.add(normalCandidate.slotKey);
      }
      const goodCandidate = findNextAvailableRequest(entry.bonne, trigram, assignedSlots, occupiedSlots, tempOccupied);

      if (normalCandidate && goodCandidate) {
        const normalRequest = normalCandidate.request;
        const goodRequest = goodCandidate.request;
        const normalDay = getRequestDayKey(normalRequest);
        const goodDay = getRequestDayKey(goodRequest);
        let pairConflict = false;
        if (normalDay && goodDay && normalDay === goodDay) {
          if (normalRequest.columnNumber === goodRequest.columnNumber) {
            pairConflict = true;
          } else {
            const normalRange = getColumnRange(normalRequest.columnNumber);
            const goodRange = getColumnRange(goodRequest.columnNumber);
            if (normalRange && goodRange && rangesOverlap(normalRange, goodRange)) {
              pairConflict = true;
            }
          }
        }

        if (!pairConflict) {
          rotationsUsed += 1;
          result.rotation = rotationsUsed;
          result.assigned = true;
          result.normal = normalRequest;
          result.good = goodRequest;
          assignments.push(result);
          passAssignments += 1;
          normals += 1;
          good += 1;

          entry.normale = entry.normale.filter((item) => item.id !== normalRequest.id);
          entry.bonne = entry.bonne.filter((item) => item.id !== goodRequest.id);

          const assignedList = assignedSlots.get(trigram) ?? [];
          if (normalDay) {
            assignedList.push({
              dayKey: normalDay,
              columnNumber: normalRequest.columnNumber,
              range: getColumnRange(normalRequest.columnNumber)
            });
          }
          if (goodDay) {
            assignedList.push({
              dayKey: goodDay,
              columnNumber: goodRequest.columnNumber,
              range: getColumnRange(goodRequest.columnNumber)
            });
          }
          assignedSlots.set(trigram, assignedList);

          if (normalCandidate.slotKey) {
            occupiedSlots.add(normalCandidate.slotKey);
          }
          if (goodCandidate.slotKey) {
            occupiedSlots.add(goodCandidate.slotKey);
          }
          if (rotationsUsed >= params.rotations) {
            break;
          }
          continue;
        }
        result.reason = 'Conflit horaire entre les gardes proposées.';
      } else if (!normalCandidate && !goodCandidate) {
        if (!entry.normale.length && !entry.bonne.length) {
          result.reason = 'Aucune demande éligible.';
        } else if (!entry.normale.length) {
          result.reason = 'Aucune garde normale disponible.';
        } else if (!entry.bonne.length) {
          result.reason = 'Aucune bonne garde disponible.';
        } else {
          result.reason = 'Aucune garde compatible disponible.';
        }
      } else if (!normalCandidate) {
        result.reason = 'Aucune garde normale disponible.';
      } else {
        result.reason = 'Aucune bonne garde disponible.';
      }

      skips += 1;
      assignments.push(result);
    }
    if (passAssignments === 0) {
      break;
    }
  }

  return {
    assignments,
    summary: {
      analysed,
      normals,
      good,
      skips,
      rotationsUsed,
      maxRotations: params.rotations
    }
  };
};

const runAutoAssignment = (params) => {
  if (params.algorithm === 'simple') {
    return runSimpleAutoAssignment(params);
  }
  throw new Error('Algorithme non pris en charge.');
};

const prepareAutoAssignmentWorkspace = async () => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return { success: false, message: 'Connexion à Supabase requise.' };
  }
  if (!state.planningReference || !state.activeTourId) {
    return { success: false, message: 'Référence de planning introuvable.' };
  }

  const { error: cleanupError } = await supabase
    .from(AUTO_ASSIGNMENT_WORK_TABLE)
    .delete()
    .eq('planning_reference', state.planningReference)
    .eq('tour_number', state.activeTourId);
  if (cleanupError) {
    console.error(cleanupError);
    return { success: false, message: 'Impossible de réinitialiser la table de travail.' };
  }

  const { data, error } = await supabase
    .from(CHOICES_TABLE)
    .select(
      'id, user_id, trigram, user_type, day, column_number, column_label, planning_day_label, slot_type_code, guard_nature, activity_type, choice_index, choice_rank, etat, is_active, planning_reference, tour_number, created_at'
    )
    .eq('planning_reference', state.planningReference)
    .eq('tour_number', state.activeTourId)
    .eq('etat', 'en attente');

  if (error) {
    console.error(error);
    return { success: false, message: 'Impossible de préparer les choix en attente.' };
  }

  const entries = (data ?? [])
    .filter((record) => record && record.is_active !== false)
    .map((record) => {
      const consolidatedIndex = computeConsolidatedIndex(record.choice_index, record.choice_rank);
      const rootChoiceIndex = getRootChoiceIndex(record.choice_index);
      const choiceIndexValue = getRootChoiceIndex(record.choice_index);
      const choiceRankValue = Number.parseInt(record.choice_rank, 10);
      const priorityValue = Number.parseInt(record.choice_rank, 10);
      return {
        choice_id: record.id,
        planning_reference: record.planning_reference ?? state.planningReference,
        planning_version: record.planning_reference ?? state.planningReference,
        tour_number: record.tour_number ?? state.activeTourId,
        trigram: normalizeTrigram(record.trigram ?? ''),
        user_id: record.user_id ?? null,
        user_type: normalizeUserTypeFilter(record.user_type ?? '') || null,
        day: record.day ?? null,
        column_number: record.column_number ?? null,
        column_label: record.column_label ?? null,
        planning_day_label: record.planning_day_label ?? null,
        slot_type_code: record.slot_type_code ?? null,
        guard_nature: record.guard_nature ?? null,
        activity_type: record.activity_type ?? null,
        choice_index: Number.isFinite(choiceIndexValue) ? choiceIndexValue : null,
        root_choice_index: rootChoiceIndex,
        choice_rank: Number.isFinite(choiceRankValue) ? choiceRankValue : null,
        consolidated_index: consolidatedIndex,
        priority: Number.isFinite(priorityValue) ? priorityValue : null,
        status: record.etat ?? 'en attente',
        is_active: record.is_active ?? null,
        created_at: record.created_at ?? new Date().toISOString(),
        metadata: {
          planning_day_label: record.planning_day_label ?? null,
          slot_type_code: record.slot_type_code ?? null
        }
      };
    });

  const chunkSize = 100;
  for (let index = 0; index < entries.length; index += chunkSize) {
    const chunk = entries.slice(index, index + chunkSize);
    if (!chunk.length) {
      continue;
    }
    const { error: insertError } = await supabase.from(AUTO_ASSIGNMENT_WORK_TABLE).insert(chunk);
    if (insertError) {
      console.error(insertError);
      return { success: false, message: 'Impossible de sauvegarder la table de travail.' };
    }
  }

  return { success: true };
};

const formatAutoAssignmentGuard = (request) => {
  if (!request) {
    return '—';
  }
  const parts = [];
  parts.push(formatDate(request.day));
  const label =
    request.columnLabel ||
    request.planningDayLabel ||
    request.slotTypeCode ||
    (request.columnNumber != null ? `Colonne ${request.columnNumber}` : null);
  if (label) {
    parts.push(label);
  }
  return parts.join(' • ');
};

const formatAutoAssignmentSummary = (summary) => {
  if (!summary) {
    return '';
  }
  const analysedLabel = `${summary.analysed ?? 0} médecin${summary.analysed === 1 ? '' : 's'} analysé${summary.analysed === 1 ? '' : 's'}`;
  const normalsLabel = `${summary.normals ?? 0} normale${summary.normals === 1 ? '' : 's'}`;
  const goodLabel = `${summary.good ?? 0} bonne${summary.good === 1 ? '' : 's'}`;
  const skipsLabel = `${summary.skips ?? 0} skip${summary.skips === 1 ? '' : 's'}`;
  const rotationsLabel = `${summary.rotationsUsed ?? 0}/${summary.maxRotations ?? 0}`;
  return `Terminé — ${analysedLabel} • ${normalsLabel} • ${goodLabel} • ${skipsLabel} • Rotations utilisées : ${rotationsLabel}.`;
};

const renderAutoAssignmentResults = (result) => {
  state.autoAssignment.lastResult = result ?? null;
  if (!elements.autoResultsSection || !elements.autoResultsBody) {
    return;
  }
  if (!result) {
    elements.autoResultsSection.classList.add('hidden');
    if (elements.autoSummary) {
      elements.autoSummary.textContent = '';
    }
    return;
  }

  elements.autoResultsSection.classList.remove('hidden');
  elements.autoResultsBody.innerHTML = '';
  const rows = result.assignments ?? [];
  if (!rows.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 5;
    cell.textContent = 'Aucune donnée à afficher.';
    cell.className = 'requests-empty-cell';
    row.appendChild(cell);
    elements.autoResultsBody.appendChild(row);
  } else {
    rows.forEach((item) => {
      const row = document.createElement('tr');
      const rotationCell = document.createElement('td');
      rotationCell.textContent = String(item.rotation ?? '—');
      row.appendChild(rotationCell);

      const doctorCell = document.createElement('td');
      const trigram = normalizeTrigram(item.trigram ?? '');
      let userTypeLabel = '';
      if (item.userType === 'medecin') {
        userTypeLabel = ' (Médecin)';
      } else if (item.userType === 'remplacant') {
        userTypeLabel = ' (Remplaçant)';
      }
      doctorCell.textContent = trigram ? `${trigram}${userTypeLabel}` : `—${userTypeLabel}`;
      row.appendChild(doctorCell);

      const normalCell = document.createElement('td');
      normalCell.textContent = formatAutoAssignmentGuard(item.normal);
      row.appendChild(normalCell);

      const goodCell = document.createElement('td');
      goodCell.textContent = formatAutoAssignmentGuard(item.good);
      row.appendChild(goodCell);

      const statusCell = document.createElement('td');
      if (item.assigned) {
        const badge = document.createElement('span');
        badge.className = 'badge badge-success';
        badge.textContent = 'Attribué';
        statusCell.appendChild(badge);
      } else {
        const badge = document.createElement('span');
        badge.className = 'badge badge-warning';
        badge.textContent = item.reason || 'Non attribué';
        statusCell.appendChild(badge);
      }
      row.appendChild(statusCell);

      elements.autoResultsBody.appendChild(row);
    });
  }

  if (elements.autoSummary) {
    elements.autoSummary.textContent = formatAutoAssignmentSummary(result.summary);
  }
};

const updateUndoButtonState = () => {
  if (!elements.autoUndoButton) {
    return;
  }
  const shouldDisable = state.autoAssignment.isRunning || state.autoAssignment.lastRunId == null;
  elements.autoUndoButton.disabled = shouldDisable;
};

const fetchLastAutoAssignmentRun = async () => {
  const supabase = await ensureSupabase();
  if (!supabase || !state.planningReference || !state.activeTourId) {
    state.autoAssignment.lastRunId = null;
    updateUndoButtonState();
    return;
  }
  const { data, error } = await supabase
    .from(AUTO_ASSIGNMENT_RUNS_TABLE)
    .select('id')
    .eq('planning_reference', state.planningReference)
    .eq('tour_number', state.activeTourId)
    .order('created_at', { ascending: false })
    .limit(1)
    .maybeSingle();
  if (error) {
    console.error(error);
    state.autoAssignment.lastRunId = null;
  } else {
    state.autoAssignment.lastRunId = data?.id ?? null;
  }
  updateUndoButtonState();
};

const cleanupAutoAssignmentWorkspaceAfterAcceptance = async ({
  supabase,
  accepted,
  alternativeIds = [],
  additionalAlternatives = [],
  competingIds = []
}) => {
  if (!supabase || !state.planningReference || !state.activeTourId) {
    return;
  }

  const idsToRemove = new Set();
  if (accepted?.id) {
    idsToRemove.add(accepted.id);
  }
  (alternativeIds ?? []).forEach((id) => idsToRemove.add(id));
  (additionalAlternatives ?? []).forEach((record) => {
    if (record?.id) {
      idsToRemove.add(record.id);
    }
  });
  (competingIds ?? []).forEach((id) => idsToRemove.add(id));

  if (idsToRemove.size) {
    const { error: deleteByIdError } = await supabase
      .from(AUTO_ASSIGNMENT_WORK_TABLE)
      .delete()
      .eq('planning_reference', state.planningReference)
      .eq('tour_number', state.activeTourId)
      .in('choice_id', Array.from(idsToRemove));
    if (deleteByIdError) {
      console.error(deleteByIdError);
    }
  }

  const trigram = typeof accepted?.trigram === 'string' ? accepted.trigram.trim().toUpperCase() : null;
  const rootChoiceIndex = getRootChoiceIndex(accepted?.choiceIndex ?? accepted?.choice_index);
  if (trigram && rootChoiceIndex != null) {
    const { error: deleteByRootError } = await supabase
      .from(AUTO_ASSIGNMENT_WORK_TABLE)
      .delete()
      .eq('planning_reference', state.planningReference)
      .eq('tour_number', state.activeTourId)
      .eq('trigram', trigram)
      .eq('root_choice_index', rootChoiceIndex);
    if (deleteByRootError) {
      console.error(deleteByRootError);
    }
  }
};

const applyAutomaticAcceptance = async (request, { guardType }) => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return { success: false, message: 'Connexion à Supabase requise.' };
  }
  if (!request?.id) {
    return { success: false, message: 'Demande introuvable.' };
  }

  const stored = state.requests.find((item) => item.id === request.id) ?? request;
  const competing = await fetchCompetingRequests(stored);

  const alternativeIds = collectAlternativeIds(competing, stored);
  const competingIds = competing
    .filter((item) => item.trigram !== stored.trigram)
    .map((item) => item.id);
  const primaryRefused = competing.filter(
    (item) => item.trigram !== stored.trigram && item.choice_rank === 1 && item.id !== stored.id
  );

  const idsToFetch = Array.from(new Set([stored.id, ...alternativeIds, ...competingIds]));
  const stateById = new Map();
  if (idsToFetch.length) {
    const { data: states, error } = await supabase
      .from(CHOICES_TABLE)
      .select('id, etat, is_active, choice_rank, choice_order')
      .in('id', idsToFetch);
    if (error) {
      console.error(error);
      return { success: false, message: "Impossible de récupérer l'état des demandes." };
    }
    (states ?? []).forEach((record) => {
      stateById.set(record.id, record);
    });
  }

  const operations = [];
  const updates = [];
  const targetState = stateById.get(stored.id) ?? {
    etat: stored.status ?? 'en attente',
    is_active: stored.isActive ?? true,
    choice_rank: stored.choiceRank ?? 1,
    choice_order: stored.choiceOrder ?? stored.choice_order ?? null
  };
  const parsedTargetRank = parseNumeric(targetState.choice_rank);
  const parsedStoredRank = parseNumeric(stored.choiceRank ?? stored.choice_rank);
  const selectedRankThreshold = parsedTargetRank ?? parsedStoredRank ?? 1;

  const selectedForAlternatives = {
    ...stored,
    choiceRank: selectedRankThreshold
  };

  const normalizedAlternativeIds = collectAlternativeIds(competing, selectedForAlternatives);
  const alternativeIdSet = new Set([...alternativeIds, ...normalizedAlternativeIds]);

  const additionalAlternatives = await findAlternativesToDeactivate({
    supabase,
    trigram: stored.trigram,
    choiceIndex: stored.choiceIndex,
    selectedId: stored.id,
    selectedRank: selectedRankThreshold,
    excludedIds: alternativeIds
  });

  additionalAlternatives.forEach((record) => {
    if (record?.id) {
      alternativeIdSet.add(record.id);
    }
  });

  const storedChoiceOrder = parseNumeric(
    targetState.choice_order ?? stored.choiceOrder ?? stored.choice_order
  );

  const choiceOrderAlternatives = await findChoiceOrderAlternatives({
    supabase,
    trigram: stored.trigram,
    choiceOrder: storedChoiceOrder,
    selectedId: stored.id,
    excludedIds: Array.from(alternativeIdSet)
  });

  const additionalAlternativeRecords = [...additionalAlternatives];
  choiceOrderAlternatives.forEach((record) => {
    if (record?.id && !alternativeIdSet.has(record.id)) {
      alternativeIdSet.add(record.id);
      additionalAlternativeRecords.push(record);
    }
  });

  alternativeIds.length = 0;
  alternativeIdSet.forEach((id) => alternativeIds.push(id));

  const missingStateIds = alternativeIds.filter((id) => !stateById.has(id));
  if (missingStateIds.length) {
    const { data: extraStates, error: extraStatesError } = await supabase
      .from(CHOICES_TABLE)
      .select('id, etat, is_active, choice_rank, choice_order')
      .in('id', missingStateIds);
    if (extraStatesError) {
      console.error(extraStatesError);
      return { success: false, message: "Impossible de récupérer l'état des alternatives." };
    }
    (extraStates ?? []).forEach((record) => {
      stateById.set(record.id, record);
    });
  }

  if (targetState.etat !== 'validé' || targetState.is_active !== true) {
    updates.push(async () => {
      const { error } = await supabase
        .from(CHOICES_TABLE)
        .update({ etat: 'validé', is_active: true })
        .eq('id', stored.id);
      if (error) {
        throw error;
      }
    });
    operations.push({
      choiceId: stored.id,
      action: 'accept',
      previous: {
        etat: targetState.etat ?? null,
        is_active: targetState.is_active ?? null,
        choice_rank: targetState.choice_rank ?? null
      },
      next: {
        etat: 'validé',
        is_active: true,
        choice_rank: targetState.choice_rank ?? stored.choiceRank ?? 1
      },
      reason: `Attribution automatique (${guardType ?? 'garde'})`
    });
  }

  alternativeIds.forEach((id) => {
    const state = stateById.get(id);
    const shouldUpdate = !state || state.etat !== 'refusé' || state.is_active !== false;
    if (!shouldUpdate) {
      return;
    }
    updates.push(async () => {
      const { error } = await supabase
        .from(CHOICES_TABLE)
        .update({ etat: 'refusé', is_active: false })
        .eq('id', id);
      if (error) {
        throw error;
      }
    });
    operations.push({
      choiceId: id,
      action: 'refuse_alternative',
      previous: {
        etat: state?.etat ?? null,
        is_active: state?.is_active ?? null,
        choice_rank: state?.choice_rank ?? null
      },
      next: {
        etat: 'refusé',
        is_active: false,
        choice_rank: state?.choice_rank ?? null
      },
      reason: 'Alternative refusée automatiquement'
    });
  });

  additionalAlternativeRecords.forEach((record) => {
    updates.push(async () => {
      const { error } = await supabase
        .from(CHOICES_TABLE)
        .update({ etat: 'refusé', is_active: false })
        .eq('id', record.id);
      if (error) {
        throw error;
      }
    });
    operations.push({
      choiceId: record.id,
      action: 'refuse_alternative',
      previous: {
        etat: record.etat ?? null,
        is_active: record.is_active ?? null,
        choice_rank: record.choice_rank ?? null
      },
      next: {
        etat: 'refusé',
        is_active: false,
        choice_rank: record.choice_rank ?? null
      },
      reason: 'Alternative retirée automatiquement'
    });
  });

  competingIds.forEach((id) => {
    const state = stateById.get(id);
    if (!state || state.etat !== 'en attente') {
      return;
    }
    updates.push(async () => {
      const { error } = await supabase
        .from(CHOICES_TABLE)
        .update({ etat: 'refusé' })
        .eq('id', id)
        .eq('etat', 'en attente');
      if (error) {
        throw error;
      }
    });
    operations.push({
      choiceId: id,
      action: 'refuse',
      previous: {
        etat: state.etat ?? null,
        is_active: state.is_active ?? null,
        choice_rank: state.choice_rank ?? null
      },
      next: {
        etat: 'refusé',
        is_active: state.is_active ?? null,
        choice_rank: state.choice_rank ?? null
      },
      reason: `Conflit avec ${stored.trigram}`
    });
  });

  try {
    for (const operation of updates) {
      await operation();
    }
  } catch (error) {
    console.error(error);
    return { success: false, message: "Impossible d'appliquer une attribution automatique." };
  }

  await recordAudit({
    action: 'accept',
    choiceId: stored.id,
    targetTrigram: stored.trigram,
    targetDay: stored.day,
    targetColumnNumber: stored.columnNumber,
    reason: 'Attribution automatique',
    metadata: { guardType }
  });

  for (const competitorId of competingIds) {
    const competitor = competing.find((item) => item.id === competitorId);
    if (competitor) {
      await recordAudit({
        action: 'refuse',
        choiceId: competitor.id,
        targetTrigram: competitor.trigram,
        targetDay: competitor.day,
        targetColumnNumber: competitor.column_number,
        reason: `Attribué automatiquement à ${stored.trigram}`
      });
    }
  }

  for (const competitor of primaryRefused) {
    const promotion = await promoteAlternative(competitor.trigram, competitor.choice_index);
    if (promotion) {
      operations.push({
        choiceId: promotion.choiceId,
        action: promotion.action,
        previous: promotion.previous,
        next: promotion.next,
        reason: 'Promotion automatique après refus'
      });
    }
  }

  await cleanupAutoAssignmentWorkspaceAfterAcceptance({
    supabase,
    accepted: stored,
    alternativeIds,
    additionalAlternatives: additionalAlternativeRecords,
    competingIds
  });

  return { success: true, operations };
};

const recordAutoAssignmentRun = async ({ params, result, operations }) => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return null;
  }
  const actor = await ensureActor();
  const payload = {
    actor_id: actor?.id ?? null,
    actor_trigram: actor?.trigram ?? null,
    actor_username: actor?.username ?? null,
    planning_reference: state.planningReference,
    tour_number: state.activeTourId,
    rotations_used: result?.summary?.rotationsUsed ?? 0,
    parameters: {
      ...params,
      populations: params.populations,
      order: params.order
    },
    summary: {
      ...result?.summary,
      assignments: (result?.assignments ?? []).map((item) => ({
        trigram: normalizeTrigram(item.trigram ?? ''),
        userType: item.userType ?? '',
        rotation: item.rotation ?? null,
        assigned: item.assigned ?? false,
        reason: item.reason ?? null,
        normalChoiceId: item.normal?.id ?? null,
        goodChoiceId: item.good?.id ?? null
      }))
    }
  };

  const { data, error } = await supabase
    .from(AUTO_ASSIGNMENT_RUNS_TABLE)
    .insert(payload)
    .select('id')
    .single();
  if (error) {
    console.error(error);
    return null;
  }
  const runId = data?.id ?? null;
  if (!runId) {
    return null;
  }

  if (operations?.length) {
    const entriesPayload = operations.map((operation) => ({
      run_id: runId,
      choice_id: operation.choiceId,
      action: operation.action,
      previous_state: operation.previous ?? {},
      next_state: operation.next ?? {},
      reason: operation.reason ?? null
    }));
    const { error: entriesError } = await supabase
      .from(AUTO_ASSIGNMENT_RUN_ENTRIES_TABLE)
      .insert(entriesPayload);
    if (entriesError) {
      console.error(entriesError);
    }
  }

  state.autoAssignment.lastRunId = runId;
  updateUndoButtonState();
  return runId;
};

const applyAutoAssignmentResult = async (result, params) => {
  const assignments = (result.assignments ?? []).filter((item) => item.assigned && item.normal && item.good);
  if (!assignments.length) {
    return { success: false, message: 'Aucune attribution à appliquer.' };
  }
  const collectedOperations = [];
  for (const assignment of assignments) {
    const normalOutcome = await applyAutomaticAcceptance(assignment.normal, { guardType: 'normale' });
    if (!normalOutcome.success) {
      return normalOutcome;
    }
    collectedOperations.push(...(normalOutcome.operations ?? []));
    const goodOutcome = await applyAutomaticAcceptance(assignment.good, { guardType: 'bonne' });
    if (!goodOutcome.success) {
      return goodOutcome;
    }
    collectedOperations.push(...(goodOutcome.operations ?? []));
  }

  const runId = await recordAutoAssignmentRun({ params, result, operations: collectedOperations });
  return { success: true, runId };
};

const undoLastAutoAssignment = async () => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    setAutoAssignmentFeedback('Connexion à Supabase requise.');
    return;
  }
  if (!state.planningReference || !state.activeTourId) {
    setAutoAssignmentFeedback("Aucun lot à annuler.");
    return;
  }

  setAutoAssignmentRunning(true);
  try {
    const { data: run, error } = await supabase
      .from(AUTO_ASSIGNMENT_RUNS_TABLE)
      .select('id')
      .eq('planning_reference', state.planningReference)
      .eq('tour_number', state.activeTourId)
      .order('created_at', { ascending: false })
      .limit(1)
      .maybeSingle();
    if (error || !run?.id) {
      setAutoAssignmentFeedback("Aucun lot d'attribution automatique à annuler.");
      return;
    }

    const runId = run.id;
    const { data: entries, error: entriesError } = await supabase
      .from(AUTO_ASSIGNMENT_RUN_ENTRIES_TABLE)
      .select('choice_id, previous_state')
      .eq('run_id', runId);
    if (entriesError) {
      console.error(entriesError);
      setAutoAssignmentFeedback("Impossible de récupérer le détail du lot.");
      return;
    }

    for (const entry of entries ?? []) {
      const previous = entry.previous_state ?? {};
      const updatePayload = {};
      if (Object.prototype.hasOwnProperty.call(previous, 'etat')) {
        updatePayload.etat = previous.etat;
      }
      if (Object.prototype.hasOwnProperty.call(previous, 'is_active')) {
        updatePayload.is_active = previous.is_active;
      }
      if (Object.prototype.hasOwnProperty.call(previous, 'choice_rank')) {
        updatePayload.choice_rank = previous.choice_rank;
      }
      if (Object.keys(updatePayload).length) {
        const { error: updateError } = await supabase
          .from(CHOICES_TABLE)
          .update(updatePayload)
          .eq('id', entry.choice_id);
        if (updateError) {
          console.error(updateError);
          setAutoAssignmentFeedback("Impossible d'annuler le lot en entier.");
          return;
        }
      }
    }

    await supabase.from(AUTO_ASSIGNMENT_RUN_ENTRIES_TABLE).delete().eq('run_id', runId);
    await supabase.from(AUTO_ASSIGNMENT_RUNS_TABLE).delete().eq('id', runId);
    state.autoAssignment.lastRunId = null;
    updateUndoButtonState();
    await loadRequests();
    const workspaceReset = await prepareAutoAssignmentWorkspace();
    if (!workspaceReset.success) {
      console.warn('Préparation de la table de travail échouée après annulation.', workspaceReset.message);
    }
    await fetchLastAutoAssignmentRun();
    setAutoAssignmentFeedback('Dernière attribution automatique annulée.');
  } catch (error) {
    console.error(error);
    setAutoAssignmentFeedback("Une erreur est survenue lors de l'annulation.");
  } finally {
    setAutoAssignmentRunning(false);
  }
};

const executeAutoAssignment = async ({ apply }) => {
  const params = getAutoAssignmentParameters();
  const validation = validateAutoAssignmentParameters(params);
  if (!validation.isValid) {
    setAutoAssignmentFeedback(validation.errors.join(' '));
    return;
  }
  setAutoAssignmentRunning(true);
  setAutoAssignmentFeedback(apply ? 'Attribution automatique en cours…' : 'Prévisualisation en cours…');
  try {
    await loadRequests();
    const preparation = await prepareAutoAssignmentWorkspace();
    if (!preparation.success) {
      setAutoAssignmentFeedback(preparation.message ?? "Impossible de préparer l'attribution automatique.");
      return;
    }
    const result = runAutoAssignment(params);
    renderAutoAssignmentResults(result);
    if (apply) {
      const applyOutcome = await applyAutoAssignmentResult(result, params);
      if (!applyOutcome.success) {
        setAutoAssignmentFeedback(applyOutcome.message ?? "Impossible d'appliquer l'attribution automatique.");
        return;
      }
      await loadRequests();
      await fetchLastAutoAssignmentRun();
    }
    setAutoAssignmentFeedback(formatAutoAssignmentSummary(result.summary));
  } catch (error) {
    console.error('Erreur lors de l\'attribution automatique', error);
    setAutoAssignmentFeedback("Une erreur est survenue lors de l'attribution automatique.");
  } finally {
    setAutoAssignmentRunning(false);
  }
};

const handleAutoAssignmentAction = (event) => {
  const button = event.target.closest('[data-auto-action]');
  if (!button || state.autoAssignment.isRunning) {
    return;
  }
  const action = button.dataset.autoAction;
  if (action === 'preview') {
    executeAutoAssignment({ apply: false });
  } else if (action === 'apply') {
    executeAutoAssignment({ apply: true });
  } else if (action === 'undo') {
    undoLastAutoAssignment();
  }
};

const attachEventListeners = () => {
  if (elements.logoutBtn) {
    elements.logoutBtn.addEventListener('click', () => {
      setCurrentUser(null);
      window.location.replace('index.html');
    });
  }
  if (elements.backBtn) {
    elements.backBtn.addEventListener('click', () => {
      window.location.assign('admin.html');
    });
  }
  if (elements.modeTabs) {
    elements.modeTabs.addEventListener('click', handleModeTabClick);
  }
  if (elements.tabsNav) {
    elements.tabsNav.addEventListener('click', handleTabClick);
  }
  if (elements.userTypeTabsNav) {
    elements.userTypeTabsNav.addEventListener('click', handleUserTypeTabClick);
  }
  if (elements.tableBody) {
    elements.tableBody.addEventListener('click', handleActionClick);
  }
  if (elements.filtersForm) {
    elements.filtersForm.addEventListener('change', handleFiltersChange);
    elements.filtersForm.addEventListener('input', handleFiltersChange);
    elements.filtersForm.addEventListener('reset', () => {
      window.setTimeout(() => {
        state.filters = {
          date: '',
          type: '',
          doctor: '',
          column: '',
          status: '',
          userType: ''
        };
        updateUserTypeTabs();
        renderRequests();
      }, 0);
    });
  }
  if (elements.autoForm) {
    elements.autoForm.addEventListener('change', handleAutoAssignmentFormChange);
    elements.autoForm.addEventListener('click', handleAutoAssignmentAction);
  }
};

const initialize = async () => {
  const currentUser = getCurrentUser();
  if (!currentUser) {
    window.location.replace('index.html');
    return;
  }
  requireRole('administrateur');

  initializeConnectionModal();
  attachEventListeners();
  setViewMode('requests');
  updateUserTypeTabs();
  buildTabs();
  updateTabState();

  await ensureSupabase();
  await ensureActor();
  await loadAdministrativeSettings();
  await loadPlanningColumns();
  await loadRequests();
  await loadAutoAssignmentDirectory();
  await fetchLastAutoAssignmentRun();
  updateUndoButtonState();
};

initialize().catch((error) => {
  console.error('Failed to initialize request processing page', error);
  setFeedback('Une erreur est survenue lors du chargement de la page.');
});
