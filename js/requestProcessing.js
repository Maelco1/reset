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

const CHOICE_SERIES_LABELS = new Map([
  ['normale', 'Gardes normales'],
  ['bonne', 'Bonnes gardes']
]);

const ADMIN_SETTINGS_TABLE = 'parametres_administratifs';
const AUDIT_TABLE = 'planning_choice_audit';
const CHOICES_TABLE = 'planning_choices';

const SEMI_ASSIGNMENT_DEFAULTS = {
  populations: ['medecin', 'remplacant'],
  order: 'asc',
  startTrigram: '',
  rotations: 1,
  normalThreshold: 2,
  goodQuota: 1,
  stepValidation: true
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
  semiAssignment: {
    isRunning: false,
    directory: [],
    steps: [],
    currentIndex: -1,
    params: null,
    summary: null
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
  semiPanel: document.querySelector('#semi-assignment-panel'),
  semiForm: document.querySelector('#semi-assignment-form'),
  semiFeedback: document.querySelector('#semi-assignment-feedback'),
  semiActiveTour: document.querySelector('#semi-assignment-active-tour'),
  semiResultsSection: document.querySelector('#semi-assignment-results'),
  semiResultsBody: document.querySelector('#semi-assignment-results-body'),
  semiSummary: document.querySelector('#semi-assignment-summary'),
  semiTrigramOptions: document.querySelector('#semi-assignment-trigram-options'),
  semiStartInput: document.querySelector('#semi-assignment-start')
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

const normalizeGuardNature = (value) => {
  if (typeof value !== 'string') {
    return 'normale';
  }
  const normalized = value.toString().trim().toLowerCase();
  return normalized === 'bonne' ? 'bonne' : 'normale';
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

const haveMatchingChoiceOrder = (candidate, selectedOrder) => {
  if (selectedOrder == null) {
    return false;
  }
  const candidateOrder = parseNumeric(candidate.choice_order ?? candidate.choiceOrder);
  return candidateOrder != null && candidateOrder === selectedOrder;
};

const getAlternativeRelation = (candidate, selectedRank, selectedIndex, selectedOrder) => {
  if (haveMatchingChoiceOrder(candidate, selectedOrder)) {
    return 'order';
  }
  if (haveMatchingChoiceIndex(candidate, selectedIndex)) {
    return 'index';
  }
  const candidateRank = parseNumeric(candidate.choice_rank ?? candidate.choiceRank);
  if (candidateRank == null) {
    return 'rank';
  }
  const threshold = selectedRank ?? 1;
  return candidateRank > threshold ? 'rank' : null;
};

const shouldTreatAsAlternative = (candidate, selectedRank, selectedIndex, selectedOrder) =>
  getAlternativeRelation(candidate, selectedRank, selectedIndex, selectedOrder) != null;

const getRequestGroupKey = (request) => {
  if (!request) {
    return null;
  }
  const order = parseNumeric(request.choiceOrder ?? request.choice_order);
  if (Number.isFinite(order)) {
    return `order:${order}`;
  }
  const index = parseNumeric(request.choiceIndex ?? request.choice_index);
  if (Number.isFinite(index)) {
    return `index:${index}`;
  }
  if (request.id != null) {
    return `id:${String(request.id)}`;
  }
  return null;
};

const buildGroupStateKey = (trigram, guardType, groupKey) => {
  if (!groupKey) {
    return null;
  }
  const normalizedTrigram = normalizeTrigram(trigram ?? '');
  const normalizedGuard = guardType === 'bonne' ? 'bonne' : 'normale';
  return `${normalizedTrigram}:${normalizedGuard}:${groupKey}`;
};

const collectAlternativeIds = (competing, request) => {
  const selectedRank = parseNumeric(request.choiceRank ?? request.choice_rank);
  const selectedIndex = parseNumeric(request.choiceIndex ?? request.choice_index);
  const selectedOrder = parseNumeric(request.choiceOrder ?? request.choice_order);
  const selectedNature = normalizeGuardNature(request.guardNature ?? request.guard_nature);
  const alternatives = competing
    .filter((item) => item.trigram === request.trigram && item.id !== request.id)
    .filter((item) => normalizeGuardNature(item.guard_nature ?? item.guardNature) === selectedNature)
    .filter((item) => shouldTreatAsAlternative(item, selectedRank, selectedIndex, selectedOrder))
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

const setSemiAssignmentFeedback = (message) => {
  if (elements.semiFeedback) {
    elements.semiFeedback.textContent = message ?? '';
  }
};

const setSemiAssignmentRunning = (isRunning) => {
  state.semiAssignment.isRunning = isRunning;
  const form = elements.semiForm;
  if (form) {
    form.classList.toggle('is-loading', isRunning);
    form.querySelectorAll('input, select, button').forEach((control) => {
      if (!control) {
        return;
      }
      control.disabled = isRunning;
    });
  }
  updateSemiAssignmentControlState();
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
    const mode = button.dataset.mode === 'semi' ? 'semi' : 'requests';
    const isActive = mode === active;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-selected', isActive ? 'true' : 'false');
    button.setAttribute('tabindex', isActive ? '0' : '-1');
  });
};

const updateModePanels = () => {
  const isSemi = state.viewMode === 'semi';
  elements.requestPanels.forEach((panel) => {
    if (!panel) {
      return;
    }
    panel.classList.toggle('hidden', isSemi);
  });
  if (elements.semiPanel) {
    elements.semiPanel.classList.toggle('hidden', !isSemi);
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
  const indexA = Number.isFinite(a?.choiceIndex) ? a.choiceIndex : Number.POSITIVE_INFINITY;
  const indexB = Number.isFinite(b?.choiceIndex) ? b.choiceIndex : Number.POSITIVE_INFINITY;
  if (indexA !== indexB) {
    return indexA - indexB;
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
  updateSemiAssignmentActiveTour();
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

const updateSemiAssignmentActiveTour = () => {
  if (!elements.semiActiveTour) {
    return;
  }
  if (!state.activeTourId) {
    elements.semiActiveTour.textContent = '—';
    return;
  }
  const parts = [`Tour ${state.activeTourId}`];
  if (Number.isInteger(state.planningYear)) {
    parts.push(String(state.planningYear));
  }
  elements.semiActiveTour.textContent = parts.join(' • ');
};

const getSelectedSemiPopulations = () => {
  const form = elements.semiForm;
  if (!form) {
    return new Set();
  }
  const inputs = Array.from(form.querySelectorAll('input[name="population"]'));
  const selected = inputs
    .filter((input) => input.checked)
    .map((input) => normalizeUserTypeFilter(input.value));
  return new Set(selected.filter(Boolean));
};

const ensureSemiAssignmentStartTrigram = (options) => {
  const input = elements.semiStartInput;
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

const updateSemiAssignmentTrigramOptions = () => {
  const datalist = elements.semiTrigramOptions;
  if (!datalist) {
    return;
  }
  datalist.innerHTML = '';
  const populations = getSelectedSemiPopulations();
  const filtered = state.semiAssignment.directory.filter((entry) => {
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
  ensureSemiAssignmentStartTrigram(filtered);
};

const loadSemiAssignmentDirectory = async () => {
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
  state.semiAssignment.directory = (data ?? []).map((record) => ({
    id: record.id ?? null,
    username: record.username ?? '',
    trigram: normalizeTrigram(record.trigram ?? record.username ?? ''),
    role: normalizeUserTypeFilter(record.role ?? '')
  }));
  updateSemiAssignmentTrigramOptions();
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
    guardNature: normalizeGuardNature(record.guard_nature),
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
    .order('choice_index', { ascending: true })
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

const promoteAlternative = async (trigram, choiceIndex, guardNature) => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return null;
  }
  const normalizedNature = normalizeGuardNature(guardNature);
  const { data: alternative, error } = await supabase
    .from(CHOICES_TABLE)
    .select('id, choice_rank, day, column_number, guard_nature')
    .eq('planning_reference', state.planningReference)
    .eq('tour_number', state.activeTourId)
    .eq('trigram', trigram)
    .eq('choice_index', choiceIndex)
    .eq('guard_nature', normalizedNature)
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
  guardNature,
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
  const normalizedNature = normalizeGuardNature(guardNature);

  const { data, error } = await supabase
    .from(CHOICES_TABLE)
    .select('id, etat, is_active, choice_rank, day, column_number, guard_nature')
    .eq('planning_reference', state.planningReference)
    .eq('tour_number', state.activeTourId)
    .eq('trigram', normalizedTrigram)
    .eq('choice_index', numericChoiceIndex)
    .eq('guard_nature', normalizedNature)
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
  guardNature,
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
  const normalizedNature = normalizeGuardNature(guardNature);
  const { data, error } = await supabase
    .from(CHOICES_TABLE)
    .select('id, etat, is_active, choice_rank, day, column_number, choice_order, guard_nature')
    .eq('planning_reference', state.planningReference)
    .eq('tour_number', state.activeTourId)
    .eq('trigram', normalizedTrigram)
    .eq('choice_order', parsedChoiceOrder)
    .eq('guard_nature', normalizedNature)
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
    .select('id, trigram, choice_index, choice_rank, etat, day, column_number, is_active, guard_nature')
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
    guardNature: request.guardNature,
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
    guardNature: request.guardNature,
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
    await promoteAlternative(
      competitor.trigram,
      competitor.choice_index,
      competitor.guard_nature ?? competitor.guardNature
    );
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
  const normalized = mode === 'semi' ? 'semi' : 'requests';
  state.viewMode = normalized;
  updateModeTabs();
  updateModePanels();
  if (normalized === 'semi') {
    updateSemiAssignmentActiveTour();
    updateSemiAssignmentTrigramOptions();
    renderSemiAssignmentState();
  }
};

const handleModeTabClick = (event) => {
  const button = event.target.closest('[data-mode]');
  if (!button) {
    return;
  }
  const mode = button.dataset.mode === 'semi' ? 'semi' : 'requests';
  setViewMode(mode);
};

const handleSemiAssignmentFormChange = (event) => {
  const target = event.target;
  if (!target) {
    return;
  }
  if (target.name === 'population') {
    updateSemiAssignmentTrigramOptions();
  }
};

const getSemiAssignmentParameters = () => {
  const defaults = SEMI_ASSIGNMENT_DEFAULTS;
  const form = elements.semiForm;
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
  let rotations = Number.parseInt(formData.get('rotations'), 10);
  if (!Number.isInteger(rotations) || rotations < 1) {
    rotations = defaults.rotations;
  }
  let normalThreshold = Number.parseInt(formData.get('normalThreshold'), 10);
  if (!Number.isInteger(normalThreshold) || normalThreshold < 1) {
    normalThreshold = defaults.normalThreshold;
  }
  let goodQuota = Number.parseInt(formData.get('goodQuota'), 10);
  if (!Number.isInteger(goodQuota) || goodQuota < 0) {
    goodQuota = defaults.goodQuota;
  }
  const stepValidation = formData.get('stepValidation') != null;
  return {
    populations,
    order,
    startTrigram,
    rotations,
    normalThreshold,
    goodQuota,
    stepValidation
  };
};

const validateSemiAssignmentParameters = (params) => {
  const errors = [];
  if (!params.populations || !params.populations.length) {
    errors.push('Sélectionnez au moins une population.');
  }
  if (!Number.isInteger(params.rotations) || params.rotations < 1) {
    errors.push('Le nombre de rotations doit être un entier positif.');
  }
  if (!Number.isInteger(params.normalThreshold) || params.normalThreshold < 1) {
    errors.push('Le palier de gardes normales doit être un entier positif.');
  }
  if (!Number.isInteger(params.goodQuota) || params.goodQuota < 0) {
    errors.push('Le nombre de bonnes gardes doit être positif ou nul.');
  }
  const availableTrigrams = new Set(
    state.semiAssignment.directory
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
  const directoryEntry = state.semiAssignment.directory.find((entry) => entry.trigram === normalized);
  return directoryEntry?.role ?? '';
};

const getOrderedTrigrams = (params) => {
  const candidates = state.semiAssignment.directory
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




const formatChoicePrimaryLabel = (choiceIndex) => {
  if (!Number.isFinite(choiceIndex)) {
    return '—';
  }
  return String(Math.trunc(choiceIndex));
};

const formatChoiceAlternativeLabel = (choiceIndex, choiceRank, choiceOrder) => {
  const hasIndex = Number.isFinite(choiceIndex);
  const hasRank = Number.isFinite(choiceRank);
  if (hasIndex && hasRank) {
    const normalizedIndex = Math.trunc(choiceIndex);
    const normalizedRank = Math.max(1, Math.trunc(choiceRank));
    return `${normalizedIndex}.${normalizedRank}`;
  }
  if (hasRank) {
    return String(Math.max(1, Math.trunc(choiceRank)));
  }
  if (Number.isFinite(choiceOrder)) {
    return `Ord.${Math.trunc(choiceOrder)}`;
  }
  return '—';
};

const describeAlternativeRelation = (relation) => {
  if (relation === 'order') {
    return "Alternative (ordre identique)";
  }
  if (relation === 'index') {
    return "Alternative (index identique)";
  }
  if (relation === 'rank') {
    return 'Alternative (rang supérieur)';
  }
  return 'Sélection';
};

const createSemiAssignmentImpact = (request, guardType, role, relation = null) => {
  if (!request) {
    return null;
  }
  const dayKey = getRequestDayKey(request);
  const normalizedGuard = guardType === 'bonne' ? 'bonne' : 'normale';
  const columnLabel =
    request.columnLabel ||
    request.planningDayLabel ||
    request.slotTypeCode ||
    (Number.isFinite(request.columnNumber) ? `Colonne ${request.columnNumber}` : '');
  return {
    id: request.id ?? null,
    trigram: normalizeTrigram(request.trigram),
    guardType: normalizedGuard,
    role,
    relation,
    day: request.day ?? null,
    dayKey,
    columnNumber: Number.isFinite(request.columnNumber) ? request.columnNumber : null,
    choiceIndex: Number.isFinite(request.choiceIndex) ? request.choiceIndex : null,
    choiceRank: Number.isFinite(request.choiceRank) ? request.choiceRank : null,
    choiceOrder: Number.isFinite(request.choiceOrder) ? request.choiceOrder : null,
    primaryLabel: formatChoicePrimaryLabel(request.choiceIndex),
    alternativeLabel: formatChoiceAlternativeLabel(
      request.choiceIndex,
      request.choiceRank,
      request.choiceOrder
    ),
    description: describeAlternativeRelation(relation),
    location: columnLabel,
    summary: formatSemiAssignmentGuard(request)
  };
};

const findLocalAlternativesForRequest = (request) => {
  if (!request) {
    return [];
  }
  const trigram = normalizeTrigram(request.trigram);
  const guardNature = normalizeGuardNature(request.guardNature);
  const selectedRank = parseNumeric(request.choiceRank ?? request.choice_rank);
  const selectedIndex = parseNumeric(request.choiceIndex ?? request.choice_index);
  const selectedOrder = parseNumeric(request.choiceOrder ?? request.choice_order);
  return state.requests
    .filter((item) => item.id !== request.id)
    .filter((item) => normalizeTrigram(item.trigram) === trigram)
    .filter((item) => normalizeGuardNature(item.guardNature ?? item.guard_nature) === guardNature)
    .filter((item) => (item.status ?? '').toLowerCase() === 'en attente')
    .map((item) => ({
      record: item,
      relation: getAlternativeRelation(item, selectedRank, selectedIndex, selectedOrder)
    }))
    .filter((entry) => entry.relation != null)
    .map((entry) => ({
      record: entry.record,
      relation: entry.relation
    }));
};

const buildSemiAssignmentImpacts = (request, guardType) => {
  const selection = createSemiAssignmentImpact(request, guardType, 'assigned');
  const alternatives = findLocalAlternativesForRequest(request).map((entry) =>
    createSemiAssignmentImpact(entry.record, guardType, 'alternative', entry.relation)
  );
  return {
    assigned: selection ? [selection] : [],
    alternatives: alternatives.filter(Boolean)
  };
};

const computeSemiAssignmentSteps = (params) => {
  const orderedTrigrams = getOrderedTrigrams(params);
  const pendingMap = buildPendingRequestsMap(params);
  const assignedSlots = buildAssignedSlotsMap();
  const occupiedSlots = buildOccupiedSlotsSet();
  const usedGroupKeys = new Set();
  const validatedCounts = buildValidatedGuardCounts();

  state.requests.forEach((request) => {
    if (request?.status !== 'validé') {
      return;
    }
    const trigram = normalizeTrigram(request.trigram);
    if (!trigram) {
      return;
    }
    const guardType =
      request.guardNature === 'bonne'
        ? 'bonne'
        : request.guardNature === 'normale'
          ? 'normale'
          : '';
    if (!guardType) {
      return;
    }
    const groupKey = getRequestGroupKey(request);
    if (!groupKey) {
      return;
    }
    const stateKey = buildGroupStateKey(trigram, guardType, groupKey);
    if (stateKey) {
      usedGroupKeys.add(stateKey);
    }
    const entry = pendingMap.get(trigram);
    if (entry) {
      const listKey = guardType === 'bonne' ? 'bonne' : 'normale';
      entry[listKey] = entry[listKey].filter((item) => getRequestGroupKey(item) !== groupKey);
    }
  });

  const simulatedCounts = new Map();
  orderedTrigrams.forEach((trigram) => {
    const base = validatedCounts.get(trigram) ?? { normal: 0, good: 0 };
    simulatedCounts.set(trigram, { normal: base.normal, good: base.good });
  });

  const steps = [];
  let analysed = 0;
  let normals = 0;
  let good = 0;

  for (let rotation = 1; rotation <= params.rotations; rotation += 1) {
    for (const trigram of orderedTrigrams) {
      analysed += 1;
      const entry = pendingMap.get(trigram);
      if (!entry) {
        continue;
      }
      const counts = simulatedCounts.get(trigram) ?? { normal: 0, good: 0 };
      simulatedCounts.set(trigram, counts);
      const extraOccupied = new Set();

      const selectCandidate = (listKey, guardType) => {
        const candidate = findNextAvailableRequest(
          entry[listKey],
          trigram,
          assignedSlots,
          occupiedSlots,
          extraOccupied,
          { usedGroups: usedGroupKeys, guardType }
        );
        if (candidate?.slotKey) {
          extraOccupied.add(candidate.slotKey);
        }
        return candidate;
      };

      const registerSelection = (guardType, candidate) => {
        if (!candidate?.request) {
          return;
        }
        const listKey = guardType === 'bonne' ? 'bonne' : 'normale';
        entry[listKey] = entry[listKey].filter((item) => item.id !== candidate.request.id);
        const groupKey = candidate.groupKey ?? getRequestGroupKey(candidate.request);
        if (groupKey) {
          const stateKey = buildGroupStateKey(trigram, guardType, groupKey);
          if (stateKey) {
            usedGroupKeys.add(stateKey);
          }
          entry[listKey] = entry[listKey].filter((item) => getRequestGroupKey(item) !== groupKey);
        }
        if (candidate.slotKey) {
          occupiedSlots.add(candidate.slotKey);
        }
      };

      const addStep = (guardType, candidate) => {
        if (!candidate?.request) {
          return;
        }
        const impacts = buildSemiAssignmentImpacts(candidate.request, guardType);
        steps.push({
          trigram,
          userType: entry?.userType ?? getUserTypeForTrigram(trigram),
          rotation,
          guardType,
          request: candidate.request,
          status: 'pending',
          impacts
        });
        registerSelection(guardType, candidate);
        if (guardType === 'bonne') {
          counts.good += 1;
          good += 1;
        } else {
          counts.normal += 1;
          normals += 1;
        }
      };

      const normalCandidate = selectCandidate('normale', 'normale');
      if (normalCandidate) {
        addStep('normale', normalCandidate);
      }

      const grantedBlocks = params.normalThreshold > 0 ? Math.floor(counts.normal / params.normalThreshold) : 0;
      const allowedGood = grantedBlocks * params.goodQuota;
      if (params.goodQuota > 0 && allowedGood > counts.good) {
        const goodCandidate = selectCandidate('bonne', 'bonne');
        if (goodCandidate) {
          addStep('bonne', goodCandidate);
        }
      }
    }
  }

  return {
    steps,
    summary: {
      analysed,
      normals,
      good,
      rotations: params.rotations,
      uniqueDoctors: orderedTrigrams.length
    }
  };
};

const formatSemiAssignmentGuard = (request) => {
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

const formatSemiAssignmentSummary = (summary) => {
  if (!summary) {
    return '';
  }
  const analysedLabel = `${summary.analysed ?? 0} passage${summary.analysed === 1 ? '' : 's'}`;
  const normalsLabel = `${summary.normals ?? 0} garde${summary.normals === 1 ? '' : 's'} normale${summary.normals === 1 ? '' : 's'}`;
  const goodLabel = `${summary.good ?? 0} bonne${summary.good === 1 ? '' : 's'}`;
  const rotationsLabel = `${summary.rotations ?? 0} rotation${summary.rotations === 1 ? '' : 's'}`;
  return `${analysedLabel} analysé${summary.analysed === 1 ? '' : 's'} • ${normalsLabel} proposées • ${goodLabel} proposées • ${rotationsLabel}`;
};

const getCurrentSemiAssignmentStep = () => {
  if (state.semiAssignment.currentIndex < 0) {
    return null;
  }
  return state.semiAssignment.steps[state.semiAssignment.currentIndex] ?? null;
};

const updateSemiAssignmentControlState = () => {
  const form = elements.semiForm;
  if (!form) {
    return;
  }
  const { steps, currentIndex, isRunning } = state.semiAssignment;
  const hasSteps = steps.length > 0;
  const currentStep = getCurrentSemiAssignmentStep();
  form.querySelectorAll('[data-semi-action]').forEach((control) => {
    if (!(control instanceof HTMLButtonElement)) {
      return;
    }
    const action = control.dataset.semiAction;
    if (action === 'prepare') {
      control.disabled = isRunning;
    } else if (action === 'previous') {
      control.disabled = isRunning || !hasSteps || currentIndex <= 0;
    } else if (action === 'next') {
      control.disabled = isRunning || !hasSteps || currentIndex >= steps.length - 1;
    } else if (action === 'accept') {
      control.disabled = isRunning || !currentStep?.request;
    } else if (action === 'skip') {
      control.disabled = isRunning || !hasSteps;
    } else {
      control.disabled = isRunning;
    }
  });
};

const setSemiAssignmentSteps = (steps, summary, params) => {
  state.semiAssignment.steps = steps.map((step) => ({ ...step, status: step.status ?? 'pending' }));
  state.semiAssignment.summary = summary ?? null;
  state.semiAssignment.params = params ?? state.semiAssignment.params;
  state.semiAssignment.currentIndex = state.semiAssignment.steps.length ? 0 : -1;
  renderSemiAssignmentState();
};

const moveSemiAssignmentIndex = (offset) => {
  if (!state.semiAssignment.steps.length) {
    state.semiAssignment.currentIndex = -1;
    renderSemiAssignmentState();
    return;
  }
  const nextIndex = clamp(
    state.semiAssignment.currentIndex + offset,
    0,
    state.semiAssignment.steps.length - 1
  );
  state.semiAssignment.currentIndex = nextIndex;
  renderSemiAssignmentState();
};


const createSemiAssignmentMiniSlot = (item) => {
  if (!item) {
    return null;
  }
  const slot = document.createElement('div');
  slot.className = 'semi-assignment-mini-slot';
  const typeKey = item.guardType === 'bonne' ? 'good' : 'normal';
  const roleKey = item.role === 'alternative' ? 'alternative' : 'assigned';
  slot.classList.add(`semi-assignment-mini-slot--${typeKey}-${roleKey}`);
  const trigram = document.createElement('span');
  trigram.className = 'semi-assignment-mini-trigram';
  trigram.textContent = item.trigram || '—';
  slot.appendChild(trigram);
  const primary = document.createElement('span');
  primary.className = 'semi-assignment-mini-label semi-assignment-mini-label--primary';
  primary.textContent = `P: ${item.primaryLabel ?? '—'}`;
  slot.appendChild(primary);
  const alternative = document.createElement('span');
  alternative.className = 'semi-assignment-mini-label semi-assignment-mini-label--alt';
  alternative.textContent = `A: ${item.alternativeLabel ?? '—'}`;
  slot.appendChild(alternative);
  const sr = document.createElement('span');
  sr.className = 'sr-only';
  const descriptors = [item.trigram || '', item.summary || '', item.description || ''].filter(Boolean);
  sr.textContent = descriptors.join(' • ');
  slot.appendChild(sr);
  slot.title = descriptors.join(' • ');
  return slot;
};

const createSemiAssignmentPreview = (step) => {
  const container = document.createElement('div');
  container.className = 'semi-assignment-impacts';
  if (!step) {
    const empty = document.createElement('p');
    empty.className = 'semi-assignment-mini-empty';
    empty.textContent = 'Aucun impact identifié.';
    container.appendChild(empty);
    return container;
  }
  const impacts = step.impacts ?? buildSemiAssignmentImpacts(step.request, step.guardType);
  if (step && !step.impacts) {
    step.impacts = impacts;
  }
  const assigned = impacts?.assigned ?? [];
  const alternatives = impacts?.alternatives ?? [];
  const items = [...assigned, ...alternatives].filter(Boolean);
  if (!items.length) {
    const empty = document.createElement('p');
    empty.className = 'semi-assignment-mini-empty';
    empty.textContent = 'Aucun impact identifié.';
    container.appendChild(empty);
    return container;
  }
  const positioned = items.filter((item) => item.dayKey && Number.isFinite(item.columnNumber));
  if (positioned.length) {
    const columnNumbers = Array.from(new Set(positioned.map((item) => item.columnNumber))).sort((a, b) => a - b);
    const dayEntries = new Map();
    positioned.forEach((item) => {
      const key = item.dayKey;
      if (!dayEntries.has(key)) {
        const label = formatDate(item.day ?? key);
        dayEntries.set(key, { key, label, cells: new Map() });
      }
      const entry = dayEntries.get(key);
      if (!entry.cells.has(item.columnNumber)) {
        entry.cells.set(item.columnNumber, []);
      }
      entry.cells.get(item.columnNumber).push(item);
    });
    const dayRows = Array.from(dayEntries.values()).sort((a, b) => a.key.localeCompare(b.key));
    const table = document.createElement('table');
    table.className = 'semi-assignment-mini';
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const dayHeader = document.createElement('th');
    dayHeader.scope = 'col';
    dayHeader.textContent = 'Jour';
    headerRow.appendChild(dayHeader);
    columnNumbers.forEach((columnNumber) => {
      const th = document.createElement('th');
      th.scope = 'col';
      th.textContent = `Col. ${columnNumber}`;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    dayRows.forEach((dayRow) => {
      const row = document.createElement('tr');
      const dayCell = document.createElement('th');
      dayCell.scope = 'row';
      dayCell.textContent = dayRow.label;
      row.appendChild(dayCell);
      columnNumbers.forEach((columnNumber) => {
        const td = document.createElement('td');
        const cellItems = dayRow.cells.get(columnNumber) ?? [];
        if (!cellItems.length) {
          td.classList.add('is-empty');
        } else {
          td.classList.add('has-content');
          cellItems.forEach((item) => {
            const slot = createSemiAssignmentMiniSlot(item);
            if (slot) {
              td.appendChild(slot);
            }
          });
        }
        row.appendChild(td);
      });
      tbody.appendChild(row);
    });
    table.appendChild(tbody);
    container.appendChild(table);
  }
  const otherItems = items.filter((item) => !item.dayKey || !Number.isFinite(item.columnNumber));
  if (otherItems.length) {
    const chips = document.createElement('div');
    chips.className = 'semi-assignment-mini-meta';
    otherItems.forEach((item) => {
      const chip = document.createElement('span');
      chip.className = 'semi-assignment-mini-chip';
      const typeKey = item.guardType === 'bonne' ? 'good' : 'normal';
      const roleKey = item.role === 'alternative' ? 'alternative' : 'assigned';
      chip.classList.add(`semi-assignment-mini-chip--${typeKey}-${roleKey}`);
      chip.textContent = `${item.trigram || '—'} • P:${item.primaryLabel} • A:${item.alternativeLabel}`;
      chip.title = item.summary || '';
      chips.appendChild(chip);
    });
    container.appendChild(chips);
  }
  return container;
};


const renderSemiAssignmentState = () => {
  if (!elements.semiResultsSection || !elements.semiResultsBody) {
    return;
  }
  const { steps, currentIndex, summary } = state.semiAssignment;
  if (!steps.length) {
    elements.semiResultsSection.classList.add('hidden');
    if (elements.semiSummary) {
      elements.semiSummary.textContent = '';
    }
    updateSemiAssignmentControlState();
    return;
  }
  elements.semiResultsSection.classList.remove('hidden');
  elements.semiResultsBody.innerHTML = '';
  steps.forEach((step, index) => {
    const row = document.createElement('tr');
    if (index === currentIndex) {
      row.classList.add('is-active');
    }

    const rotationCell = document.createElement('td');
    rotationCell.textContent = String(step.rotation ?? '—');
    row.appendChild(rotationCell);

    const doctorCell = document.createElement('td');
    const trigram = normalizeTrigram(step.trigram ?? '');
    let userTypeLabel = '';
    if (step.userType === 'medecin') {
      userTypeLabel = ' (Médecin)';
    } else if (step.userType === 'remplacant') {
      userTypeLabel = ' (Remplaçant)';
    }
    doctorCell.textContent = trigram ? `${trigram}${userTypeLabel}` : `—${userTypeLabel}`;
    row.appendChild(doctorCell);

    const typeCell = document.createElement('td');
    typeCell.textContent = CHOICE_SERIES_LABELS.get(step.guardType) ?? step.guardType ?? '—';
    row.appendChild(typeCell);

    const slotCell = document.createElement('td');
    slotCell.textContent = formatSemiAssignmentGuard(step.request);
    row.appendChild(slotCell);

    const impactsCell = document.createElement('td');
    impactsCell.appendChild(createSemiAssignmentPreview(step));
    row.appendChild(impactsCell);

    const statusCell = document.createElement('td');
    const badge = document.createElement('span');
    badge.className = 'badge';
    if (step.status === 'accepted') {
      badge.classList.add('badge-success');
      badge.textContent = 'Validée';
    } else if (step.status === 'skipped') {
      badge.classList.add('badge-danger');
      badge.textContent = 'Ignorée';
    } else if (index === currentIndex) {
      badge.classList.add('badge-warning');
      badge.textContent = 'Étape active';
    } else {
      badge.textContent = 'En attente';
    }
    statusCell.appendChild(badge);
    row.appendChild(statusCell);

    elements.semiResultsBody.appendChild(row);
  });

  if (elements.semiSummary) {
    elements.semiSummary.textContent = formatSemiAssignmentSummary(summary);
  }
  updateSemiAssignmentControlState();
};

const prepareSemiAssignmentSequence = () => {
  const params = getSemiAssignmentParameters();
  const validation = validateSemiAssignmentParameters(params);
  if (!validation.isValid) {
    setSemiAssignmentFeedback(validation.errors.join(' '));
    return;
  }
  const { steps, summary } = computeSemiAssignmentSteps(params);
  setSemiAssignmentSteps(steps, summary, params);
  if (!steps.length) {
    setSemiAssignmentFeedback('Aucune attribution proposée pour ces paramètres.');
  } else {
    setSemiAssignmentFeedback(formatSemiAssignmentSummary(summary));
  }
};

const applySemiAssignmentAcceptance = async () => {
  const currentStep = getCurrentSemiAssignmentStep();
  if (!currentStep?.request) {
    setSemiAssignmentFeedback('Aucune attribution à valider.');
    return;
  }
  setSemiAssignmentRunning(true);
  try {
    await acceptRequest(currentStep.request.id);
    state.semiAssignment.steps[state.semiAssignment.currentIndex].status = 'accepted';
    if (state.semiAssignment.params) {
      const { steps, summary } = computeSemiAssignmentSteps(state.semiAssignment.params);
      setSemiAssignmentSteps(steps, summary, state.semiAssignment.params);
      if (steps.length) {
        setSemiAssignmentFeedback('Garde validée. Séquence recalculée.');
      } else {
        setSemiAssignmentFeedback('Garde validée. Plus aucune proposition.');
      }
    }
  } catch (error) {
    console.error('Erreur lors de la validation semi-automatique', error);
    setSemiAssignmentFeedback('Impossible de valider cette attribution.');
  } finally {
    setSemiAssignmentRunning(false);
  }
};

const skipSemiAssignmentStep = () => {
  const currentStep = getCurrentSemiAssignmentStep();
  if (!currentStep) {
    setSemiAssignmentFeedback('Aucune étape à ignorer.');
    return;
  }
  currentStep.status = 'skipped';
  moveSemiAssignmentIndex(1);
  setSemiAssignmentFeedback('Étape ignorée.');
};

const handleSemiAssignmentAction = (event) => {
  const button = event.target.closest('[data-semi-action]');
  if (!button || state.semiAssignment.isRunning) {
    return;
  }
  const action = button.dataset.semiAction;
  if (action === 'prepare') {
    prepareSemiAssignmentSequence();
  } else if (action === 'previous') {
    moveSemiAssignmentIndex(-1);
  } else if (action === 'next') {
    moveSemiAssignmentIndex(1);
  } else if (action === 'accept') {
    applySemiAssignmentAcceptance();
  } else if (action === 'skip') {
    skipSemiAssignmentStep();
  }
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

const buildValidatedGuardCounts = () => {
  const map = new Map();
  state.requests.forEach((request) => {
    if (request.status !== 'validé') {
      return;
    }
    const trigram = normalizeTrigram(request.trigram);
    if (!trigram) {
      return;
    }
    if (!map.has(trigram)) {
      map.set(trigram, { normal: 0, good: 0 });
    }
    const entry = map.get(trigram);
    if (request.guardNature === 'bonne') {
      entry.good += 1;
    } else if (request.guardNature === 'normale') {
      entry.normal += 1;
    }
  });
  return map;
};

const findNextAvailableRequest = (
  list,
  trigram,
  assignedSlots,
  occupiedSlots,
  extraOccupied = new Set(),
  options = {}
) => {
  if (!Array.isArray(list) || !list.length) {
    return null;
  }
  const { usedGroups = null, guardType = '' } = options;
  for (let index = 0; index < list.length; index += 1) {
    const candidate = list[index];
    if (!candidate) {
      continue;
    }
    if (candidate.status && candidate.status !== 'en attente') {
      continue;
    }
    const groupKey = getRequestGroupKey(candidate);
    const stateKey = buildGroupStateKey(trigram, guardType, groupKey);
    if (stateKey && usedGroups?.has(stateKey)) {
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
      slotKey,
      groupKey
    };
  }
  return null;
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
  if (elements.semiForm) {
    elements.semiForm.addEventListener('change', handleSemiAssignmentFormChange);
    elements.semiForm.addEventListener('click', handleSemiAssignmentAction);
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
  await loadSemiAssignmentDirectory();
  updateSemiAssignmentActiveTour();
  updateSemiAssignmentTrigramOptions();
};

initialize().catch((error) => {
  console.error('Failed to initialize request processing page', error);
  setFeedback('Une erreur est survenue lors du chargement de la page.');
});
