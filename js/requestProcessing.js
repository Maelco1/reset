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
  }
};

const elements = {
  logoutBtn: document.querySelector('#logout'),
  disconnectBtn: document.querySelector('#disconnect'),
  backBtn: document.querySelector('#back-to-admin'),
  userTypeTabsNav: document.querySelector('#request-user-tabs'),
  tabsNav: document.querySelector('#request-tabs'),
  filtersForm: document.querySelector('#request-filters'),
  feedback: document.querySelector('#request-feedback'),
  tableBody: document.querySelector('#requests-body'),
  userTypeTabButtons: Array.from(document.querySelectorAll('[data-user-type-tab]'))
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

const NEARLY_EQUAL_EPSILON = 1e-6;

const areNearlyEqual = (a, b, epsilon = NEARLY_EQUAL_EPSILON) => {
  if (!Number.isFinite(a) || !Number.isFinite(b)) {
    return a === b;
  }
  return Math.abs(a - b) <= epsilon;
};

const getChoiceIndexParts = (rawValue) => {
  const numeric = parseNumeric(rawValue);
  if (!Number.isFinite(numeric)) {
    return {
      numeric: null,
      primary: Number.POSITIVE_INFINITY,
      secondary: Number.POSITIVE_INFINITY
    };
  }
  const primary = Math.trunc(numeric);
  const secondary = Math.abs(numeric - primary);
  return {
    numeric,
    primary,
    secondary
  };
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

const hasHigherSecondaryIndexWithinPrimary = (candidate, selectedIndex) => {
  if (selectedIndex == null) {
    return false;
  }
  const candidateParts = getChoiceIndexParts(candidate.choice_index ?? candidate.choiceIndex);
  const selectedParts = getChoiceIndexParts(selectedIndex);
  if (!Number.isFinite(candidateParts.primary) || !Number.isFinite(selectedParts.primary)) {
    return false;
  }
  if (candidateParts.primary !== selectedParts.primary) {
    return false;
  }
  if (!Number.isFinite(candidateParts.secondary) || !Number.isFinite(selectedParts.secondary)) {
    return false;
  }
  if (areNearlyEqual(candidateParts.secondary, selectedParts.secondary)) {
    return false;
  }
  return candidateParts.secondary > selectedParts.secondary;
};

const getAlternativeRelation = (candidate, selectedRank, selectedIndex, selectedOrder) => {
  if (haveMatchingChoiceOrder(candidate, selectedOrder)) {
    return 'order';
  }
  if (haveMatchingChoiceIndex(candidate, selectedIndex)) {
    return 'index';
  }
  if (hasHigherSecondaryIndexWithinPrimary(candidate, selectedIndex)) {
    return 'secondary';
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

// Remplacer la fonction existante
function getRequestGroupKey(request) {
  // On privilégie l'INDEX PRINCIPAL (partie entière de 1.x) pour regrouper les alternatives
  const indexRaw = request?.choiceIndex ?? request?.choice_index;
  const indexNum = parseNumeric(indexRaw);
  if (Number.isFinite(indexNum)) {
    // Math.trunc() garantit que 1.0, 1.2, 1.9 => groupe "1"
    return `index:${Math.trunc(indexNum)}`;
  }

  // Si pas d'index, on retombe sur l'ordre
  const orderRaw = request?.choiceOrder ?? request?.choice_order;
  const orderNum = parseNumeric(orderRaw);
  if (Number.isFinite(orderNum)) {
    return `order:${orderNum}`;
  }

  // Fallback stable pour éviter des regroupements incohérents
  const id = request?.id ?? request?.request_id ?? JSON.stringify(request);
  return `misc:${String(id)}`;
}

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
  updateUserTypeTabs();
  buildTabs();
  updateTabState();

  await ensureSupabase();
  await ensureActor();
  await loadAdministrativeSettings();
  await loadPlanningColumns();
  await loadRequests();
};

initialize().catch((error) => {
  console.error('Failed to initialize request processing page', error);
  setFeedback('Une erreur est survenue lors du chargement de la page.');
});
