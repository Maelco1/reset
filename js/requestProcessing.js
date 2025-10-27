import {
  initializeConnectionModal,
  requireRole,
  getCurrentUser,
  setCurrentUser,
  getSupabaseClient,
  onSupabaseReady
} from './supabaseClient.js';

const PLANNING_TOURS = [
  { id: 1, table: 'planning_columns' },
  { id: 2, table: 'planning_columns_tour2' },
  { id: 3, table: 'planning_columns_tour3' },
  { id: 4, table: 'planning_columns_tour4' },
  { id: 5, table: 'planning_columns_tour5' },
  { id: 6, table: 'planning_columns_tour6' }
];

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
  tabsNav: document.querySelector('#request-tabs'),
  filtersForm: document.querySelector('#request-filters'),
  feedback: document.querySelector('#request-feedback'),
  tableBody: document.querySelector('#requests-body'),
  userTypeButtons: Array.from(document.querySelectorAll('[data-user-type-filter]'))
};

const normalizeUserTypeFilter = (value) => {
  if (typeof value !== 'string') {
    return '';
  }
  const lowered = value.toLowerCase();
  return USER_TYPE_FILTERS.has(lowered) ? lowered : '';
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

const getPlanningTableName = (tourId = state.activeTourId) => getTourConfig(tourId).table;

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
  if (!Number.isFinite(index)) {
    return '—';
  }
  if (!Number.isFinite(rank) || rank <= 1) {
    return String(index);
  }
  return `${index}.${rank}`;
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
      const columnLabel = `${request.planningDayLabel ?? ''} ${request.columnLabel ?? ''} ${request.slotTypeCode ?? ''}`.toLowerCase();
      if (!columnLabel.includes(filters.column.toLowerCase())) {
        return false;
      }
    }
    return true;
  });
};

const sortRequests = (list) =>
  [...list].sort((a, b) => {
    const indexA = Number.isFinite(a.choiceIndex) ? a.choiceIndex : Number.MAX_SAFE_INTEGER;
    const indexB = Number.isFinite(b.choiceIndex) ? b.choiceIndex : Number.MAX_SAFE_INTEGER;
    if (indexA !== indexB) {
      return indexA - indexB;
    }
    const rankA = Number.isFinite(a.choiceRank) ? a.choiceRank : Number.MAX_SAFE_INTEGER;
    const rankB = Number.isFinite(b.choiceRank) ? b.choiceRank : Number.MAX_SAFE_INTEGER;
    if (rankA !== rankB) {
      return rankA - rankB;
    }
    const timeA = a.createdAt?.getTime?.() ?? new Date(a.createdAt ?? 0).getTime();
    const timeB = b.createdAt?.getTime?.() ?? new Date(b.createdAt ?? 0).getTime();
    return timeA - timeB;
  });

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
    cell.colSpan = 10;
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
    addCell(request.planningDayLabel || request.columnLabel || `Colonne ${request.columnNumber}`, 'column');
    addCell(TYPE_LABELS.get(request.activityType) ?? TYPE_LABELS.get((request.activityType ?? '').toLowerCase()) ?? 'Visite', 'type');
    addCell(formatHours(request.columnNumber), 'hours');
    addCell((request.trigram ?? '').toUpperCase(), 'doctor');
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
  const tableName = getPlanningTableName();
  const { data, error } = await supabase
    .from(tableName)
    .select('position, label, type_category, start_time, end_time');
  if (error) {
    console.error(error);
    return;
  }
  state.columns = new Map((data ?? []).map((column) => [column.position, column]));
};

const mapRequestRecord = (record) => {
  const createdAt = record.created_at ? new Date(record.created_at) : null;
  return {
    id: record.id,
    trigram: (record.trigram ?? '').toUpperCase(),
    userType: normalizeUserTypeFilter(record.user_type ?? ''),
    day: record.day ?? null,
    columnNumber: record.column_number ?? null,
    columnLabel: record.column_label ?? null,
    planningDayLabel: record.planning_day_label ?? null,
    slotTypeCode: record.slot_type_code ?? null,
    guardNature: record.guard_nature ?? 'normale',
    activityType: (record.activity_type ?? 'visite').toLowerCase(),
    choiceIndex: Number.parseFloat(record.choice_index ?? record.choiceIndex ?? 0) || 0,
    choiceRank: Number.parseFloat(record.choice_rank ?? record.choiceRank ?? 1) || 1,
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
      'id, trigram, user_type, day, column_number, column_label, planning_day_label, slot_type_code, guard_nature, activity_type, choice_index, choice_rank, created_at, etat, is_active, planning_reference, tour_number'
    );
  if (state.planningReference) {
    query.eq('planning_reference', state.planningReference);
  }
  if (state.activeTourId) {
    query.eq('tour_number', state.activeTourId);
  }
  query.order('choice_index', { ascending: true }).order('choice_rank', { ascending: true }).order('created_at', { ascending: true });
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
    return;
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
    return;
  }
  const { error: updateError } = await supabase
    .from(CHOICES_TABLE)
    .update({ choice_rank: 1 })
    .eq('id', alternative.id);
  if (updateError) {
    console.error(updateError);
    return;
  }
  await recordAudit({
    action: 'auto_promote',
    choiceId: alternative.id,
    targetTrigram: trigram,
    targetDay: alternative.day,
    targetColumnNumber: alternative.column_number,
    reason: "Promotion automatique de l'alternative"
  });
};

const fetchCompetingRequests = async (request) => {
  const supabase = await ensureSupabase();
  if (!supabase) {
    return [];
  }
  const { data, error } = await supabase
    .from(CHOICES_TABLE)
    .select('id, trigram, choice_index, choice_rank, etat, day, column_number')
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

  const alternativeIds = competing
    .filter((item) => item.trigram === request.trigram && item.id !== request.id && item.choice_rank > 1)
    .map((item) => item.id);
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
        .update({ is_active: false })
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

const updateUserTypeFilterButtons = () => {
  if (!Array.isArray(elements.userTypeButtons)) {
    return;
  }
  const active = normalizeUserTypeFilter(state.filters.userType);
  elements.userTypeButtons.forEach((button) => {
    if (!button) {
      return;
    }
    const value = normalizeUserTypeFilter(button.dataset.userTypeFilter ?? '');
    const isActive = !!value && value === active;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
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
  updateUserTypeFilterButtons();
  renderRequests();
};

const handleUserTypeFilterClick = (event) => {
  const button = event.currentTarget ?? event.target;
  if (!button) {
    return;
  }
  const value = button.dataset.userTypeFilter;
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
  updateUserTypeFilterButtons();
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
        updateUserTypeFilterButtons();
        renderRequests();
      }, 0);
    });
  }
  if (Array.isArray(elements.userTypeButtons)) {
    elements.userTypeButtons.forEach((button) => {
      if (button) {
        button.addEventListener('click', handleUserTypeFilterClick);
      }
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
  updateUserTypeFilterButtons();
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
