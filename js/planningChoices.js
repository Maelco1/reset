import {
  getSupabaseClient,
  onSupabaseReady,
  getCurrentUser
} from './supabaseClient.js';

const DEFAULT_COLOR = '#1e293b';
const DEFAULT_COLUMN_TINT = 'rgba(30, 41, 59, 0.12)';
const COLUMN_DEFAULTS = [
  { position: 1, typeCode: '1N', color: '#1e293b' },
  { position: 2, typeCode: '2N', color: '#1e293b' },
  { position: 3, typeCode: '3N', color: '#1e293b' },
  { position: 4, typeCode: '4C', color: '#1e293b' },
  { position: 5, typeCode: '5S', color: '#1e293b' },
  { position: 6, typeCode: '6S', color: '#1e293b' },
  { position: 7, typeCode: 'VIS', color: '#1e293b' },
  { position: 8, typeCode: 'VIS', color: '#1e293b' },
  { position: 9, typeCode: 'VIS', color: '#1e293b' },
  { position: 10, typeCode: 'VIS', color: '#1e293b' },
  { position: 11, typeCode: 'VIS', color: '#1e293b' },
  { position: 12, typeCode: 'TC', color: '#1e293b' },
  { position: 13, typeCode: 'C1COU', color: '#1e293b' },
  { position: 14, typeCode: 'C2COU', color: '#1e293b' },
  { position: 15, typeCode: 'C1BOU', color: '#1e293b' },
  { position: 16, typeCode: 'C2BOU', color: '#1e293b' },
  { position: 17, typeCode: 'PFG', color: '#1e293b' },
  { position: 18, typeCode: 'C1ANT', color: '#1e293b' },
  { position: 19, typeCode: 'C2ANT', color: '#1e293b' },
  { position: 20, typeCode: 'TC', color: '#1e293b' },
  { position: 21, typeCode: 'N', color: '#1e293b' },
  { position: 22, typeCode: 'C', color: '#1e293b' },
  { position: 23, typeCode: 'S', color: '#1e293b' },
  { position: 24, typeCode: 'VIS', color: '#1e293b' },
  { position: 25, typeCode: 'VIS', color: '#1e293b' },
  { position: 26, typeCode: 'VIS', color: '#1e293b' },
  { position: 27, typeCode: 'C1COU', color: '#1e293b' },
  { position: 28, typeCode: 'C2COU', color: '#1e293b' },
  { position: 29, typeCode: 'C1BOU', color: '#1e293b' },
  { position: 30, typeCode: 'C2BOU', color: '#1e293b' },
  { position: 31, typeCode: 'PFG', color: '#1e293b' },
  { position: 32, typeCode: 'C1ANT', color: '#1e293b' },
  { position: 33, typeCode: 'C2ANT', color: '#1e293b' },
  { position: 34, typeCode: 'TC', color: '#1e293b' },
  { position: 35, typeCode: 'VIS', color: '#1e293b' },
  { position: 36, typeCode: 'PFG', color: '#1e293b' },
  { position: 37, typeCode: 'VIS', color: '#1e293b' },
  { position: 38, typeCode: 'VIS', color: '#1e293b' },
  { position: 39, typeCode: 'VIS', color: '#1e293b' },
  { position: 40, typeCode: 'VIS', color: '#1e293b' },
  { position: 41, typeCode: 'VIS', color: '#1e293b' },
  { position: 42, typeCode: 'VIS', color: '#1e293b' },
  { position: 43, typeCode: 'TCN', color: '#1e293b' },
  { position: 44, typeCode: 'VIS', color: '#1e293b' },
  { position: 45, typeCode: 'VIS', color: '#1e293b' },
  { position: 46, typeCode: 'VIS', color: '#1e293b' }
];

const TYPE_CATEGORY_MAP = new Map([
  ...[
    '1N',
    '2N',
    '3N',
    '4C',
    '5S',
    '6S',
    'N',
    'C',
    'S',
    'VIS'
  ].map((code) => [code, 'Visite']),
  ...[
    'C1COU',
    'C2COU',
    'C1BOU',
    'C2BOU',
    'PFG',
    'C1ANT',
    'C2ANT',
    'C1',
    'C2',
    'C3'
  ].map((code) => [code, 'Consultation']),
  ...['TC', 'TCN'].map((code) => [code, 'Téléconsultation'])
]);

const PLANNING_TOURS = [
  { id: 1, label: 'Tour 1', table: 'planning_columns' },
  { id: 2, label: 'Tour 2', table: 'planning_columns_tour2' },
  { id: 3, label: 'Tour 3', table: 'planning_columns_tour3' },
  { id: 4, label: 'Tour 4', table: 'planning_columns_tour4' },
  { id: 5, label: 'Tour 5', table: 'planning_columns_tour5' },
  { id: 6, label: 'Tour 6', table: 'planning_columns_tour6' }
];

const ADMIN_SETTINGS_TABLE = 'parametres_administratifs';
const WEEKDAY_LABELS = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
const MONTH_NAMES = [
  'Janvier',
  'Février',
  'Mars',
  'Avril',
  'Mai',
  'Juin',
  'Juillet',
  'Août',
  'Septembre',
  'Octobre',
  'Novembre',
  'Décembre'
];

const ACTIVITY_TYPES = new Map([
  ['Visite', 'visite'],
  ['Consultation', 'consultation'],
  ['Téléconsultation', 'téléconsultation']
]);

const USER_TYPE_LABELS = new Set(['medecin', 'remplacant']);
const CHOICE_SERIES = ['mauvaise', 'bonus'];
const CHOICE_SERIES_LABELS = new Map([
  ['mauvaise', 'Mauvaises gardes'],
  ['bonus', 'Gardes bonus']
]);
const CHOICE_SERIES_STEPS = new Map([
  ['mauvaise', 1],
  ['bonus', 2]
]);
const CHOICE_INDEX_MIN = 1;
const CHOICE_INDEX_MAX = 20;

const clamp = (value, min, max) => Math.min(Math.max(value, min), max);

const sanitizeColor = (value) => {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  if (/^#[0-9a-fA-F]{6}$/.test(trimmed)) {
    return trimmed.toLowerCase();
  }
  if (/^#[0-9a-fA-F]{3}$/.test(trimmed)) {
    const [r, g, b] = trimmed
      .slice(1)
      .split('')
      .map((char) => char + char);
    return `#${r}${g}${b}`.toLowerCase();
  }
  return null;
};

const normalizeColor = (value) => sanitizeColor(value) ?? DEFAULT_COLOR;

const hexToRgb = (value) => {
  const sanitized = sanitizeColor(value);
  if (!sanitized) {
    return null;
  }
  const hex = sanitized.slice(1);
  const numeric = Number.parseInt(hex, 16);
  if (Number.isNaN(numeric)) {
    return null;
  }
  return {
    r: (numeric >> 16) & 255,
    g: (numeric >> 8) & 255,
    b: numeric & 255
  };
};

const getSlotColumnTint = (color) => {
  const rgb = hexToRgb(color);
  if (!rgb) {
    return DEFAULT_COLUMN_TINT;
  }
  const alpha = 0.14;
  return `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${alpha})`;
};

const getRelativeLuminance = ({ r, g, b }) => {
  const transform = (channel) => {
    const scaled = channel / 255;
    return scaled <= 0.03928 ? scaled / 12.92 : Math.pow((scaled + 0.055) / 1.055, 2.4);
  };
  const red = transform(r);
  const green = transform(g);
  const blue = transform(b);
  return 0.2126 * red + 0.7152 * green + 0.0722 * blue;
};

const getSlotTextColor = (color) => {
  const rgb = hexToRgb(color);
  if (!rgb) {
    return '#1f2937';
  }
  const luminance = getRelativeLuminance(rgb);
  return luminance > 0.55 ? '#0f172a' : '#f8fafc';
};

const inferTypeCategory = (typeCode) => {
  if (typeof typeCode !== 'string') {
    return 'Visite';
  }
  const normalized = typeCode.trim().toUpperCase();
  return TYPE_CATEGORY_MAP.get(normalized) ?? 'Visite';
};

const sanitizeQuality = (value) => (value === 'Bonus' ? 'Bonus' : 'Mauvaise');

const toBoolean = (value, fallback = true) => {
  if (value === null || value === undefined) {
    return fallback;
  }
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  const str = String(value).toLowerCase();
  return str === 'true' || str === '1' || str === 'yes' || str === 'oui';
};

const getSlotTimeRange = (slot) => {
  const start = (slot.start_time ?? '').trim();
  const end = (slot.end_time ?? '').trim();
  if (start && end) {
    return `${start} – ${end}`;
  }
  if (start) {
    return `Début ${start}`;
  }
  if (end) {
    return `Fin ${end}`;
  }
  return '';
};

const getDefaultSlot = (position) => {
  const defaults = COLUMN_DEFAULTS.find((item) => item.position === position);
  const typeCode = defaults?.typeCode ?? `COL${String(position).padStart(2, '0')}`;
  const category = inferTypeCategory(typeCode);
  return {
    position,
    label: typeCode,
    type_code: typeCode,
    type_category: category,
    start_time: null,
    end_time: null,
    color: defaults?.color ?? DEFAULT_COLOR,
    quality_weekdays: 'Mauvaise',
    quality_saturday: 'Mauvaise',
    quality_sunday: 'Mauvaise',
    open_mauvaise_weekdays: true,
    open_mauvaise_saturday: true,
    open_mauvaise_sunday: true,
    open_bonus_weekdays: true,
    open_bonus_saturday: true,
    open_bonus_sunday: true
  };
};

const sortSlotsByPosition = (slots) =>
  [...slots].sort((a, b) => {
    const positionA = Number(a?.position ?? Number.POSITIVE_INFINITY);
    const positionB = Number(b?.position ?? Number.POSITIVE_INFINITY);
    return positionA - positionB;
  });

const addDays = (date, days) => {
  const clone = new Date(date);
  clone.setDate(clone.getDate() + days);
  return clone;
};

const formatDateKey = (date) =>
  `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;

const calculateEaster = (year) => {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
};

const getFrenchHolidays = (year) => {
  const holidays = new Map();
  const addFixed = (month, day, label) => {
    holidays.set(formatDateKey(new Date(year, month, day)), label);
  };
  addFixed(0, 1, "Jour de l'An");
  addFixed(4, 1, 'Fête du Travail');
  addFixed(4, 8, 'Victoire 1945');
  addFixed(6, 14, 'Fête Nationale');
  addFixed(7, 15, 'Assomption');
  addFixed(10, 1, 'Toussaint');
  addFixed(10, 11, 'Armistice');
  addFixed(11, 25, 'Noël');

  const easter = calculateEaster(year);
  holidays.set(formatDateKey(addDays(easter, 1)), 'Lundi de Pâques');
  holidays.set(formatDateKey(addDays(easter, 39)), 'Ascension');
  holidays.set(formatDateKey(addDays(easter, 50)), 'Lundi de Pentecôte');

  return holidays;
};

const getMonthDays = (year, month) => {
  const days = [];
  const cursor = new Date(year, month, 1);
  while (cursor.getMonth() === month) {
    days.push(new Date(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  return days;
};

const getDaySegment = (date, isHoliday) => {
  if (isHoliday || date.getDay() === 0) {
    return 'sunday';
  }
  if (date.getDay() === 6) {
    return 'saturday';
  }
  return 'weekday';
};

const getQualityForSegment = (slot, segment) => {
  if (segment === 'saturday') {
    return sanitizeQuality(slot.quality_saturday);
  }
  if (segment === 'sunday') {
    return sanitizeQuality(slot.quality_sunday);
  }
  return sanitizeQuality(slot.quality_weekdays);
};

const getSlotOpeningsForSegment = (slot, segment) => {
  const segmentKey = segment === 'saturday' ? 'saturday' : segment === 'sunday' ? 'sunday' : 'weekdays';
  const bonusKey = `open_bonus_${segmentKey}`;
  const mauvaiseKey = `open_mauvaise_${segmentKey}`;
  return {
    bonus: toBoolean(slot[bonusKey]),
    mauvaise: toBoolean(slot[mauvaiseKey])
  };
};

const isSlotOpen = (slot, segment, visionOverride = null) => {
  const openings = getSlotOpeningsForSegment(slot, segment);
  if (visionOverride === 'bonus' || visionOverride === 'mauvaise') {
    return openings[visionOverride];
  }
  if (visionOverride === null) {
    return openings.bonus || openings.mauvaise;
  }
  const quality = getQualityForSegment(slot, segment);
  const preferredVision = quality === 'Bonus' ? 'bonus' : 'mauvaise';
  return openings[preferredVision];
};

const buildSlotTitle = (slot, dayLabel, isHoliday, holidayName, isOpen, quality) => {
  const details = [];
  const label = slot.label ?? `Créneau ${slot.position}`;
  details.push(label);
  if (slot.type_code) {
    details.push(`Type : ${slot.type_code}`);
  }
  if (slot.type_category) {
    details.push(slot.type_category);
  }
  const slotTime = getSlotTimeRange(slot);
  if (slotTime) {
    details.push(slotTime);
  }
  details.push(`Qualité : ${quality}`);
  details.push(isOpen ? 'Ouvert' : 'Indisponible');
  if (isHoliday && holidayName) {
    details.push(holidayName);
  }
  return `${dayLabel} · ${details.join(' · ')}`;
};

const getTourConfig = (tourId) =>
  PLANNING_TOURS.find((tour) => tour.id === tourId) ?? PLANNING_TOURS[0];

const getPlanningTableName = (tourId) => getTourConfig(tourId).table;

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

const ensurePlanningColumns = async (supabase, slots, tableName) => {
  const missing = [];
  for (let position = 1; position <= 46; position += 1) {
    if (!slots.some((slot) => slot.position === position)) {
      missing.push(getDefaultSlot(position));
    }
  }
  if (missing.length === 0) {
    return slots;
  }
  const { error } = await supabase.from(tableName).upsert(missing, { onConflict: 'position' });
  if (error) {
    console.error(error);
    return slots;
  }
  const { data, error: reloadError } = await supabase
    .from(tableName)
    .select(
      `id,
        position,
        label,
        type_code,
        type_category,
        start_time,
        end_time,
        color,
        quality_weekdays,
        quality_saturday,
        quality_sunday,
        open_mauvaise_weekdays,
        open_mauvaise_saturday,
        open_mauvaise_sunday,
        open_bonus_weekdays,
        open_bonus_saturday,
        open_bonus_sunday`
    )
    .order('position');
  if (reloadError) {
    console.error(reloadError);
    return slots;
  }
  return data ?? slots;
};

const getActivityType = (category) => ACTIVITY_TYPES.get(category) ?? 'visite';

const createSelectionFromButton = (button, mode) => {
  const dayIso = button.dataset.day;
  const dayDate = new Date(dayIso);
  const monthNumber = Number.parseInt(button.dataset.month, 10);
  const yearNumber = Number.parseInt(button.dataset.year, 10);
  return {
    slotKey: button.dataset.slotKey,
    nature: mode,
    dayIso,
    dayDate,
    dayLabel: button.dataset.dayLabel ?? dayIso,
    monthNumber,
    yearNumber,
    columnPosition: Number.parseInt(button.dataset.position, 10),
    columnLabel: button.dataset.columnLabel ?? '',
    slotTypeCode: button.dataset.typeCode ?? '',
    activityCategory: button.dataset.activityCategory ?? 'Visite',
    activityType: getActivityType(button.dataset.activityCategory ?? 'Visite'),
    summary: button.dataset.summary ?? '',
    button
  };
};

const getPlanningReference = ({ tourId, year, monthOne, monthTwo }) => {
  const toPart = (month) => String(month + 1).padStart(2, '0');
  return `tour${tourId}-${year}-${toPart(monthOne)}-${toPart(monthTwo)}`;
};

const formatSummaryLabel = (selection) => {
  const monthName = MONTH_NAMES[selection.dayDate.getMonth()];
  const dayNumber = String(selection.dayDate.getDate()).padStart(2, '0');
  const weekday = WEEKDAY_LABELS[selection.dayDate.getDay()];
  return `${weekday} ${dayNumber} ${monthName} ${selection.dayDate.getFullYear()}`;
};

const DRAG_DATA_TYPE = 'text/plain';
export function initializePlanningChoices({ userRole }) {
  const normalizedUserRole = USER_TYPE_LABELS.has(userRole) ? userRole : 'medecin';
  const planningSection = document.querySelector('[data-planning-section]');
  const planningTables = document.querySelector('#planning-tables');
  const planningFeedback = document.querySelector('#planning-feedback');
  const planningBoard = document.querySelector('#planning-board');
  const stepper = document.querySelector('.stepper');
  const stepperButtons = stepper ? Array.from(stepper.querySelectorAll('.stepper-step')) : [];
  const stepPanes = Array.from(document.querySelectorAll('[data-step-pane]'));
  const stepHosts = new Map(
    Array.from(document.querySelectorAll('[data-step-host]')).map((element) => [
      Number.parseInt(element.dataset.stepHost, 10),
      element
    ])
  );
  const summaryList = document.querySelector('#summary-list');
  const summaryFeedback = document.querySelector('#summary-feedback');
  const saveButton = document.querySelector('#save-choices');
  const saveFeedback = document.querySelector('#save-feedback');

  if (!planningSection || !planningTables || !planningBoard) {
    console.error('Planning markup missing from page.');
    return;
  }

  const state = {
    supabase: null,
    planningSlots: [],
    slotButtons: new Map(),
    selectionMap: new Map(),
    selections: [],
    selectionMode: 'mauvaise',
    choiceSeries: {
      mauvaise: { activeIndex: CHOICE_INDEX_MIN, buttons: new Map(), container: null },
      bonus: { activeIndex: CHOICE_INDEX_MIN, buttons: new Map(), container: null }
    },
    currentStep: 1,
    selectedTourId: PLANNING_TOURS[0].id,
    planningYear: new Date().getFullYear(),
    planningMonthOne: new Date().getMonth(),
    planningMonthTwo: (new Date().getMonth() + 1) % 12,
    planningReference: getPlanningReference({
      tourId: PLANNING_TOURS[0].id,
      year: new Date().getFullYear(),
      monthOne: new Date().getMonth(),
      monthTwo: (new Date().getMonth() + 1) % 12
    }),
    holidaysByYear: new Map(),
    planningCellSizingFrame: null,
    stickyHeaderFrame: null,
    userProfile: null,
    isSaving: false
  };

  const getHolidaysForYear = (year) => {
    if (!state.holidaysByYear.has(year)) {
      state.holidaysByYear.set(year, getFrenchHolidays(year));
    }
    return state.holidaysByYear.get(year);
  };

  const setPlanningFeedback = (message) => {
    if (planningFeedback) {
      planningFeedback.textContent = message ?? '';
    }
  };

  const movePlanningSectionToStep = (step) => {
    const host = stepHosts.get(step);
    if (!host || planningSection.parentElement === host) {
      return;
    }
    host.appendChild(planningSection);
  };

  const sanitizeChoiceNature = (nature) => (nature === 'bonus' ? 'bonus' : 'mauvaise');

  const sanitizeChoiceIndex = (value) => {
    if (value == null || value === '') {
      return CHOICE_INDEX_MIN;
    }
    const numeric = Number.parseInt(value, 10);
    if (Number.isNaN(numeric)) {
      return CHOICE_INDEX_MIN;
    }
    return clamp(numeric, CHOICE_INDEX_MIN, CHOICE_INDEX_MAX);
  };

  const sanitizeChoiceRank = (value) => {
    if (value == null || value === '') {
      return 1;
    }
    const numeric = Number.parseInt(value, 10);
    if (!Number.isFinite(numeric) || numeric < 1) {
      return 1;
    }
    return numeric;
  };

  const getChoiceSeries = (nature) => state.choiceSeries[sanitizeChoiceNature(nature)];

  const getActiveChoiceIndex = (nature) => {
    const series = getChoiceSeries(nature);
    return sanitizeChoiceIndex(series?.activeIndex);
  };

  const updateChoiceIndexActiveState = (nature) => {
    const series = getChoiceSeries(nature);
    if (!series) {
      return;
    }
    const activeIndex = sanitizeChoiceIndex(series.activeIndex);
    series.buttons.forEach((button, key) => {
      const indexNumber = sanitizeChoiceIndex(key);
      const isActive = indexNumber === activeIndex;
      button.classList.toggle('is-active', isActive);
      button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
    });
  };

  const setActiveChoiceIndex = (nature, index) => {
    const series = getChoiceSeries(nature);
    if (!series) {
      return;
    }
    const sanitizedIndex = sanitizeChoiceIndex(index);
    if (series.activeIndex === sanitizedIndex) {
      updateChoiceIndexActiveState(nature);
      return;
    }
    series.activeIndex = sanitizedIndex;
    updateChoiceIndexActiveState(nature);
  };

  const createEmptyChoiceGroups = () =>
    CHOICE_SERIES.reduce((acc, nature) => {
      acc[nature] = new Map();
      return acc;
    }, {});

  const updateChoiceIndexButtons = (groupedSelections = null) => {
    const groups = groupedSelections ?? createEmptyChoiceGroups();
    CHOICE_SERIES.forEach((nature) => {
      const series = getChoiceSeries(nature);
      if (!series) {
        return;
      }
      const natureGroups = groups[nature] ?? new Map();
      if (!groups[nature]) {
        groups[nature] = natureGroups;
      }
      for (let index = CHOICE_INDEX_MIN; index <= CHOICE_INDEX_MAX; index += 1) {
        const button = series.buttons.get(index);
        if (!button) {
          continue;
        }
        const entries = natureGroups.get(index) ?? [];
        const hasPrimary = entries.length > 0;
        const alternativeCount = Math.max(0, entries.length - 1);
        button.classList.toggle('has-selection', hasPrimary || alternativeCount > 0);
        button.dataset.hasPrimary = hasPrimary ? 'true' : 'false';
        button.dataset.alternativeCount = String(alternativeCount);
        const indicator = button.querySelector('.choice-index-indicator');
        if (indicator) {
          if (!hasPrimary && alternativeCount === 0) {
            indicator.textContent = '';
          } else if (alternativeCount <= 3) {
            const alternatives = alternativeCount > 0 ? '○'.repeat(alternativeCount) : '';
            indicator.textContent = `${hasPrimary ? '●' : ''}${alternatives}`;
          } else {
            indicator.textContent = `${hasPrimary ? '● ' : ''}${alternativeCount}×○`;
          }
        }
        const sr = button.querySelector('.sr-only');
        if (sr) {
          const parts = [`Choix ${index}`];
          if (!hasPrimary && alternativeCount === 0) {
            parts.push('aucun créneau sélectionné');
          } else {
            parts.push(hasPrimary ? 'principal sélectionné' : 'aucun principal');
            parts.push(
              alternativeCount > 0
                ? `${alternativeCount} alternative${alternativeCount > 1 ? 's' : ''}`
                : 'aucune alternative'
            );
          }
          sr.textContent = parts.join(', ');
        }
      }
      updateChoiceIndexActiveState(nature);
    });
  };

  const createChoiceIndexControls = (nature) => {
    const series = getChoiceSeries(nature);
    if (!series) {
      return null;
    }
    series.buttons = new Map();
    const container = document.createElement('section');
    container.className = 'choice-index-panel';
    container.dataset.choiceNature = nature;

    const title = document.createElement('h3');
    title.className = 'choice-index-title';
    title.textContent = `Numérotation des choix — ${CHOICE_SERIES_LABELS.get(nature) ?? nature}`;
    container.appendChild(title);

    const description = document.createElement('p');
    description.className = 'choice-index-description';
    description.textContent = "Choisissez le numéro de choix actif puis cliquez sur les créneaux pour l'associer.";
    container.appendChild(description);

    const grid = document.createElement('div');
    grid.className = 'choice-index-grid';
    grid.setAttribute(
      'aria-label',
      `Sélection du numéro de choix pour ${(
        CHOICE_SERIES_LABELS.get(nature) ?? nature
      ).toLowerCase()}`
    );
    grid.setAttribute('role', 'group');

    for (let index = CHOICE_INDEX_MIN; index <= CHOICE_INDEX_MAX; index += 1) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'choice-index-button';
      button.dataset.choiceIndex = String(index);
      button.innerHTML = `
        <span class="choice-index-number">${index}</span>
        <span class="choice-index-indicator" aria-hidden="true"></span>
        <span class="sr-only">Choix ${index}</span>
      `;
      grid.appendChild(button);
      series.buttons.set(index, button);
    }

    grid.addEventListener('click', (event) => {
      const button = event.target.closest('.choice-index-button');
      if (!button || !grid.contains(button)) {
        return;
      }
      const index = Number.parseInt(button.dataset.choiceIndex ?? '', 10);
      if (Number.isNaN(index)) {
        return;
      }
      setActiveChoiceIndex(nature, index);
    });

    container.appendChild(grid);
    return container;
  };

  const initializeChoiceSelectors = () => {
    CHOICE_SERIES.forEach((nature) => {
      const step = CHOICE_SERIES_STEPS.get(nature);
      const host = stepHosts.get(step);
      if (!host) {
        return;
      }
      const container = createChoiceIndexControls(nature);
      if (!container) {
        return;
      }
      state.choiceSeries[sanitizeChoiceNature(nature)].container = container;
      host.prepend(container);
      updateChoiceIndexActiveState(nature);
    });
    updateChoiceIndexButtons();
  };

  const updateStepperButtons = () => {
    stepperButtons.forEach((button) => {
      const step = Number.parseInt(button.dataset.step, 10);
      const isActive = step === state.currentStep;
      button.classList.toggle('is-active', isActive);
      button.setAttribute('aria-current', isActive ? 'step' : 'false');
    });
  };

  const updateStepPanes = () => {
    stepPanes.forEach((pane) => {
      const step = Number.parseInt(pane.dataset.stepPane, 10);
      const isActive = step === state.currentStep;
      pane.classList.toggle('is-active', isActive);
    });
  };

  const updateSummaryFeedback = () => {
    if (!summaryFeedback) {
      return;
    }
    const total = state.selections.length;
    summaryFeedback.textContent = total
      ? `${total} sélection${total > 1 ? 's' : ''}`
      : 'Aucune sélection pour le moment.';
  };

  const updateSelectionVisuals = () => {
    const groupedSelections = createEmptyChoiceGroups();

    state.slotButtons.forEach((button) => {
      const tag = button.querySelector('.planning-assignment-tag');
      const roleLabel = button.querySelector('.planning-assignment-role');
      const srLabel = button.querySelector('.sr-only');
      if (tag) {
        tag.textContent = '';
        tag.classList.add('is-empty');
      }
      if (roleLabel) {
        roleLabel.textContent = '';
        roleLabel.classList.add('is-empty');
      }
      if (srLabel) {
        srLabel.textContent = button.dataset.summary ?? '';
      }
      button.removeAttribute('data-choice-order');
      button.removeAttribute('data-choice-type');
      button.removeAttribute('data-choice-rank');
      button.removeAttribute('data-choice-index');
      button.removeAttribute('data-choice-role');
      button.removeAttribute('data-choice-label');
      button.classList.remove('is-selected');
    });

    state.selections.forEach((selection, index) => {
      selection.order = index + 1;
      const nature = sanitizeChoiceNature(selection.nature);
      selection.nature = nature;
      const choiceIndex = selection.choiceIndex
        ? sanitizeChoiceIndex(selection.choiceIndex)
        : getActiveChoiceIndex(nature);
      selection.choiceIndex = choiceIndex;
      if (!groupedSelections[nature].has(choiceIndex)) {
        groupedSelections[nature].set(choiceIndex, []);
      }
      groupedSelections[nature].get(choiceIndex).push(selection);
    });

    CHOICE_SERIES.forEach((nature) => {
      groupedSelections[nature].forEach((list) => {
        list.forEach((selection, position) => {
          selection.choiceRank = position + 1;
          selection.isPrimary = position === 0;
          selection.choiceLabel = `${selection.choiceIndex}.${selection.choiceRank}`;
        });
      });
    });

    state.selections.forEach((selection) => {
      const button = selection.button;
      if (!button) {
        return;
      }
      const tag = button.querySelector('.planning-assignment-tag');
      const roleLabel = button.querySelector('.planning-assignment-role');
      const srLabel = button.querySelector('.sr-only');
      const choiceLabel = selection.choiceLabel ?? String(selection.order);
      if (tag) {
        tag.textContent = choiceLabel;
        tag.classList.remove('is-empty');
      }
      if (roleLabel) {
        const bullet = selection.isPrimary ? '●' : '○';
        const natureLabel = selection.nature === 'bonus' ? 'Bonus' : 'Mauvaise';
        roleLabel.textContent = `${bullet} ${natureLabel}`;
        roleLabel.classList.remove('is-empty');
      }
      if (srLabel) {
        const roleDescription = selection.isPrimary
          ? 'principal'
          : `alternative ${selection.choiceRank - 1}`;
        srLabel.textContent = selection.summary
          ? `${selection.summary} · Choix ${choiceLabel} (${roleDescription})`
          : `Choix ${choiceLabel} (${roleDescription})`;
      }
      button.dataset.choiceOrder = choiceLabel;
      button.dataset.choiceType = selection.nature;
      button.dataset.choiceRank = String(selection.choiceRank ?? 1);
      button.dataset.choiceIndex = String(selection.choiceIndex ?? CHOICE_INDEX_MIN);
      button.dataset.choiceRole = selection.isPrimary ? 'principal' : 'alternative';
      button.dataset.choiceLabel = choiceLabel;
      button.classList.add('is-selected');
    });

    updateChoiceIndexButtons(groupedSelections);
  };

  const renderSummaryList = () => {
    if (!summaryList) {
      return;
    }
    summaryList.innerHTML = '';
    state.selections.forEach((selection) => {
      const item = document.createElement('li');
      item.className = 'summary-item';
      item.draggable = true;
      item.dataset.slotKey = selection.slotKey;
      item.dataset.choiceIndex = String(selection.choiceIndex ?? CHOICE_INDEX_MIN);
      item.dataset.choiceRank = String(selection.choiceRank ?? 1);
      item.innerHTML = `
        <div class="summary-item-order">${selection.choiceLabel ?? selection.order}</div>
        <div class="summary-item-body">
          <span class="summary-item-date">${formatSummaryLabel(selection)}</span>
          <span class="summary-item-details">Col. ${selection.columnPosition} · ${
        selection.columnLabel || selection.slotTypeCode || 'Créneau'
      }</span>
          <span class="summary-item-tags">
            <span class="badge ${selection.nature === 'bonus' ? 'badge-success' : 'badge-warning'}">${
        (selection.isPrimary ? '●' : '○') + ' ' + (selection.nature === 'bonus' ? 'Bonus' : 'Mauvaise')
      }</span>
            <span class="badge">${selection.activityType}</span>
          </span>
        </div>
        <button type="button" class="summary-item-remove" aria-label="Retirer ce choix">&times;</button>
      `;
      summaryList.appendChild(item);
    });
    updateSummaryFeedback();
  };

  const refreshSelections = () => {
    updateSelectionVisuals();
    renderSummaryList();
  };

  const removeSelection = (slotKey) => {
    if (!state.selectionMap.has(slotKey)) {
      return;
    }
    const selection = state.selectionMap.get(slotKey);
    state.selectionMap.delete(slotKey);
    state.selections = state.selections.filter((item) => item.slotKey !== slotKey);
    if (selection?.button) {
      selection.button.classList.remove('is-selected');
    }
    refreshSelections();
  };

  const addSelection = (selection) => {
    if (state.selectionMap.has(selection.slotKey)) {
      removeSelection(selection.slotKey);
    }
    state.selectionMap.set(selection.slotKey, selection);
    state.selections.push(selection);
    refreshSelections();
  };

  const schedulePlanningCellSize = () => {
    if (state.planningCellSizingFrame) {
      cancelAnimationFrame(state.planningCellSizingFrame);
    }
    state.planningCellSizingFrame = requestAnimationFrame(() => {
      state.planningCellSizingFrame = null;
      if (!planningTables) {
        return;
      }
      planningTables.style.removeProperty('--planning-cell-width');
      planningTables.style.removeProperty('--planning-cell-height');
      const widthElements = new Set([
        ...planningTables.querySelectorAll('.planning-slot-toggle'),
        ...planningTables.querySelectorAll('.planning-slot-button'),
        ...planningTables.querySelectorAll('.planning-slot-index')
      ]);
      const heightElements = new Set([
        ...planningTables.querySelectorAll('.planning-slot-toggle'),
        ...planningTables.querySelectorAll('.planning-slot-button')
      ]);
      if (!widthElements.size && !heightElements.size) {
        return;
      }
      let maxWidth = 0;
      let maxHeight = 0;
      widthElements.forEach((element) => {
        const rect = element.getBoundingClientRect();
        maxWidth = Math.max(maxWidth, rect.width);
      });
      heightElements.forEach((element) => {
        const rect = element.getBoundingClientRect();
        maxHeight = Math.max(maxHeight, rect.height);
      });
      const clampedWidth = clamp(Math.ceil(maxWidth) + 8, 120, 160);
      const clampedHeight = clamp(Math.ceil(maxHeight) + 12, 80, 140);
      planningTables.style.setProperty('--planning-cell-width', `${clampedWidth}px`);
      planningTables.style.setProperty('--planning-cell-height', `${clampedHeight}px`);
    });
  };

  const scheduleStickyHeaderUpdate = () => {
    if (state.stickyHeaderFrame) {
      cancelAnimationFrame(state.stickyHeaderFrame);
    }
    state.stickyHeaderFrame = requestAnimationFrame(() => {
      state.stickyHeaderFrame = null;
      if (!planningTables) {
        return;
      }
      const sections = planningTables.querySelectorAll('.planning-month');
      sections.forEach((section) => {
        const header = section.querySelector('.planning-month-header');
        const table = section.querySelector('.planning-table');
        if (!header || !table) {
          return;
        }
        const offset = table.offsetTop - section.offsetTop;
        section.style.setProperty('--planning-month-header-offset', `${offset}px`);
      });
    });
  };

  const updateButtonAvailability = () => {
    state.slotButtons.forEach((button) => {
      const isSelected = button.classList.contains('is-selected');
      const canSelect = state.selectionMode === 'mauvaise'
        ? button.dataset.openMauvaise === 'true'
        : state.selectionMode === 'bonus'
          ? button.dataset.openBonus === 'true'
          : false;
      if (!state.selectionMode) {
        button.dataset.state = canSelect ? 'available' : 'closed';
        button.disabled = true;
        button.setAttribute('aria-disabled', 'true');
        return;
      }
      if (isSelected || canSelect) {
        button.dataset.state = 'available';
        button.disabled = false;
        button.removeAttribute('aria-disabled');
        button.removeAttribute('tabindex');
      } else {
        button.dataset.state = 'closed';
        button.disabled = true;
        button.setAttribute('aria-disabled', 'true');
        button.tabIndex = -1;
      }
    });
  };

  const setStep = (step) => {
    if (!Number.isInteger(step) || step < 1 || step > 3) {
      return;
    }
    state.currentStep = step;
    state.selectionMode = step === 1 ? 'mauvaise' : step === 2 ? 'bonus' : null;
    updateStepperButtons();
    updateStepPanes();
    updateChoiceIndexActiveState('mauvaise');
    updateChoiceIndexActiveState('bonus');
    if (step <= 2) {
      movePlanningSectionToStep(step);
      planningSection.classList.remove('hidden');
    } else {
      planningSection.classList.add('hidden');
      renderSummaryList();
    }
    updateButtonAvailability();
  };

  const getSelectionForButton = (button) => state.selectionMap.get(button.dataset.slotKey);

  const handleSlotToggle = (button) => {
    if (!state.selectionMode) {
      return;
    }
    if (button.dataset.state === 'closed') {
      return;
    }
    const existing = getSelectionForButton(button);
    if (existing) {
      removeSelection(existing.slotKey);
      return;
    }
    const selection = createSelectionFromButton(button, state.selectionMode);
    selection.nature = sanitizeChoiceNature(selection.nature);
    selection.choiceIndex = getActiveChoiceIndex(selection.nature);
    addSelection(selection);
  };
  const buildPlanningTable = (year, month) => {
    const monthSection = document.createElement('section');
    monthSection.className = 'planning-month';
    const heading = document.createElement('header');
    heading.className = 'planning-month-header';
    heading.innerHTML = `
      <h4>${MONTH_NAMES[month]} ${year}</h4>
      <p>${PLANNING_TOURS.find((tour) => tour.id === state.selectedTourId)?.label ?? ''}</p>
    `;
    monthSection.appendChild(heading);

    const table = document.createElement('table');
    table.className = 'planning-table';
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const dayHeader = document.createElement('th');
    dayHeader.scope = 'col';
    dayHeader.className = 'planning-day-col';
    dayHeader.textContent = 'Jour';
    headerRow.appendChild(dayHeader);

    const columnStyles = new Map();
    const columnColors = new Map();

    state.planningSlots.forEach((slot) => {
      const th = document.createElement('th');
      th.scope = 'col';
      const slotColor = normalizeColor(slot.color);
      const slotTextColor = getSlotTextColor(slotColor);
      const columnTint = getSlotColumnTint(slotColor);
      columnStyles.set(slot.position, columnTint);
      columnColors.set(slot.position, { color: slotColor, text: slotTextColor });
      th.style.setProperty('--slot-header-background', columnTint);
      th.style.setProperty('--slot-color', slotColor);
      th.style.setProperty('--slot-text-color', slotTextColor);
      const headerContent = document.createElement('div');
      headerContent.className = 'planning-slot-header';
      const slotNumber = String(slot.position).padStart(2, '0');
      const rawLabel = (slot.label ?? '').trim();
      const slotLabel = rawLabel || `Colonne ${slot.position}`;
      const slotType = slot.type_code ?? '';
      const slotTime = getSlotTimeRange(slot);
      const buttonTitleParts = [slotNumber, slotLabel];
      if (slotType) {
        buttonTitleParts.push(slotType);
      }
      if (slotTime) {
        buttonTitleParts.push(slotTime);
      }
      const buttonTitle = buttonTitleParts.join(' · ');
      const indexBadge = document.createElement('span');
      indexBadge.className = 'planning-slot-index';
      indexBadge.textContent = slotNumber;
      indexBadge.setAttribute('aria-hidden', 'true');
      headerContent.appendChild(indexBadge);
      const button = document.createElement('button');
      button.type = 'button';
      button.dataset.position = String(slot.position);
      button.className = 'planning-slot-button';
      button.style.setProperty('--slot-color', slotColor);
      button.style.setProperty('--slot-text-color', slotTextColor);
      button.title = buttonTitle;
      button.setAttribute('aria-label', buttonTitle);
      const headingSpan = document.createElement('span');
      headingSpan.className = 'planning-slot-heading';
      if (slotType) {
        const typeLabel = document.createElement('span');
        typeLabel.className = 'planning-slot-type';
        typeLabel.textContent = slotType;
        headingSpan.appendChild(typeLabel);
      }
      const titleLabel = document.createElement('span');
      titleLabel.className = 'planning-slot-title';
      titleLabel.textContent = slotLabel;
      headingSpan.appendChild(titleLabel);
      const metaParts = [];
      if (slotTime) {
        metaParts.push(slotTime);
      }
      if (metaParts.length) {
        const metaLabel = document.createElement('span');
        metaLabel.className = 'planning-slot-meta';
        metaLabel.textContent = metaParts.join(' · ');
        headingSpan.appendChild(metaLabel);
      }
      button.appendChild(headingSpan);
      const srLabel = document.createElement('span');
      srLabel.className = 'sr-only';
      srLabel.textContent = buttonTitle;
      button.appendChild(srLabel);
      headerContent.appendChild(button);
      th.appendChild(headerContent);
      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    const holidays = getHolidaysForYear(year);
    const days = getMonthDays(year, month);
    days.forEach((day) => {
      const isoKey = formatDateKey(day);
      const holidayName = holidays?.get(isoKey) ?? null;
      const isHoliday = Boolean(holidayName);
      const row = document.createElement('tr');
      if (isHoliday) {
        row.classList.add('planning-holiday');
      }
      if (day.getDay() === 0) {
        row.classList.add('planning-sunday');
      }

      const dayCell = document.createElement('th');
      dayCell.scope = 'row';
      dayCell.className = 'planning-day-header';
      const dayName = WEEKDAY_LABELS[day.getDay()];
      const dayNumber = String(day.getDate()).padStart(2, '0');
      dayCell.innerHTML = `
        <span class="day-name">${dayName}</span>
        <span class="day-number">${dayNumber} ${MONTH_NAMES[month]}</span>
      `;
      if (holidayName) {
        const badge = document.createElement('span');
        badge.className = 'holiday-badge';
        badge.textContent = holidayName;
        dayCell.appendChild(badge);
      }
      row.appendChild(dayCell);

      state.planningSlots.forEach((slot) => {
        const cell = document.createElement('td');
        cell.className = 'planning-summary-cell';
        const columnTint = columnStyles.get(slot.position);
        const columnColor = columnColors.get(slot.position);
        if (columnTint) {
          cell.style.setProperty('--slot-column-background', columnTint);
        }
        const segment = getDaySegment(day, isHoliday);
        const quality = getQualityForSegment(slot, segment);
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'planning-slot-toggle';
        const slotKey = `${isoKey}:${slot.position}`;
        button.dataset.slotKey = slotKey;
        button.dataset.day = isoKey;
        button.dataset.position = String(slot.position);
        button.dataset.summary = buildSlotTitle(
          slot,
          `${dayNumber} ${MONTH_NAMES[month]} ${year}`,
          isHoliday,
          holidayName,
          isSlotOpen(slot, segment, null),
          quality
        );
        const openings = getSlotOpeningsForSegment(slot, segment);
        button.dataset.openMauvaise = String(openings.mauvaise);
        button.dataset.openBonus = String(openings.bonus);
        button.dataset.month = String(day.getMonth() + 1);
        button.dataset.year = String(day.getFullYear());
        button.dataset.dayLabel = `${dayName} ${dayNumber} ${MONTH_NAMES[month]} ${year}`;
        button.dataset.typeCode = slot.type_code ?? '';
        const typeCategory = slot.type_category ?? inferTypeCategory(slot.type_code);
        button.dataset.activityCategory = typeCategory;
        button.dataset.columnLabel = (slot.label ?? '').trim() || `Colonne ${slot.position}`;
        if (columnColor) {
          button.style.setProperty('--slot-color', columnColor.color);
          button.style.setProperty('--slot-text-color', columnColor.text);
        }
        const tag = document.createElement('span');
        tag.className = 'planning-assignment-tag is-empty';
        button.appendChild(tag);
        const roleLabel = document.createElement('span');
        roleLabel.className = 'planning-assignment-role is-empty';
        button.appendChild(roleLabel);
        const sr = document.createElement('span');
        sr.className = 'sr-only';
        sr.textContent = button.dataset.summary ?? '';
        button.appendChild(sr);
        cell.appendChild(button);
        row.appendChild(cell);
        state.slotButtons.set(slotKey, button);
      });

      tbody.appendChild(row);
    });

    table.appendChild(tbody);
    monthSection.appendChild(table);
    return monthSection;
  };

  const renderPlanningTables = () => {
    if (!planningTables) {
      return;
    }
    planningTables.innerHTML = '';
    state.slotButtons.clear();
    if (!state.planningSlots.length) {
      planningTables.innerHTML = '<p class="empty-planning">Chargement du planning…</p>';
      return;
    }
    const months = [state.planningMonthOne, state.planningMonthTwo];
    months.forEach((month) => {
      if (!Number.isInteger(month)) {
        return;
      }
      const section = buildPlanningTable(state.planningYear, month);
      planningTables.appendChild(section);
    });
    state.selections.forEach((selection) => {
      const nextButton = state.slotButtons.get(selection.slotKey);
      if (nextButton) {
        selection.button = nextButton;
      }
    });
    schedulePlanningCellSize();
    scheduleStickyHeaderUpdate();
    updateButtonAvailability();
    refreshSelections();
  };

  const loadPlanningColumns = async () => {
    await onSupabaseReady();
    const supabase = getSupabaseClient();
    state.supabase = supabase;
    if (!supabase) {
      setPlanningFeedback('Connexion à Supabase requise.');
      return;
    }
    setPlanningFeedback('Chargement des colonnes…');
    const tableName = getPlanningTableName(state.selectedTourId);
    const { data, error } = await supabase
      .from(tableName)
      .select(
        `id,
          position,
          label,
          type_code,
          type_category,
          start_time,
          end_time,
          color,
          quality_weekdays,
          quality_saturday,
          quality_sunday,
          open_mauvaise_weekdays,
          open_mauvaise_saturday,
          open_mauvaise_sunday,
          open_bonus_weekdays,
          open_bonus_saturday,
          open_bonus_sunday`
      )
      .order('position');
    if (error) {
      console.error(error);
      setPlanningFeedback('Impossible de charger les colonnes.');
      planningTables.innerHTML = '<p class="empty-planning">Erreur de chargement.</p>';
      return;
    }
    let slots = data ?? [];
    slots = await ensurePlanningColumns(supabase, slots, tableName);
    state.planningSlots = sortSlotsByPosition(slots).map((slot) => ({
      ...slot,
      color: normalizeColor(slot.color),
      type_category: slot.type_category ?? inferTypeCategory(slot.type_code)
    }));
    setPlanningFeedback(`${state.planningSlots.length} colonne${state.planningSlots.length > 1 ? 's' : ''}.`);
    renderPlanningTables();
  };

  const fetchAdministrativeSettings = async () => {
    await onSupabaseReady();
    const supabase = getSupabaseClient();
    state.supabase = supabase;
    if (!supabase) {
      return null;
    }
    const { data, error } = await supabase
      .from(ADMIN_SETTINGS_TABLE)
      .select('id, active_tour, planning_year, planning_month_one, planning_month_two')
      .order('id', { ascending: true })
      .limit(1);
    if (error) {
      console.error(error);
      return null;
    }
    if (!data || data.length === 0) {
      return null;
    }
    return data[0];
  };

  const applyAdministrativeSettings = (settings) => {
    const now = new Date();
    const year = Number.isInteger(settings?.planning_year) ? settings.planning_year : now.getFullYear();
    const monthOne = Number.isInteger(settings?.planning_month_one)
      ? settings.planning_month_one
      : now.getMonth();
    const monthTwo = Number.isInteger(settings?.planning_month_two)
      ? settings.planning_month_two
      : (monthOne + 1) % 12;
    const sanitizedActiveTour = sanitizeActiveTour(settings?.active_tour) ?? state.selectedTourId;
    state.selectedTourId = sanitizedActiveTour;
    state.planningYear = year;
    state.planningMonthOne = monthOne;
    state.planningMonthTwo = monthTwo;
    state.planningReference = getPlanningReference({
      tourId: sanitizedActiveTour,
      year,
      monthOne,
      monthTwo
    });
  };

  const loadAdministrativeSettings = async () => {
    const settings = await fetchAdministrativeSettings();
    applyAdministrativeSettings(settings);
  };

  const fetchUserProfile = async () => {
    await onSupabaseReady();
    const supabase = getSupabaseClient();
    state.supabase = supabase;
    const currentUser = getCurrentUser();
    if (!supabase || !currentUser) {
      return null;
    }
    const { data, error } = await supabase
      .from('users')
      .select('id, username, trigram')
      .eq('id', currentUser.id)
      .maybeSingle();
    if (error) {
      console.error(error);
      return null;
    }
    if (!data) {
      return {
        id: currentUser.id,
        username: currentUser.username,
        trigram: (currentUser.username ?? '').slice(0, 3).toUpperCase()
      };
    }
    return {
      ...data,
      trigram: (data.trigram ?? data.username ?? '').slice(0, 3).toUpperCase()
    };
  };

  const saveSelections = async () => {
    if (state.isSaving) {
      return;
    }
    if (!state.selections.length) {
      if (saveFeedback) {
        saveFeedback.textContent = 'Aucune sélection à enregistrer.';
      }
      return;
    }
    await onSupabaseReady();
    const supabase = getSupabaseClient();
    if (!supabase) {
      if (saveFeedback) {
        saveFeedback.textContent = 'Connexion à Supabase requise.';
      }
      return;
    }
    if (!state.userProfile) {
      state.userProfile = await fetchUserProfile();
    }
    const trigram = state.userProfile?.trigram ?? '???';
    const userId = state.userProfile?.id ?? null;
    const nowIso = new Date().toISOString();
    const payload = state.selections.map((selection, index) => ({
      user_id: userId,
      trigram,
      user_type: normalizedUserRole,
      created_at: nowIso,
      day: selection.dayIso,
      month: selection.dayDate.getMonth() + 1,
      year: selection.dayDate.getFullYear(),
      column_number: selection.columnPosition,
      guard_nature: selection.nature,
      activity_type: selection.activityType,
      choice_order: index + 1,
      choice_index: sanitizeChoiceIndex(selection.choiceIndex),
      choice_rank: sanitizeChoiceRank(selection.choiceRank),
      tour_number: state.selectedTourId,
      planning_reference: state.planningReference,
      is_active: true,
      slot_type_code: selection.slotTypeCode,
      column_label: selection.columnLabel,
      planning_day_label: selection.dayLabel
    }));

    state.isSaving = true;
    saveButton?.setAttribute('disabled', 'true');
    if (saveFeedback) {
      saveFeedback.textContent = 'Enregistrement en cours…';
    }

    const { error } = await supabase.from('planning_choices').insert(payload);
    state.isSaving = false;
    saveButton?.removeAttribute('disabled');

    if (error) {
      console.error(error);
      if (saveFeedback) {
        saveFeedback.textContent = "Impossible d'enregistrer vos choix.";
      }
      return;
    }

    if (saveFeedback) {
      saveFeedback.textContent = 'Choix enregistrés avec succès !';
    }
  };
  if (stepper) {
    stepper.addEventListener('click', (event) => {
      const button = event.target.closest('.stepper-step');
      if (!button) {
        return;
      }
      const step = Number.parseInt(button.dataset.step, 10);
      setStep(step);
    });
  }

  document.addEventListener('click', (event) => {
    const control = event.target.closest('.step-nav');
    if (!control) {
      return;
    }
    const action = control.dataset.action;
    if (action === 'next') {
      setStep(Math.min(state.currentStep + 1, 3));
    } else if (action === 'previous') {
      setStep(Math.max(state.currentStep - 1, 1));
    }
  });

  if (planningTables) {
    planningTables.addEventListener('click', (event) => {
      const button = event.target.closest('.planning-slot-toggle');
      if (!button || !planningTables.contains(button)) {
        return;
      }
      handleSlotToggle(button);
    });
  }

  if (summaryList) {
    summaryList.addEventListener('click', (event) => {
      const removeBtn = event.target.closest('.summary-item-remove');
      if (!removeBtn) {
        return;
      }
      const item = removeBtn.closest('.summary-item');
      if (!item) {
        return;
      }
      const slotKey = item.dataset.slotKey;
      removeSelection(slotKey);
    });

    summaryList.addEventListener('dragstart', (event) => {
      const item = event.target.closest('.summary-item');
      if (!item) {
        return;
      }
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData(DRAG_DATA_TYPE, item.dataset.slotKey);
      item.classList.add('is-dragging');
    });

    const getDragAfterElement = (container, y) => {
      const draggableElements = [...container.querySelectorAll('.summary-item:not(.is-dragging)')];
      return draggableElements.reduce(
        (closest, child) => {
          const box = child.getBoundingClientRect();
          const offset = y - box.top - box.height / 2;
          if (offset < 0 && offset > closest.offset) {
            return { offset, element: child };
          }
          return closest;
        },
        { offset: Number.NEGATIVE_INFINITY, element: null }
      ).element;
    };

    summaryList.addEventListener('dragover', (event) => {
      event.preventDefault();
      const afterElement = getDragAfterElement(summaryList, event.clientY);
      const dragging = summaryList.querySelector('.summary-item.is-dragging');
      if (!dragging) {
        return;
      }
      if (afterElement == null) {
        summaryList.appendChild(dragging);
      } else {
        summaryList.insertBefore(dragging, afterElement);
      }
    });

    summaryList.addEventListener('drop', (event) => {
      event.preventDefault();
    });

    summaryList.addEventListener('dragend', () => {
      const dragging = summaryList.querySelector('.summary-item.is-dragging');
      if (dragging) {
        dragging.classList.remove('is-dragging');
      }
      const orderedKeys = [...summaryList.querySelectorAll('.summary-item')].map((item) => item.dataset.slotKey);
      const newSelections = [];
      orderedKeys.forEach((key) => {
        const selection = state.selectionMap.get(key);
        if (selection) {
          newSelections.push(selection);
        }
      });
      if (newSelections.length === state.selections.length) {
        state.selections = newSelections;
        refreshSelections();
      }
    });
  }

  if (saveButton) {
    saveButton.addEventListener('click', () => {
      saveSelections();
    });
  }

  if (planningBoard) {
    planningBoard.addEventListener('scroll', () => {
      scheduleStickyHeaderUpdate();
    });
  }

  window.addEventListener('resize', () => {
    schedulePlanningCellSize();
    scheduleStickyHeaderUpdate();
  });

  const initialize = async () => {
    await loadAdministrativeSettings();
    await loadPlanningColumns();
    state.userProfile = await fetchUserProfile();
    setStep(1);
  };

  initializeChoiceSelectors();
  initialize();
}
