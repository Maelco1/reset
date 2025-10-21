import { createClient } from 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js/+esm';

const STORAGE_KEYS = {
  url: 'guardes:supabaseUrl',
  key: 'guardes:supabaseKey',
  user: 'guardes:currentUser'
};

const ROLE_ALIASES = new Map([
  ['administrateur', 'administrateur'],
  ['administrateurs', 'administrateur'],
  ['administratrice', 'administrateur'],
  ['administratrices', 'administrateur'],
  ['admin', 'administrateur'],
  ['admins', 'administrateur'],
  ['administration', 'administrateur'],
  ['administrations', 'administrateur'],
  ['medecin', 'medecin'],
  ['medecins', 'medecin'],
  ['medecine', 'medecin'],
  ['medecines', 'medecin'],
  ['docteur', 'medecin'],
  ['docteurs', 'medecin'],
  ['remplacant', 'remplacant'],
  ['remplacants', 'remplacant'],
  ['remplacante', 'remplacant'],
  ['remplacantes', 'remplacant'],
  ['remplacement', 'remplacant'],
  ['remplacements', 'remplacant']
]);

let client = null;
let resolveReady;
const readyPromise = new Promise((resolve) => {
  resolveReady = resolve;
});

const DIACRITICS_REGEX = /[\u0300-\u036f]/g;

export function normalizeRole(value) {
  if (value == null) {
    return '';
  }

  const normalizedInput = value
    .toString()
    .normalize('NFD')
    .replace(DIACRITICS_REGEX, '')
    .toLowerCase()
    .trim();

  if (!normalizedInput) {
    return '';
  }

  const compact = normalizedInput.replace(/[^a-z]/g, '');
  if (!compact) {
    return '';
  }

  const candidates = [compact];
  if (compact.endsWith('s')) {
    candidates.push(compact.slice(0, -1));
  }

  for (const candidate of candidates) {
    if (ROLE_ALIASES.has(candidate)) {
      return ROLE_ALIASES.get(candidate);
    }
  }

  return '';
}

function buildClient(url, key) {
  if (!url || !key) {
    return null;
  }
  client = createClient(url, key);
  resolveReady?.(client);
  return client;
}

export function getStoredConnection() {
  return {
    url: localStorage.getItem(STORAGE_KEYS.url) ?? '',
    key: localStorage.getItem(STORAGE_KEYS.key) ?? ''
  };
}

export function setStoredConnection(url, key) {
  localStorage.setItem(STORAGE_KEYS.url, url);
  localStorage.setItem(STORAGE_KEYS.key, key);
  return buildClient(url, key);
}

export function clearStoredConnection() {
  localStorage.removeItem(STORAGE_KEYS.url);
  localStorage.removeItem(STORAGE_KEYS.key);
  client = null;
}

export function getSupabaseClient() {
  if (client) {
    return client;
  }
  const { url, key } = getStoredConnection();
  if (!url || !key) {
    return null;
  }
  return buildClient(url, key);
}

export function onSupabaseReady() {
  return readyPromise;
}

export function getCurrentUser() {
  const raw = localStorage.getItem(STORAGE_KEYS.user);
  if (!raw) {
    return null;
  }
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') {
      return null;
    }

    const normalizedRole = normalizeRole(parsed.role);
    if (normalizedRole) {
      if (parsed.role !== normalizedRole) {
        const updated = { ...parsed, role: normalizedRole };
        localStorage.setItem(STORAGE_KEYS.user, JSON.stringify(updated));
        return updated;
      }
      return { ...parsed, role: normalizedRole };
    }

    return parsed;
  } catch (error) {
    console.error('Failed to parse stored user', error);
    return null;
  }
}

export function setCurrentUser(user) {
  if (user) {
    const normalizedRole = normalizeRole(user.role);
    const payload = normalizedRole ? { ...user, role: normalizedRole } : { ...user };
    localStorage.setItem(STORAGE_KEYS.user, JSON.stringify(payload));
    return payload;
  }

  localStorage.removeItem(STORAGE_KEYS.user);
  return null;
}

export function requireRole(role) {
  const expectedRole = normalizeRole(role);
  const user = getCurrentUser();
  const actualRole = normalizeRole(user?.role);
  if (!user || !expectedRole || actualRole !== expectedRole) {
    window.location.assign('index.html');
  }
}

export function initializeConnectionModal() {
  const modal = document.querySelector('#connection-modal');
  const form = document.querySelector('#connection-form');
  const disconnectBtn = document.querySelector('#disconnect');

  if (!modal || !form) {
    throw new Error('Connection modal markup missing from page.');
  }

  const { url, key } = getStoredConnection();
  if (url) {
    form.elements['supabaseUrl'].value = url;
  }
  if (key) {
    form.elements['supabaseKey'].value = key;
  }

  const connect = (event) => {
    event?.preventDefault();
    const supabaseUrl = form.elements['supabaseUrl'].value.trim();
    const supabaseKey = form.elements['supabaseKey'].value.trim();

    if (!supabaseUrl || !supabaseKey) {
      alert('Merci de renseigner l\'URL et la clé API Supabase.');
      return;
    }

    setStoredConnection(supabaseUrl, supabaseKey);
    modal.classList.add('hidden');
  };

  form.addEventListener('submit', connect);

  if (disconnectBtn) {
    disconnectBtn.addEventListener('click', () => {
      clearStoredConnection();
      modal.classList.remove('hidden');
    });
  }

  if (url && key) {
    buildClient(url, key);
    modal.classList.add('hidden');
  } else {
    modal.classList.remove('hidden');
  }
}
