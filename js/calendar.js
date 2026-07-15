const COLORS = ['#4f46e5','#0ea5e9','#10b981','#f59e0b','#ef4444','#8b5cf6','#ec4899','#14b8a6'];

// ── State ────────────────────────────────────────────────────────────────────
let events = [];
let calendars = [];
let calendarGroups = [];
let openCalendarIds = [];       // visible calendars (max 5)
let primaryCalendarId = null;   // receives new events
let selectedCalId = null;       // calendar whose events are highlighted
let pendingCommand = null;
let editingCalId = null;
let editingCalColor = null;     // color being edited in calendar modal
let addCalGroupId = null;
let editingGroupId = null;
let view = 'week', weekDays = 7;
let cur = new Date(), today = new Date();
let editingId = null, dragEv = null, dragOffset = 0;
let resizeEv = null, resizeEdge = null, resizeEl = null, resizeCol = null;
let resizeOrigStart = '', resizeOrigEnd = '';
let selColor = COLORS[0];
let isAllDay = false;
let recurActive    = false;
let editingRecurDate = null;   // date string of virtual occurrence being edited/deleted
let editingRecurEv   = null  // full event object of the virtual occurrence being edited
const DOW_LABELS   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

// Theme persistence
const THEME_STORAGE_KEY = 'accessCalendarTheme';
let darkMode = false;

function readSavedThemePreference() {
  try {
    const v = localStorage.getItem(THEME_STORAGE_KEY);
    if (v === 'dark') return true;
    if (v === 'light') return false;
  } catch (e) {}
  return false;
}

function saveThemePreference(isDark) {
  try {
    localStorage.setItem(THEME_STORAGE_KEY, isDark ? 'dark' : 'light');
  } catch (e) {}
}

function getSavedThemePreference() {
  try {
    const v = localStorage.getItem(THEME_STORAGE_KEY);
    if (v === 'dark' || v === 'light') return v;
  } catch (e) {}
  return darkMode ? 'dark' : 'light';
}

window.getSavedThemePreference = getSavedThemePreference;

// ── Access detection ──────────────────────────────────────────────────────────
const inAccess = !!(window.chrome && window.chrome.webview);
const standalone = !inAccess;

// ── VBA Bridge ───────────────────────────────────────────────────────────────
// JS→VBA command channel. The Access EdgeBrowserControl (WebView2) exposes NO
// WebMessageReceived event, and RetrieveJavascriptValue (the old polling
// channel) is broken by WebView2 Runtime 149+. So each command is pushed to VBA
// by navigating to a sentinel URL that the form's WebBrowser0_BeforeNavigate
// handler intercepts and cancels. See cls/vba_form_code_behind.txt.
const VBA_CMD_URL = 'https://msaccess/__cmd__?data=';
let _cmdQueue = [];
let _cmdDraining = false;

function getPendingCommand() { return pendingCommand != null ? JSON.stringify(pendingCommand) : null; }
function clearPendingCommand() { pendingCommand = null; setStatus('idle'); }

function queueCommand(cmd) {
  pendingCommand = cmd;                       // retained for standalone/legacy readers
  setStatus('pending', 'Saving\u2026');
  if (!inAccess) return;                      // standalone browser preview: no VBA to notify
  _cmdQueue.push(cmd);
  _drainCommandQueue();
}

// Serialize navigations so rapid back-to-back commands are not lost: fire one
// sentinel navigation, let BeforeNavigate cancel it, then release for the next.
function _drainCommandQueue() {
  if (_cmdDraining || !_cmdQueue.length) return;
  _cmdDraining = true;
  const cmd = _cmdQueue.shift();
  try {
    window.location.href = VBA_CMD_URL +
      encodeURIComponent(JSON.stringify(cmd)) + '&t=' + Date.now();
  } catch (e) {}
  setTimeout(function () { _cmdDraining = false; _drainCommandQueue(); }, 60);
}

// ── Primary data entry point (called by VBA) ──────────────────────────────────
window.loadData = function(json) {
  try {
    const d = JSON.parse(json);
    calendars      = d.calendars      || [];
    calendarGroups = d.calendarGroups  || [];
    events         = d.appointments   || [];
    
    const storedTheme = getSavedThemePreference();
    if (storedTheme === 'dark') {
      darkMode = true;
    } else if (storedTheme === 'light') {
      darkMode = false;
    } else if (typeof d.darkMode === 'boolean') {
      darkMode = d.darkMode;
    } else {
      darkMode = false;
    }
    applyTheme(darkMode, false);
    // Restore open calendars and primary
    if (Array.isArray(d.openCalendarIds) && d.openCalendarIds.length) {
      openCalendarIds = d.openCalendarIds.map(String).filter(id => calendars.some(c => String(c.id) === id));
    }
    const incomingPrimary = d.primaryCalendarId ? String(d.primaryCalendarId) : (d.activeCalendarId ? String(d.activeCalendarId) : null);
    if (incomingPrimary && calendars.some(c => String(c.id) === incomingPrimary)) {
      primaryCalendarId = incomingPrimary;
    } else if (!primaryCalendarId && calendars.length) {
      primaryCalendarId = String(calendars[0].id);
    } else if (primaryCalendarId && !calendars.some(c => String(c.id) === String(primaryCalendarId))) {
      primaryCalendarId = calendars.length ? String(calendars[0].id) : null;
    }
    // Ensure primary is in the open set; default to primary if open set is empty
    if (!openCalendarIds.length && primaryCalendarId) {
      openCalendarIds = [primaryCalendarId];
    }
    if (primaryCalendarId && !openCalendarIds.includes(primaryCalendarId)) {
      openCalendarIds.unshift(primaryCalendarId);
    }
    renderSidebar();
    render();
    setStatus('idle');
  } catch(e) {
    setStatus('error', 'Bad data from Access');
    console.error(e);
  }
};

// ── Backward-compat shim (old VBA still calling loadEventsFromAccess) ─────────
function loadEventsFromAccess(eventsArg) {
  try {
    events = Array.isArray(eventsArg) ? eventsArg : JSON.parse(eventsArg);
    setStatus('idle');
    render();
  } catch(e) { setStatus('error', 'Bad data from Access'); console.error(e); }
}

// ── Filtered events for active calendar ──────────────────────────────────────
function calEvents() {
  if (!openCalendarIds.length) return events;
  return events.filter(e => openCalendarIds.includes(String(e.calendarId)));
}

// ── Multi-calendar helpers ───────────────────────────────────────────────────
function isCalendarOpen(id) {
  return openCalendarIds.includes(String(id));
}

function getCalendarColor(calId) {
  const cal = calendars.find(c => String(c.id) === String(calId));
  return (cal && cal.color) || COLORS[0];
}

function eventDisplayColor(ev) {
  const cal = calendars.find(c => String(c.id) === String(ev.calendarId));
  return (cal && cal.color) ? cal.color : (ev.color || COLORS[0]);
}
function eventAccentColor(ev) {
  // Use the appointment's own color for the accent strip; darken it for contrast
  const c = ev.color || COLORS[0];
  return darkenColor(c, 40);
}
function isEvSelected(ev) {
  return selectedCalId && String(ev.calendarId) === String(selectedCalId);
}
function darkenColor(hex, amount) {
  let c = hex.replace('#','');
  if (c.length===3) c=c[0]+c[0]+c[1]+c[1]+c[2]+c[2];
  const r=Math.max(0,parseInt(c.substring(0,2),16)-amount);
  const g=Math.max(0,parseInt(c.substring(2,4),16)-amount);
  const b=Math.max(0,parseInt(c.substring(4,6),16)-amount);
  return '#'+[r,g,b].map(v=>v.toString(16).padStart(2,'0')).join('');
}

function toggleCalendar(id) {
  id = String(id);
  const idx = openCalendarIds.indexOf(id);
  if (idx >= 0) {
    if (openCalendarIds.length <= 1) return; // don't close the last one
    openCalendarIds.splice(idx, 1);
    if (primaryCalendarId === id) primaryCalendarId = openCalendarIds[0];
    if (selectedCalId === id) selectedCalId = null;
  } else {
    if (openCalendarIds.length >= 5) return; // max 5 overlay
    openCalendarIds.push(id);
  }
  renderSidebar();
  render();
  notifyOpenCalendars();
}

function setPrimaryCalendar(id) {
  id = String(id);
  if (!openCalendarIds.includes(id)) return;
  primaryCalendarId = id;
  renderSidebar();
  renderCalTabs();
}

function notifyOpenCalendars() {
  queueCommand({
    action: 'setOpenCalendars',
    calendarIds: openCalendarIds.join(','),
    primaryId: primaryCalendarId
  });
}

// ── Status bar ───────────────────────────────────────────────────────────────
function setStatus(type, msg) {
  const el = document.getElementById('status-bar');
  const labels = { idle: '\u25CF Connected', pending: 'Saving\u2026', error: 'Error' };
  el.className = 'status-bar ' + type;
  el.textContent = msg || labels[type] || type;
}
setStatus('idle');

// ── Sidebar ──────────────────────────────────────────────────────────────────

// ── Sidebar collapse ──────────────────────────────────────────────────────
document.getElementById('sidebar-toggle-btn').addEventListener('click', function() {
  const sidebar = document.getElementById('sidebar');
  const collapsed = sidebar.classList.toggle('collapsed');
  this.textContent = collapsed ? '\u203A' : '\u2039';   // › / ‹
  this.title       = collapsed ? 'Expand sidebar' : 'Collapse sidebar';
});
function calItemHtml(c) {
  const isOpen = openCalendarIds.includes(String(c.id));
  const isPrimary = String(c.id) === String(primaryCalendarId);
  const dotColor = c.color || COLORS[0];
  return `<div class="cal-item${isOpen?' active':''}${isPrimary?' primary':''}" data-id="${c.id}" draggable="true">
    <input type="checkbox" class="cal-chk" id="calr-${c.id}" value="${c.id}"${isOpen?' checked':''}>
    <span class="cal-color-dot" style="background:${dotColor}"></span>
    <label class="cal-item-name" for="calr-${c.id}" title="${esc(c.name)}">${esc(c.name)}</label>
    <button class="cal-edit-btn" data-id="${c.id}" title="Edit calendar">&#9998;</button>
    <button class="cal-del-btn" data-id="${c.id}" title="Delete calendar">\u2715</button>
  </div>`;
}

function renderCalTabs() {
  const bar = document.getElementById('cal-tabs');
  if (openCalendarIds.length <= 1) { bar.innerHTML = ''; return; }
  bar.innerHTML = openCalendarIds.map(cid => {
    const cal = calendars.find(c => String(c.id) === String(cid));
    if (!cal) return '';
    const isPrimary = String(cid) === String(primaryCalendarId);
    const isSelected = String(cid) === String(selectedCalId);
    const color = cal.color || '#4f46e5';
    return `<span class="cal-tab${isPrimary ? ' primary' : ''}${isSelected && !isPrimary ? ' active' : ''}" data-id="${cal.id}">`
      + `<span class="cal-tab-dot" style="background:${color}"></span>`
      + `<span class="cal-tab-name">${esc(cal.name)}</span>`
      + `<span class="cal-tab-close" title="Hide calendar">&times;</span>`
      + `</span>`;
  }).join('');
  bar.querySelectorAll('.cal-tab').forEach(tab => {
    tab.querySelector('.cal-tab-name').addEventListener('click', () => {
      selectedCalId = String(tab.dataset.id) === String(selectedCalId) ? null : String(tab.dataset.id);
      setPrimaryCalendar(tab.dataset.id);
      render();
    });
    tab.querySelector('.cal-tab-close').addEventListener('click', e => { e.stopPropagation(); toggleCalendar(tab.dataset.id); });
  });
}

function renderSidebar() {
  const list = document.getElementById('cal-list');
  if (!calendars.length && !calendarGroups.length) {
    list.innerHTML = '<div class="cal-empty">No calendars yet.<br>Click + to add a group.</div>';
    return;
  }
  let html = '';
  // Render each group
  calendarGroups.forEach(g => {
    const gCals = calendars.filter(c => String(c.groupId) === String(g.id));
    html += `<div class="cal-group" data-group-id="${g.id}">
      <div class="cal-group-header" data-group-id="${g.id}">
        <span class="cal-group-arrow">\u25BE</span>
        <span class="cal-group-name" title="${esc(g.name)}">${esc(g.name)}</span>
        <button class="cal-group-add-btn" data-group-id="${g.id}" title="Add calendar to group">+</button>
        <div class="cal-group-actions">
          <button class="grp-edit-btn" data-group-id="${g.id}" title="Rename group">&#9998;</button>
          <button class="grp-del-btn" data-group-id="${g.id}" title="Delete group">\u2715</button>
        </div>
      </div>
      <div class="cal-group-body">
        ${gCals.length ? gCals.map(c => calItemHtml(c)).join('') : '<div class="cal-empty" style="padding:8px 12px;font-size:0.72rem">Drop calendars here</div>'}
      </div>
    </div>`;
  });
  // Ungrouped calendars
  const ungrouped = calendars.filter(c => !c.groupId || !calendarGroups.some(g => String(g.id) === String(c.groupId)));
  if (ungrouped.length || !calendarGroups.length) {
    html += `<div class="cal-group" data-group-id="">
      <div class="cal-group-header" data-group-id="">
        <span class="cal-group-arrow">\u25BE</span>
        <span class="cal-group-name">Ungrouped</span>
      </div>
      <div class="cal-group-body">
        ${ungrouped.length ? ungrouped.map(c => calItemHtml(c)).join('') : '<div class="cal-empty" style="padding:8px 12px;font-size:0.72rem">No ungrouped calendars</div>'}
      </div>
    </div>`;
  }
  list.innerHTML = html;
  // Wire calendar item events
  list.querySelectorAll('.cal-chk').forEach(chk => {
    chk.addEventListener('change', () => toggleCalendar(chk.value));
  });
  list.querySelectorAll('.cal-del-btn').forEach(btn => {
    btn.addEventListener('click', e => { e.stopPropagation(); deleteCalendar(btn.dataset.id); });
  });
  list.querySelectorAll('.cal-edit-btn').forEach(btn => {
    btn.addEventListener('click', e => { e.stopPropagation(); openCalEditModal(btn.dataset.id); });
  });
  list.querySelectorAll('.cal-item').forEach(item => {
    item.addEventListener('click', e => {
      if (e.target.closest('.cal-del-btn') || e.target.closest('.cal-edit-btn') || e.target.tagName === 'LABEL' || e.target.tagName === 'INPUT') return;
      // Click on the row sets this as primary (or opens it)
      if (openCalendarIds.includes(String(item.dataset.id))) {
        setPrimaryCalendar(item.dataset.id);
      } else {
        toggleCalendar(item.dataset.id);
      }
    });
    // Drag start for calendar items (between groups)
    item.addEventListener('dragstart', e => {
      e.dataTransfer.setData('text/cal-id', item.dataset.id);
      e.dataTransfer.effectAllowed = 'move';
      requestAnimationFrame(() => item.classList.add('dragging'));
    });
    item.addEventListener('dragend', () => item.classList.remove('dragging'));
  });
  // Wire group header events
  list.querySelectorAll('.cal-group-header').forEach(hdr => {
    hdr.addEventListener('click', e => {
      if (e.target.closest('.cal-group-actions') || e.target.closest('.cal-group-add-btn')) return;
      const body = hdr.nextElementSibling;
      const collapsed = hdr.classList.toggle('collapsed');
      body.classList.toggle('collapsed', collapsed);
    });
  });
  list.querySelectorAll('.grp-edit-btn').forEach(btn => {
    btn.addEventListener('click', e => { e.stopPropagation(); openGroupModal(btn.dataset.groupId); });
  });
  list.querySelectorAll('.grp-del-btn').forEach(btn => {
    btn.addEventListener('click', e => { e.stopPropagation(); deleteGroup(btn.dataset.groupId); });
  });
  list.querySelectorAll('.cal-group-add-btn').forEach(btn => {
    btn.addEventListener('click', e => { e.stopPropagation(); openCalModal(btn.dataset.groupId); });
  });
  // Wire group drag-drop targets
  list.querySelectorAll('.cal-group').forEach(grp => {
    grp.addEventListener('dragover', e => {
      if (!e.dataTransfer.types.includes('text/cal-id')) return;
      e.preventDefault();
      grp.classList.add('drag-over');
    });
    grp.addEventListener('dragleave', e => {
      if (!grp.contains(e.relatedTarget)) grp.classList.remove('drag-over');
    });
    grp.addEventListener('drop', e => {
      e.preventDefault();
      grp.classList.remove('drag-over');
      const calId = e.dataTransfer.getData('text/cal-id');
      if (!calId) return;
      const groupId = grp.dataset.groupId || '';
      moveCalendarToGroup(calId, groupId);
    });
  });
}

function switchCalendar(id) {
  // Legacy — now just toggles and sets primary
  if (!openCalendarIds.includes(String(id))) toggleCalendar(id);
  setPrimaryCalendar(id);
}

function deleteCalendar(id) {
  const cal = calendars.find(c => String(c.id) === String(id));
  const name = cal ? cal.name : 'this calendar';
  if (!confirm('Delete "' + name + '" and all its appointments?')) return;
  // Optimistic UI update
  calendars = calendars.filter(c => String(c.id) !== String(id));
  events    = events.filter(e => String(e.calendarId) !== String(id));
  openCalendarIds = openCalendarIds.filter(oid => String(oid) !== String(id));
  if (String(primaryCalendarId) === String(id)) {
    primaryCalendarId = openCalendarIds.length ? openCalendarIds[0] : (calendars.length ? String(calendars[0].id) : null);
    if (primaryCalendarId && !openCalendarIds.includes(primaryCalendarId)) openCalendarIds.push(primaryCalendarId);
  }
  notifyOpenCalendars();
  renderSidebar();
  render();
  queueCommand({ action: 'deleteCalendar', id: String(id) });
}

// ── Add-group modal ──────────────────────────────────────────────────────────
function openGroupModal(groupId) {
  editingGroupId = groupId || null;
  if (editingGroupId) {
    const g = calendarGroups.find(g => String(g.id) === String(editingGroupId));
    document.getElementById('group-modal-title').textContent = 'Rename Group';
    document.getElementById('group-name-save').textContent = 'Save';
    document.getElementById('group-name-input').value = g ? g.name : '';
  } else {
    document.getElementById('group-modal-title').textContent = 'New Group';
    document.getElementById('group-name-save').textContent = 'Create';
    document.getElementById('group-name-input').value = '';
  }
  document.getElementById('group-name-overlay').style.display = 'flex';
  setTimeout(() => document.getElementById('group-name-input').focus(), 50);
}
document.getElementById('add-group-btn').onclick = () => openGroupModal();
document.getElementById('group-name-cancel').onclick = () => {
  document.getElementById('group-name-overlay').style.display = 'none';
};
document.getElementById('group-name-save').onclick = () => {
  const name = document.getElementById('group-name-input').value.trim();
  if (!name) { document.getElementById('group-name-input').focus(); return; }
  document.getElementById('group-name-overlay').style.display = 'none';
  if (editingGroupId) {
    const g = calendarGroups.find(g => String(g.id) === String(editingGroupId));
    if (g) g.name = name;
    renderSidebar();
    if (!standalone) queueCommand({ action: 'renameGroup', id: String(editingGroupId), name });
    editingGroupId = null;
  } else {
    if (standalone) {
      calendarGroups.push({ id: 'tmp_g_' + Date.now(), name });
      renderSidebar();
    } else {
      queueCommand({ action: 'addGroup', name });
    }
  }
};
document.getElementById('group-name-input').addEventListener('keydown', e => {
  if (e.key === 'Enter') document.getElementById('group-name-save').click();
});
document.getElementById('group-name-overlay').onclick = e => {
  if (e.target === document.getElementById('group-name-overlay'))
    document.getElementById('group-name-overlay').style.display = 'none';
};

function deleteGroup(id) {
  const g = calendarGroups.find(g => String(g.id) === String(id));
  const name = g ? g.name : 'this group';
  if (!confirm('Delete group "' + name + '"? Calendars in this group will become ungrouped.')) return;
  calendarGroups = calendarGroups.filter(g => String(g.id) !== String(id));
  calendars.forEach(c => { if (String(c.groupId) === String(id)) c.groupId = ''; });
  renderSidebar();
  if (!standalone) queueCommand({ action: 'deleteGroup', id: String(id) });
}

function moveCalendarToGroup(calId, groupId) {
  const cal = calendars.find(c => String(c.id) === String(calId));
  if (!cal) return;
  if (String(cal.groupId || '') === String(groupId || '')) return;
  cal.groupId = groupId || '';
  renderSidebar();
  if (!standalone) queueCommand({ action: 'moveCalendarToGroup', calendarId: String(calId), groupId: String(groupId || '') });
}

function renderCalColorRow() {
  const row = document.getElementById('cal-color-row');
  row.innerHTML = '';
  COLORS.forEach(c => {
    const s = document.createElement('div');
    s.className = 'color-swatch' + (c === editingCalColor ? ' selected' : '');
    s.style.background = c;
    s.onclick = () => { editingCalColor = c; renderCalColorRow(); };
    row.appendChild(s);
  });
}

// ── Add/Edit calendar modal ──────────────────────────────────────────────────
function openCalModal(groupId) {
  editingCalId = null;
  addCalGroupId = groupId || '';
  editingCalColor = COLORS[0];
  document.getElementById('cal-modal-title').textContent = 'New Calendar';
  document.getElementById('cal-name-save').textContent = 'Create';
  document.getElementById('new-cal-name').value = '';
  document.getElementById('new-cal-work-enable').checked = false;
  document.getElementById('new-cal-work-fields').style.display = 'none';
  document.getElementById('new-cal-work-start').value = '09:00';
  document.getElementById('new-cal-work-end').value   = '17:00';
  renderCalColorRow();
  document.getElementById('cal-name-overlay').style.display = 'flex';
  setTimeout(() => document.getElementById('new-cal-name').focus(), 50);
}

function openCalEditModal(calId) {
  const cal = calendars.find(c => String(c.id) === String(calId));
  if (!cal) return;
  editingCalId = calId;
  addCalGroupId = null;
  editingCalColor = cal.color || COLORS[0];
  document.getElementById('cal-modal-title').textContent = 'Edit Calendar';
  document.getElementById('cal-name-save').textContent = 'Save';
  document.getElementById('new-cal-name').value = cal.name || '';
  const hasWork = !!(cal.workStart && cal.workEnd);
  document.getElementById('new-cal-work-enable').checked = hasWork;
  document.getElementById('new-cal-work-fields').style.display = hasWork ? 'grid' : 'none';
  document.getElementById('new-cal-work-start').value = cal.workStart || '09:00';
  document.getElementById('new-cal-work-end').value   = cal.workEnd   || '17:00';
  renderCalColorRow();
  document.getElementById('cal-name-overlay').style.display = 'flex';
  setTimeout(() => document.getElementById('new-cal-name').focus(), 50);
}

document.getElementById('new-cal-work-enable').addEventListener('change', function() {
  const fields = document.getElementById('new-cal-work-fields');
  fields.style.display = this.checked ? 'grid' : 'none';
});
document.getElementById('cal-name-cancel').onclick = () => {
  document.getElementById('cal-name-overlay').style.display = 'none';
  editingCalId = null; addCalGroupId = null;
};
document.getElementById('cal-name-save').onclick = () => {
  const name = document.getElementById('new-cal-name').value.trim();
  if (!name) { document.getElementById('new-cal-name').focus(); return; }
  const enabled = document.getElementById('new-cal-work-enable').checked;
  const ws = enabled ? document.getElementById('new-cal-work-start').value : '';
  const we = enabled ? document.getElementById('new-cal-work-end').value   : '';
  const calColor = editingCalColor || COLORS[0];
  document.getElementById('cal-name-overlay').style.display = 'none';
  if (editingCalId) {
    const cal = calendars.find(c => String(c.id) === String(editingCalId));
    if (cal) { cal.name = name; cal.workStart = ws; cal.workEnd = we; cal.color = calColor; }
    renderSidebar(); render();
    if (!standalone) queueCommand({ action: 'editCalendar', id: String(editingCalId), name, workStart: ws, workEnd: we, color: calColor });
    editingCalId = null;
  } else {
    const groupId = addCalGroupId || '';
    if (standalone) {
      const tmpId = 'tmp_' + Date.now();
      calendars.push({ id: tmpId, name, color: calColor, workStart: ws, workEnd: we, groupId });
      if (!primaryCalendarId) { primaryCalendarId = tmpId; openCalendarIds.push(tmpId); }
      renderSidebar();
    } else {
      queueCommand({ action: 'addCalendar', name, color: calColor, workStart: ws, workEnd: we, groupId });
    }
    addCalGroupId = null;
  }
};
document.getElementById('new-cal-name').addEventListener('keydown', e => {
  if (e.key === 'Enter') document.getElementById('cal-name-save').click();
});
document.getElementById('cal-name-overlay').onclick = e => {
  if (e.target === document.getElementById('cal-name-overlay')) {
    document.getElementById('cal-name-overlay').style.display = 'none';
    editingCalId = null; addCalGroupId = null;
  }
};

// ── All-day toggle ───────────────────────────────────────────────────────────
const alldayToggle = document.getElementById('allday-toggle');
const timeFields   = document.getElementById('time-fields');
alldayToggle.onclick = () => {
  isAllDay = !isAllDay;
  alldayToggle.classList.toggle('on', isAllDay);
  alldayToggle.setAttribute('aria-checked', isAllDay);
  timeFields.style.display = isAllDay ? 'none' : '';
  if (isAllDay) {
    document.getElementById('ev-start').value = '00:00';
    document.getElementById('ev-end').value   = '23:59';
  }
};
function setAllDay(val) {
  isAllDay = val;
  alldayToggle.classList.toggle('on', isAllDay);
  alldayToggle.setAttribute('aria-checked', isAllDay);
  timeFields.style.display = isAllDay ? 'none' : '';
}

// ── Theme toggle ─────────────────────────────────────────────────────────────
document.getElementById('theme-btn').onclick = () => {
  darkMode = !darkMode;
  applyTheme(darkMode, true);
  queueCommand({ action: 'setTheme', dark: darkMode });
};
function applyTheme(dark, persist = true) {
  darkMode = !!dark;
  document.documentElement.setAttribute('data-theme', dark ? 'dark' : '');
  const btn = document.getElementById('theme-btn');
  if (btn) btn.textContent = dark ? '\u2600\uFE0F' : '\uD83C\uDF19';
  const lbl = document.getElementById('theme-label');
  if (lbl) lbl.textContent = dark ? 'Dark Mode' : 'Light Mode';
  if (persist) saveThemePreference(darkMode);
}

// ── Settings (color customization) ───────────────────────────────────────────
const COLOR_STORAGE_KEY = 'accessCalendarColors';

function loadSavedColors() {
  try {
    const raw = localStorage.getItem(COLOR_STORAGE_KEY);
    if (raw) {
      const colors = JSON.parse(raw);
      if (colors.accent) document.documentElement.style.setProperty('--accent', colors.accent);
      if (colors.accentHover) document.documentElement.style.setProperty('--accent-hover', colors.accentHover);
    }
  } catch (e) {}
}

function saveColors(accent, accentHover) {
  try {
    localStorage.setItem(COLOR_STORAGE_KEY, JSON.stringify({ accent, accentHover }));
  } catch (e) {}
}

function resetColors() {
  try { localStorage.removeItem(COLOR_STORAGE_KEY); } catch (e) {}
  document.documentElement.style.removeProperty('--accent');
  document.documentElement.style.removeProperty('--accent-hover');
}

function getCurrentAccent() {
  return getComputedStyle(document.documentElement).getPropertyValue('--accent').trim();
}
function getCurrentAccentHover() {
  return getComputedStyle(document.documentElement).getPropertyValue('--accent-hover').trim();
}

(function initSettings() {
  const btn       = document.getElementById('settings-btn');
  const overlay   = document.getElementById('settings-overlay');
  const mainPick  = document.getElementById('settings-main-color');
  const mainHex   = document.getElementById('settings-main-hex');
  const hoverPick = document.getElementById('settings-accent-hover-color');
  const hoverHex  = document.getElementById('settings-accent-hover-hex');
  const saveBtn   = document.getElementById('settings-save');
  const cancelBtn = document.getElementById('settings-cancel');
  const resetBtn  = document.getElementById('settings-reset');

  function openSettings() {
    const accent = getCurrentAccent();
    const hover  = getCurrentAccentHover();
    mainPick.value  = accent;
    mainHex.value   = accent;
    hoverPick.value = hover;
    hoverHex.value  = hover;
    overlay.style.display = 'flex';
  }

  function closeSettings() { overlay.style.display = 'none'; }

  // Sync pickers ↔ text inputs with live preview
  mainPick.addEventListener('input', () => {
    mainHex.value = mainPick.value;
    document.documentElement.style.setProperty('--accent', mainPick.value);
  });
  mainHex.addEventListener('input', () => {
    if (/^#[0-9a-f]{6}$/i.test(mainHex.value)) {
      mainPick.value = mainHex.value;
      document.documentElement.style.setProperty('--accent', mainHex.value);
    }
  });
  hoverPick.addEventListener('input', () => {
    hoverHex.value = hoverPick.value;
    document.documentElement.style.setProperty('--accent-hover', hoverPick.value);
  });
  hoverHex.addEventListener('input', () => {
    if (/^#[0-9a-f]{6}$/i.test(hoverHex.value)) {
      hoverPick.value = hoverHex.value;
      document.documentElement.style.setProperty('--accent-hover', hoverHex.value);
    }
  });

  btn.addEventListener('click', openSettings);
  cancelBtn.addEventListener('click', () => {
    loadSavedColors();   // revert live preview
    closeSettings();
  });
  saveBtn.addEventListener('click', () => {
    saveColors(mainPick.value, hoverPick.value);
    closeSettings();
  });
  resetBtn.addEventListener('click', () => {
    resetColors();
    openSettings();      // refresh inputs with defaults
  });
  overlay.addEventListener('click', e => {
    if (e.target === overlay) {
      loadSavedColors();
      closeSettings();
    }
  });

  // Apply saved colours on load
  loadSavedColors();
})();

// ── Recurrence helpers ────────────────────────────────────────────────────────

// AFTER — the immediate call still runs to build the DOW row before `app` exists;
// app.recur.init() (called from CalendarApp.init) will rebuild it + wire radio listeners.
function initRecurPanel() {
  const row = document.getElementById('r-dow-row');
  row.innerHTML = '';
  DOW_LABELS.forEach((lbl, i) => {
    const lab = document.createElement('label');
    lab.style.cssText = 'display:flex;align-items:center;gap:3px;font-size:0.8rem;user-select:none;cursor:pointer';
    lab.innerHTML = `<input type="checkbox" data-dow="${i}" style="accent-color:var(--accent);cursor:pointer"> ${lbl}`;
    row.appendChild(lab);
  });
}
initRecurPanel(); // ← keep: builds DOW row before `app` exists


// AFTER — shim; radio listener registration removed (now lives in RecurrencePanel.init)
function showRecurSubPanel(type) { app.recur.showSubPanel(type); }

// AFTER — shim
function getRecurData() { return app.recur.getRecurData(); }

// AFTER — shim
function setRecurData(ev) { app.recur.setRecurData(ev); }

// ── Recurrence action overlay (this occurrence vs all) ────────────────────────
let _recurActionCallback = null;

function showRecurActionOverlay(title, msg, onSingle, onAll, singleLabel, allLabel) {
  document.getElementById('recur-action-title').textContent  = title;
  document.getElementById('recur-action-msg').textContent    = msg;
  document.getElementById('recur-action-single').textContent = singleLabel || 'This occurrence';
  document.getElementById('recur-action-all').textContent    = allLabel    || 'All occurrences';
  _recurActionCallback = { onSingle, onAll };
  document.getElementById('recur-action-overlay').style.display = 'flex';
}

document.getElementById('recur-action-cancel').onclick = () => {
  document.getElementById('recur-action-overlay').style.display = 'none';
};
document.getElementById('recur-action-single').onclick = () => {
  document.getElementById('recur-action-overlay').style.display = 'none';
  if (_recurActionCallback?.onSingle) _recurActionCallback.onSingle();
};
document.getElementById('recur-action-all').onclick = () => {
  document.getElementById('recur-action-overlay').style.display = 'none';
  if (_recurActionCallback?.onAll) _recurActionCallback.onAll();
};
document.getElementById('recur-action-overlay').onclick = e => {
  if (e.target === document.getElementById('recur-action-overlay'))
    document.getElementById('recur-action-overlay').style.display = 'none';
};


// ── Navigation ───────────────────────────────────────────────────────────────
const calEl = document.getElementById('cal');
const label = document.getElementById('current-label');
const pad = n => String(n).padStart(2,'0');
const fmt = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const sameDay = (a,b) => a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate();
function addDays(d,n) { const r=new Date(d); r.setDate(r.getDate()+n); return r; }
function startOfWeek(d)     { const r=new Date(d); r.setDate(r.getDate()-r.getDay()); return r; }
function startOfWorkWeek(d) { const r=new Date(d), day=r.getDay(); r.setDate(r.getDate()-(day===0?6:day-1)); return r; }

document.getElementById('prev-btn').onclick  = () => navigate(-1);
document.getElementById('next-btn').onclick  = () => navigate(1);
document.getElementById('today-btn').onclick = () => { cur=new Date(); render(); };

// ── Month/Year picker ────────────────────────────────────────────────────────
(function initMonthPicker() {
  const popup   = document.getElementById('monthpicker');
  const mpMonth = document.getElementById('mp-month');
  const mpYear  = document.getElementById('mp-year');
  const MNAMES  = ['January','February','March','April','May','June',
                   'July','August','September','October','November','December'];
  MNAMES.forEach((n, i) => {
    const o = document.createElement('option');
    o.value = i; o.textContent = n;
    mpMonth.appendChild(o);
  });

  function openPicker() {
    mpMonth.value = cur.getMonth();
    mpYear.value  = cur.getFullYear();
    popup.classList.add('open');
  }
  function closePicker() { popup.classList.remove('open'); }

  document.getElementById('current-label').addEventListener('click', e => {
    e.stopPropagation();
    popup.classList.contains('open') ? closePicker() : openPicker();
  });
  document.getElementById('mp-go').addEventListener('click', () => {
    cur = new Date(parseInt(mpYear.value, 10), parseInt(mpMonth.value, 10), 1);
    closePicker();
    render();
  });
  document.getElementById('mp-cancel').addEventListener('click', closePicker);
  popup.addEventListener('click', e => e.stopPropagation());
  document.addEventListener('click', () => closePicker());
})();
document.querySelectorAll('.view-tab').forEach(t => { t.onclick = () => { view=t.dataset.view; render(); }; });
document.getElementById('w5-btn').onclick = () => { weekDays=5; render(); };
document.getElementById('w7-btn').onclick = () => { weekDays=7; render(); };

function navigate(dir) {
  if      (view==='month') cur.setMonth(cur.getMonth()+dir);
  else if (view==='week')  cur.setDate(cur.getDate()+dir*weekDays);
  else                     cur.setDate(cur.getDate()+dir);
  render();
}

document.addEventListener('click', e => {
  if (e.target.classList.contains('wdate')) {
    const dayEl = e.target.closest('.week-head-day');
    if (dayEl) {
      const idx = [...dayEl.parentElement.children].indexOf(dayEl) - 1;
      const ws  = weekDays===5 ? startOfWorkWeek(cur) : startOfWeek(cur);
      cur = addDays(ws, idx);
      view = 'day';
      render();
    }
  }
});

// ── Modal ────────────────────────────────────────────────────────────────────
const overlay = document.getElementById('modal-overlay');
function openModal(dateStr, startStr='09:00', endStr='10:00', ev=null) {
  // Resolve virtual occurrence ID → numeric master ID so saves hit the right DB row
  editingId        = ev ? (ev.recurMasterId != null ? ev.recurMasterId : ev.id) : null;
  editingRecurDate = ev && ev.isRecurring ? ev.date : null
  editingRecurEv   = (ev && ev.isRecurring) ? ev : null
  document.getElementById('modal-title').textContent = ev ? 'Edit Appointment' : 'New Appointment';
  // Populate calendar dropdown with open calendars
  const calSel = document.getElementById('ev-calendar');
  calSel.innerHTML = openCalendarIds.map(cid => {
    const c = calendars.find(cc => String(cc.id) === String(cid));
    return c ? `<option value="${c.id}">${c.name}</option>` : '';
  }).join('');
  calSel.value = ev ? String(ev.calendarId) : String(primaryCalendarId);
  document.getElementById('ev-title').value = ev ? ev.title : '';
  document.getElementById('ev-date').value     = ev ? ev.date  : dateStr;
  document.getElementById('ev-end-date').value = ev ? (ev.endDate || ev.date || dateStr) : dateStr;
  document.getElementById('ev-notes').value = ev ? (ev.notes||'') : '';
  setAllDay(ev ? !!ev.allDay : false);
  document.getElementById('ev-start').value = ev ? ev.start : startStr;
  document.getElementById('ev-end').value   = ev ? ev.end   : endStr;
  selColor = ev ? ev.color : COLORS[0];
  renderColorRow();
  setRecurData(ev);  // populate recurrence panel when editing an existing event
  // Set reminder
  document.getElementById('ev-reminder').value =
      ev ? (ev.reminderMinutes !== undefined && ev.reminderMinutes !== null && ev.reminderMinutes >= 0
            ? String(ev.reminderMinutes) : '-1')
         : '15';
  document.getElementById('modal-delete').style.display = ev ? '' : 'none';
  overlay.style.display = 'flex';
  setTimeout(() => document.getElementById('ev-title').focus(), 50);
}
function renderColorRow() {
  const row = document.getElementById('color-row');
  row.innerHTML = '';
  COLORS.forEach(c => {
    const s = document.createElement('div');
    s.className = 'color-swatch' + (c===selColor?' selected':'');
    s.style.background = c;
    s.onclick = () => { selColor=c; renderColorRow(); };
    row.appendChild(s);
  });
}
document.getElementById('modal-cancel').onclick = () => { overlay.style.display='none'; };
overlay.onclick = e => { if (e.target===overlay) overlay.style.display='none'; };

document.getElementById('modal-delete').onclick = () => {
  if (editingId == null) return;
  overlay.style.display = 'none';

  if (editingRecurDate !== null) {
    // Recurring event — ask what scope
    showRecurActionOverlay(
      'Delete Recurring Event',
      'Do you want to delete just this occurrence or all events in the series?',
      () => {
        // This occurrence only
        if (!standalone)
          queueCommand({action:'deleteOccurrence', id: editingId, date: editingRecurDate});
        events = events.filter(e => !(String(e.recurMasterId) === String(editingId) && e.date === editingRecurDate));
        render();
      },
      () => {
        // All occurrences
        if (!standalone)
          queueCommand({action:'deleteSeries', id: editingId});
        events = events.filter(e => String(e.recurMasterId) !== String(editingId) && e.id != editingId);
        render();
      }
    );
  } else {
    if (standalone) {
      events = events.filter(e => e.id !== editingId);
    } else {
      queueCommand({action:'delete', id: editingId});
      events = events.filter(e => e.id != editingId);
    }
    render();
  }
};


document.getElementById('modal-save').onclick = () => {
    const title   = document.getElementById('ev-title').value.trim() || 'No title'
    const date    = document.getElementById('ev-date').value
    const endDate = document.getElementById('ev-end-date').value
    const start   = document.getElementById('ev-start').value
    const end     = document.getElementById('ev-end').value
    const notes   = document.getElementById('ev-notes').value.trim()
    const allDay  = isAllDay
    if (!date) return
    const calId = document.getElementById('ev-calendar').value || primaryCalendarId
    const rd    = getRecurData()

    // ── Recurring-occurrence reschedule ──────────────────────────────────────
    // When the user opens a single occurrence of a recurring series and clicks
    // Save, ask whether they mean THIS occurrence only or ALL occurrences.
    //   "This occurrence"  → commitRescheduleOne():
    //       1. Adds the original date to master's RecurExceptions (suppresses it)
    //       2. Creates a new standalone appointment at the new date/time
    //   "All occurrences"  → doSaveAll() updates the master record as before
    if (editingRecurDate !== null && editingId !== null && editingRecurEv !== null) {
        overlay.style.display = 'none'
        showRecurActionOverlay(
            'Edit Recurring Event',
            'Update only this occurrence, or every event in the series?',
            () => {
                // This occurrence — reschedule to the (possibly changed) date/time
                commitRescheduleOne(
                    editingRecurEv,
                    date, start, end,
                    { title, color: selColor, notes, allDay,
                      reminderMinutes: parseInt(document.getElementById('ev-reminder').value, 10) }
                )
            },
            () => {
                // All occurrences — update the master record as before
                doSaveAll()
            },
            'This occurrence',
            'All occurrences'
        )
        return
    }

    doSaveAll()

    // ── Helper: save / update the master record (existing behaviour) ─────────
    function doSaveAll() {
        if (standalone) {
            if (editingId !== null) {
                const ev = events.find(e => e.id === editingId)
                if (ev) Object.assign(ev, {title,date,endDate,start,end,notes,allDay,color:selColor,...rd})
            } else {
                events.push({id:Date.now(),title,date,endDate,start,end,notes,allDay,color:selColor,calendarId:calId,...rd})
            }
        } else {
            queueCommand({ action:'save', id: editingId||0, calendarId:calId,
                title, date, endDate, start, end, notes, allDay, color:selColor,
                reminderMinutes: parseInt(document.getElementById('ev-reminder').value, 10),
                recurType:         rd.recurType,
                recurInterval:     rd.recurInterval    || 1,
                recurDaysOfWeek:   rd.recurDaysOfWeek  || '',
                recurMonthlyMode:  rd.recurMonthlyMode || 'day',
                recurMonthDay:     rd.recurMonthDay    || 1,
                recurMonthWeek:    rd.recurMonthWeek   || 1,
                recurMonthDOW:     rd.recurMonthDOW    || 0,
                recurEndType:      rd.recurEndType     || 'never',
                recurEndDate:      rd.recurEndDate     || '',
                recurCount:        rd.recurCount       || 0,
                recurRangeStart:   rd.recurRangeStart  || date })
            // Optimistic clear stale events — keep UI responsive while VBA reloads
            if (editingId !== null) {
                const hadRecurrence = events.some(e => String(e.recurMasterId) === String(editingId))
                events = events.filter(e => e.id !== editingId && String(e.recurMasterId) !== String(editingId))
                // For non-recurring edits OR converting recurring→single: keep event visible
                // with updated values as a placeholder while VBA reloads
                if (!hadRecurrence || rd.recurType === 'none')
                    events.push({id:`tmp${Date.now()}`,title,date,endDate,start,end,
                                 notes,allDay,color:selColor,calendarId:calId,recurType:'none'})
                // For recurring→still-recurring edits: no placeholder; VBA must rebuild all occurrences
            } else {
                events.push({id:`tmp${Date.now()}`,title,date,endDate,start,end,
                             notes,allDay,color:selColor,calendarId:calId,...rd})
            }
        }
        overlay.style.display = 'none'
        render()
    }
}



// ── Drag ─────────────────────────────────────────────────────────────────────
function commitMove(ev, newDate, newStart, newEnd) {
  let newEndDate = '';
  if (ev.endDate && ev.endDate > ev.date) {
    const durMs   = new Date(ev.endDate+'T00:00:00') - new Date(ev.date+'T00:00:00');
    newEndDate    = fmt(new Date(new Date(newDate+'T00:00:00').getTime() + durMs));
  }
  if (standalone) {
    ev.date=newDate; ev.endDate=newEndDate; ev.start=newStart; ev.end=newEnd; render();
  } else {
    queueCommand({action:'move',id:ev.id,date:newDate,endDate:newEndDate,start:newStart,end:newEnd});
    ev.date=newDate; ev.endDate=newEndDate; ev.start=newStart; ev.end=newEnd; render();
  }
}


// commitResize: only changes start/end time, not the date.
// Handles recurring events by sending recurMasterId so VBA updates the master record.
function commitResize(ev, newStart, newEnd) {
  if (standalone) {
    ev.start = newStart; ev.end = newEnd; render();
  } else {
    const masterId = ev.recurMasterId != null ? String(ev.recurMasterId) : '';
    queueCommand({action:'resize', id:String(ev.id), masterId:masterId, start:newStart, end:newEnd});
    ev.start = newStart; ev.end = newEnd; render();
  }
}

// commitRescheduleOne: exempts one occurrence from its recurring series and
// books a new standalone appointment with the modified date/time.
// overrides = { title, color, notes, allDay } — supplied when called from the modal
// so user-edited values flow through into the new standalone record.
function commitRescheduleOne(ev, newDate, newStart, newEnd, overrides = {}) {
    const masterId      = ev.recurMasterId != null ? ev.recurMasterId : ev.id
    const title         = overrides.title         !== undefined ? overrides.title         : ev.title
    const color         = overrides.color         !== undefined ? overrides.color         : (ev.color  || '#4f46e5')
    const notes         = overrides.notes         !== undefined ? overrides.notes         : (ev.notes  || '')
    const location      = overrides.location      !== undefined ? overrides.location      : (ev.location || '')
    const allDay        = overrides.allDay        !== undefined ? overrides.allDay        : (ev.allDay || false)
    const reminderMins  = overrides.reminderMinutes !== undefined ? overrides.reminderMinutes
                          : (ev.reminderMinutes !== undefined ? ev.reminderMinutes : -1)
    if (standalone) {
        events = events.filter(e => !(String(e.recurMasterId) === String(masterId) && e.date === ev.date))
        events.push({
            id: Date.now(), calendarId: ev.calendarId,
            title, date: newDate, endDate: '', start: newStart, end: newEnd,
            allDay, color, notes, location, recurType: 'none', isRecurring: false,
            reminderMinutes: reminderMins
        })
        render()
    } else {
        queueCommand({
            action: 'rescheduleOccurrence',
            masterId: String(masterId),
            originalDate: ev.date,
            newDate, start: newStart, end: newEnd,
            title, calendarId: String(ev.calendarId),
            color, notes, location,
            allDay: allDay ? true : false,
            reminderMinutes: reminderMins
        })
        // Optimistic: hide the virtual occurrence immediately while VBA rebuilds
        events = events.filter(e => !(String(e.recurMasterId) === String(masterId) && e.date === ev.date))
        render()
    }
}


// ── Render ───────────────────────────────────────────────────────────────────
function render() {
  document.querySelectorAll('.view-tab').forEach(t => t.classList.toggle('active', t.dataset.view===view));
  document.getElementById('week-toggle-wrap').style.display = view==='week' ? '' : 'none';
  document.getElementById('w5-btn').classList.toggle('active', weekDays===5);
  document.getElementById('w7-btn').classList.toggle('active', weekDays===7);
  if      (view==='month') renderMonth();
  else if (view==='week')  renderWeek();
  else                     renderDay();
  renderMiniCal();
  renderCalTabs();
}

// ── Month view ────────────────────────────────────────────────────────────────
function renderMonth() {
  const y=cur.getFullYear(), m=cur.getMonth();
  const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
  label.textContent = MONTHS[m]+' '+y;
  const first=new Date(y,m,1), start=startOfWeek(first);
  const days  = Array.from({length:42},(_,i)=>addDays(start,i));
  const evs   = calEvents();
  const mdEvs = evs.filter(e => e.endDate && e.endDate > e.date);   // multiday
  const sdEvs = evs.filter(e => !e.endDate || e.endDate <= e.date); // single-day

  // Build 6 week-row wrappers (position:relative lets banners span cells)
  const weeksHtml = Array.from({length:6},(_,w) => {
    const wDays = days.slice(w*7, w*7+7);
    return '<div class="week-row" data-week="'+w+'">' +
      wDays.map(day => {
        const ds=fmt(day), isToday=sameDay(day,today), other=day.getMonth()!==m;
        const dayEvs = sdEvs.filter(e=>e.date===ds).sort((a,b)=>a.start.localeCompare(b.start));
        const visible=dayEvs.slice(0,3), more=dayEvs.length-3;
        return '<div class="month-cell'+(other?' other-month':'')+(isToday?' today':'')+'" data-date="'+ds+'">' +
          '<div class="day-num">'+day.getDate()+'</div>' +
          '<div class="md-spacer"></div>' +
          visible.map(ev=>'<div class="month-event'+(isEvSelected(ev)?' ev-selected':'')+'" style="background:'+eventDisplayColor(ev)+';--ev-accent:'+eventAccentColor(ev)+
            '" data-id="'+ev.id+'" draggable="true">'+(ev.isRecurring?'↻ ':'')+esc(ev.title)+'</div>').join('') +
          (more>0?'<div class="more-events">+'+more+' more</div>':'') +
          '</div>';
      }).join('') +
      '</div>';
  }).join('');

  calEl.innerHTML =
    '<div class="month-header">'+
    ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'].map(n=>'<div>'+n+'</div>').join('')+
    '</div><div class="month-grid" id="month-grid">'+weeksHtml+'</div>';

  // ── Inject multiday banners ──────────────────────────────────────────────
  document.querySelectorAll('.week-row').forEach((wRow, w) => {
    const wDays     = days.slice(w*7, w*7+7);
    const wStartStr = fmt(wDays[0]), wEndStr = fmt(wDays[6]);
    const weekEvs   = mdEvs
      .filter(ev => ev.date <= wEndStr && ev.endDate >= wStartStr)
      .sort((a,b) => a.date.localeCompare(b.date));
    if (!weekEvs.length) return;

    // Greedy slot allocation – prevents banners overlapping vertically
    const slotLastCol = [];
    const assigns = [];
    weekEvs.forEach(ev => {
      const segS = ev.date    < wStartStr ? wStartStr : ev.date;
      const segE = ev.endDate > wEndStr   ? wEndStr   : ev.endDate;
      const colS = wDays.findIndex(d => fmt(d) === segS);
      const colE = wDays.findIndex(d => fmt(d) === segE);
      if (colS < 0) return;
      const effColE = colE >= 0 ? colE : wDays.length - 1;
      let slot = slotLastCol.findIndex(last => last < colS);
      if (slot < 0) { slot = slotLastCol.length; slotLastCol.push(effColE); }
      else slotLastCol[slot] = effColE;
      assigns.push({ ev, colS, colE: effColE, slot,
        isPartial: ev.date < wStartStr, continues: ev.endDate > wEndStr });
    });

    // Resize spacers so single-day pills render below the banners
    const maxSlotPerCol = Array(7).fill(-1);
    assigns.forEach(({colS, colE, slot}) => {
      for (let c=colS; c<=colE; c++) maxSlotPerCol[c] = Math.max(maxSlotPerCol[c], slot);
    });
    wRow.querySelectorAll('.month-cell').forEach((cell, ci) => {
      const sp = cell.querySelector('.md-spacer');
      if (sp) sp.style.height = maxSlotPerCol[ci] >= 0 ? ((maxSlotPerCol[ci]+1)*22+2)+'px' : '0';
    });

    // Render the banner elements
    const pct = 100/7;
    assigns.forEach(({ev, colS, colE, slot, isPartial, continues}) => {
      const span   = colE - colS + 1;
      const banner = document.createElement('div');
      let   cls    = 'md-banner';
      if (isPartial) cls += ' md-cont-left';
      if (continues) cls += ' md-cont-right';
      if (isEvSelected(ev)) cls += ' ev-selected';
      banner.className       = cls;
      banner.style.background = eventDisplayColor(ev);
      banner.style.setProperty('--ev-accent', eventAccentColor(ev));
      banner.style.top       = (24 + slot*22)+'px';
      banner.style.left      = 'calc('+(colS*pct)+'% + 2px)';
      banner.style.width     = 'calc('+(span*pct)+'% - 4px)';
      banner.dataset.id      = String(ev.id);
      banner.textContent     = esc(ev.title);
      wRow.appendChild(banner);
    });
  });

  // ── Event listeners ──────────────────────────────────────────────────────
  calEl.querySelectorAll('.month-cell').forEach(cell => {
    cell.addEventListener('click', e => {
      if (e.target.classList.contains('month-event') || e.target.classList.contains('more-events')) return;
      openModal(cell.dataset.date);
    });
    cell.addEventListener('dragover',  e => { e.preventDefault(); cell.style.background='var(--accent-hover)'; });
    cell.addEventListener('dragleave', () => { cell.style.background=''; });
            cell.addEventListener('drop', e => {
                e.preventDefault()
                cell.style.background = ''
                if (!dragEv) return
                if (dragEv.isRecurring) {
                    const snapEv = dragEv
                    const droppedDate = cell.dataset.date
                    showRecurActionOverlay(
                        'Reschedule Recurring Event',
                        'This is one appointment in a series. What do you want to reschedule?',
                        () => commitRescheduleOne(snapEv, droppedDate, snapEv.start, snapEv.end),
                        () => commitMove(snapEv, droppedDate, snapEv.start, snapEv.end),
                        'Just this one',
                        'The entire series'
                    )
                } else {
                    commitMove(dragEv, cell.dataset.date, dragEv.start, dragEv.end)
                }
            })
          });            
  calEl.querySelectorAll('.month-event').forEach(el => {
    el.addEventListener('click', e => {
      e.stopPropagation();
      const ev=events.find(e2=>String(e2.id)===el.dataset.id);
      if(ev) openModal('','','',ev);
    });
    el.addEventListener('dragstart', e => {
      dragEv=events.find(e2=>String(e2.id)===el.dataset.id);
      e.dataTransfer.effectAllowed='move';
    });
    el.addEventListener('dragend', () => { dragEv=null; });
  });
  calEl.querySelectorAll('.md-banner').forEach(el => {
    el.addEventListener('click', e => {
      e.stopPropagation();
      const ev=events.find(e2=>String(e2.id)===el.dataset.id);
      if(ev) openModal('','','',ev);
    });
  });
}

// ── Work-hours shading helpers ────────────────────────────────────────────
function getActiveCalWorkHours() {
  const cal = calendars.find(c => String(c.id) === String(primaryCalendarId));
  if (!cal) return null;
  const ws = parseInt((cal.workStart || '').split(':')[0]);
  const we = parseInt((cal.workEnd   || '').split(':')[0]);
  if (isNaN(ws) || isNaN(we) || ws >= we) return null;
  return { start: ws, end: we };
}
function offHoursCls(h, wh) {
  // Highlight the WORK-HOURS band; non-work hours show the base white
  return wh && (h >= wh.start && h < wh.end) ? ' off-hours' : '';
}

// ── Week view ─────────────────────────────────────────────────────────────────
function renderWeek() {
  const sw = weekDays===5 ? startOfWorkWeek(cur) : startOfWeek(cur);
  const dates = Array.from({length:weekDays},(_,i)=>addDays(sw,i));
  const MONTHS=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  label.textContent = dates[0].getMonth()!==dates[dates.length-1].getMonth()
    ? `${MONTHS[dates[0].getMonth()]} \u2013 ${MONTHS[dates[dates.length-1].getMonth()]} ${dates[0].getFullYear()}`
    : `${MONTHS[dates[0].getMonth()]} ${dates[0].getFullYear()}`;
  const DAYS=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const wh = getActiveCalWorkHours();
  calEl.innerHTML =
    `<div class="week-container">` +
    `<div class="week-head"><div class="week-head-time"></div>${dates.map(d=>
      `<div class="week-head-day"><div class="wday">${DAYS[d.getDay()]}</div>` +
      `<div class="wdate${sameDay(d,today)?' today':''}">${d.getDate()}</div></div>`).join('')}</div>` +
    `<div class="allday-row"><div class="allday-gutter">All&#8209;day</div>${dates.map(d=>
      `<div class="allday-col" data-allday="${fmt(d)}"></div>`).join('')}</div>` +
    `<div class="week-body" id="week-body">` +
    `<div class="time-col">${Array.from({length:24},(_,h)=>`<div class="time-slot${offHoursCls(h,wh)}">${h===0?'':h<12?h+'am':h===12?'12pm':(h-12)+'pm'}</div>`).join('')}</div>` +
    `<div class="days-col" id="days-col">${dates.map(d=>
      `<div class="day-col" data-date="${fmt(d)}">${Array.from({length:24},(_,h)=>`<div class="hour-row${offHoursCls(h,wh)}" data-hour="${h}"></div>`).join('')}</div>`).join('')}</div>` +
    `</div></div>`;
  renderWeekEvents(dates);
  renderAlldayEvents(dates);
  setupAlldayInteraction();
  setupWeekInteraction();
  scrollToNow();
}

// ── Day view ──────────────────────────────────────────────────────────────────
function renderDay() {
  const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const DAYS=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  label.textContent = `${DAYS[cur.getDay()]}, ${MONTHS[cur.getMonth()]} ${cur.getDate()}, ${cur.getFullYear()}`;
  const wh = getActiveCalWorkHours();
  calEl.innerHTML =
    `<div class="week-container">` +
    `<div class="week-head"><div class="week-head-time"></div>` +
    `<div class="week-head-day"><div class="wday">${DAYS[cur.getDay()].slice(0,3).toUpperCase()}</div>` +
    `<div class="wdate${sameDay(cur,today)?' today':''}">${cur.getDate()}</div></div></div>` +
    `<div class="allday-row"><div class="allday-gutter">All&#8209;day</div>` +
    `<div class="allday-col" data-allday="${fmt(cur)}"></div></div>` +
    `<div class="week-body" id="week-body">` +
    `<div class="time-col">${Array.from({length:24},(_,h)=>`<div class="time-slot${offHoursCls(h,wh)}">${h===0?'':h<12?h+'am':h===12?'12pm':(h-12)+'pm'}</div>`).join('')}</div>` +
    `<div class="days-col" id="days-col">` +
    `<div class="day-col" data-date="${fmt(cur)}">${Array.from({length:24},(_,h)=>`<div class="hour-row${offHoursCls(h,wh)}" data-hour="${h}"></div>`).join('')}</div>` +
    `</div></div></div>`;
  renderWeekEvents([new Date(cur)]);
  renderAlldayEvents([new Date(cur)]);
  setupAlldayInteraction();
  setupWeekInteraction();
  scrollToNow();
}


// ── All-day strip drag-and-drop interaction ───────────────────────────────────
function setupAlldayInteraction() {
  document.querySelectorAll('.allday-col').forEach(col => {
    col.addEventListener('dragover', e => {
      if (!dragEv || !dragEv.allDay) return;   // only accept all-day drags
      e.preventDefault();
      col.classList.add('drag-over');
    });
    col.addEventListener('dragleave', () => col.classList.remove('drag-over'));
    col.addEventListener('drop', e => {
      e.preventDefault();
      col.classList.remove('drag-over');
      if (!dragEv || !dragEv.allDay) return;
      const newDate = col.dataset.allday;
      if (newDate === dragEv.date) return;      // dropped on same day — no-op
      commitMove(dragEv, newDate, dragEv.start, dragEv.end);
    });
  });
}

// ── Shared week/day event rendering ──────────────────────────────────────────
function timeToY(t) { const [h,m]=t.split(':').map(Number); return (h*60+m)*(56/60); }
function yToTime(y) {
  const mins=Math.round(y*(60/56)/15)*15, h=Math.floor(mins/60), m=mins%60;
  return pad(Math.min(h,23))+':'+pad(m);
}
// ── Overlap layout: assign columns to concurrent events ─────────────────────
function layoutOverlappingEvents(dayEvents) {
  if (!dayEvents.length) return [];
  const items = dayEvents.map(ev => {
    const top = timeToY(ev.start);
    const bot = ev.end ? timeToY(ev.end) : top + 56;
    return { ev, top, bot: Math.max(bot, top + 22), col: 0, totalCols: 1 };
  }).sort((a, b) => a.top - b.top || (b.bot - b.top) - (a.bot - a.top));

  // Group into clusters of mutually overlapping events
  const clusters = [];
  let cluster = [items[0]];
  let clusterEnd = items[0].bot;
  for (let i = 1; i < items.length; i++) {
    if (items[i].top < clusterEnd) {
      cluster.push(items[i]);
      clusterEnd = Math.max(clusterEnd, items[i].bot);
    } else {
      clusters.push(cluster);
      cluster = [items[i]];
      clusterEnd = items[i].bot;
    }
  }
  clusters.push(cluster);

  // For each cluster, greedily assign columns
  clusters.forEach(group => {
    const columns = [];  // each column tracks the end-y of its last event
    group.forEach(item => {
      let placed = false;
      for (let c = 0; c < columns.length; c++) {
        if (item.top >= columns[c]) {
          item.col = c;
          columns[c] = item.bot;
          placed = true;
          break;
        }
      }
      if (!placed) {
        item.col = columns.length;
        columns.push(item.bot);
      }
    });
    const totalCols = columns.length;
    group.forEach(item => { item.totalCols = totalCols; });
  });
  return items;
}

function renderWeekEvents(dates) {
  const daysCol = document.getElementById('days-col');
  if (!daysCol) return;
  const nowCol = daysCol.querySelector(`[data-date="${fmt(today)}"]`);
  if (nowCol) {
    const now=new Date(), line=document.createElement('div');
    line.className='now-line';
    line.style.top=(now.getHours()*60+now.getMinutes())*(56/60)+'px';
    nowCol.appendChild(line);
  }
  const evs = calEvents();
  dates.forEach(d => {
    const ds=fmt(d), col=daysCol.querySelector(`[data-date="${ds}"]`);
    if (!col) return;
    const dayEvs = evs.filter(e=>e.date===ds && !e.allDay && (!e.endDate||e.endDate<=e.date));
    const laid = layoutOverlappingEvents(dayEvs);
    laid.forEach(({ev, top, bot, col: evCol, totalCols}) => {
      const h=Math.max(bot-top,22);
      const pct = 100 / totalCols;
      const leftPct = evCol * pct;
      const el=document.createElement('div');
      el.className='week-event'+(isEvSelected(ev)?' ev-selected':'');
      el.draggable=true;
      el.style.cssText=`top:${top}px;height:${h}px;background:${eventDisplayColor(ev)};--ev-accent:${eventAccentColor(ev)};left:calc(${leftPct}% + 1px);width:calc(${pct}% - 2px)`;
      el.dataset.id=String(ev.id);
      el.innerHTML=`<div class="ev-title">${ev.isRecurring?'↻ ':''}${esc(ev.title)}</div><div class="ev-time">${ev.start}${ev.end?' \u2013 '+ev.end:''}</div>`;
      el.addEventListener('click', e => {
        e.stopPropagation();
        const found=events.find(e2=>String(e2.id)===el.dataset.id);
        if(found) openModal('','','',found);
      });
      el.addEventListener('dragstart', e => {
        if (resizeEv) { e.preventDefault(); return; }
        dragEv=events.find(e2=>String(e2.id)===el.dataset.id);
        dragOffset=e.clientY-el.getBoundingClientRect().top;
        e.dataTransfer.effectAllowed='move';
      });
      el.addEventListener('dragend', () => { dragEv=null; });

      // ── Resize handles ──────────────────────────────────────────────────────
      const topHandle = document.createElement('div');
      topHandle.className = 'ev-resize-top';
      topHandle.title = 'Drag to change start time';
      const botHandle = document.createElement('div');
      botHandle.className = 'ev-resize-bottom';
      botHandle.title = 'Drag to change end time';

      function startResize(edge, e) {
        e.stopPropagation();
        e.preventDefault();
        const found = events.find(e2 => String(e2.id) === el.dataset.id);
        if (!found) return;
        resizeEv = found;
        resizeEdge = edge;
        resizeEl = el;
        resizeCol = col;
        resizeOrigStart = found.start;
        resizeOrigEnd = found.end || found.start;
        el.classList.add('resizing');
        document.body.classList.add('resizing-active');
        el.draggable = false;
      }
      topHandle.addEventListener('mousedown', e => startResize('top', e));
      botHandle.addEventListener('mousedown', e => startResize('bottom', e));
      el.appendChild(topHandle);
      el.appendChild(botHandle);
      col.appendChild(el);
    });
  });
}

// ── All-day event strip renderer (week & day views) ──────────────────────────
function renderAlldayEvents(dates) {
  const evs         = calEvents();
  const visStartStr = fmt(dates[0]);
  const visEndStr   = fmt(dates[dates.length-1]);

  // ── Single-day allDay events (original behaviour) ────────────────────────
  const singleAlldayEvs = evs.filter(e => e.allDay && (!e.endDate || e.endDate <= e.date));
  dates.forEach(d => {
    const ds  = fmt(d);
    const col = document.querySelector(`.allday-col[data-allday="${ds}"]`);
    if (!col) return;
    singleAlldayEvs.filter(e => e.date === ds).forEach(ev => {
      const el = document.createElement('div');
      el.className  = 'allday-event'+(isEvSelected(ev)?' ev-selected':'');
      el.style.background = eventDisplayColor(ev);
      el.style.setProperty('--ev-accent', eventAccentColor(ev));
      el.dataset.id = String(ev.id);
      el.textContent = ev.title;
      el.draggable = true;
      el.addEventListener('click', e => {
        e.stopPropagation();
        const found = events.find(e2 => String(e2.id) === el.dataset.id);
        if (found) openModal('', '', '', found);
      });
      el.addEventListener('dragstart', e => {
        dragEv = events.find(e2 => String(e2.id) === el.dataset.id);
        e.dataTransfer.effectAllowed = 'move';
        requestAnimationFrame(() => el.classList.add('dragging'));
      });
      el.addEventListener('dragend', () => {
        el.classList.remove('dragging');
        dragEv = null;
      });
      col.appendChild(el);
    });
  });

  // ── Multiday events – spanning banner in the allday strip ────────────────
  // Width trick: each allday-col is flex:1 (equal share). A child with
  // width: calc(N * 100%) grows to span N columns. z-index keeps it on top.
  const mdEvs = evs.filter(e => e.endDate && e.endDate > e.date);
  mdEvs.forEach(ev => {
    if (ev.date > visEndStr || ev.endDate < visStartStr) return;
    const segStartStr = ev.date    < visStartStr ? visStartStr : ev.date;
    const segEndStr   = ev.endDate > visEndStr   ? visEndStr   : ev.endDate;
    const colStartIdx = dates.findIndex(d => fmt(d) === segStartStr);
    if (colStartIdx < 0) return;
    let colEndIdx = dates.findIndex(d => fmt(d) === segEndStr);
    if (colEndIdx < 0) colEndIdx = dates.length - 1;
    const spanCount = colEndIdx - colStartIdx + 1;
    const allCols   = document.querySelectorAll('.allday-col');
    const startCol  = allCols[colStartIdx];
    if (!startCol) return;

    const el = document.createElement('div');
    el.className       = 'allday-event md-allday-span'+(isEvSelected(ev)?' ev-selected':'');
    el.style.background = eventDisplayColor(ev);
    el.style.setProperty('--ev-accent', eventAccentColor(ev));
    el.style.width     = 'calc('+spanCount+' * 100% + '+(spanCount-1)*2+'px)';
    el.dataset.id      = String(ev.id);
    el.textContent     = (ev.date < visStartStr ? '\u2039 ' : '') +
                          ev.title +
                         (ev.endDate > visEndStr ? ' \u203A' : '');
    el.addEventListener('click', e => {
      e.stopPropagation();
      const found = events.find(e2 => String(e2.id) === el.dataset.id);
      if (found) openModal('', '', '', found);
    });
    startCol.appendChild(el);
  });
}

function setupWeekInteraction() {
  const daysCol = document.getElementById('days-col');
  if (!daysCol) return;
  daysCol.querySelectorAll('.day-col').forEach(col => {
    col.addEventListener('click', e => {
      if (e.target.closest('.week-event')) return;
      const rect=col.getBoundingClientRect(), y=e.clientY-rect.top;
      const h=Math.floor(y/56), m=y%56>=28?30:0;
      openModal(col.dataset.date, pad(h)+':'+pad(m), pad(Math.min(h+1,23))+':'+pad(m));
    });
    col.addEventListener('dragover',  e => {
      if (!dragEv || dragEv.allDay) return;    // reject all-day drags on timed grid
      e.preventDefault(); col.style.background='var(--drag-over)';
    });
    col.addEventListener('dragleave', () => { col.style.background=''; });
    col.addEventListener('drop', e => {
      e.preventDefault(); col.style.background='';
      if (!dragEv || dragEv.allDay) return;    // reject all-day drags on timed grid
      const rect=col.getBoundingClientRect(), y=Math.max(0,e.clientY-rect.top-dragOffset);
      const newStart=yToTime(y);
      const [sh,sm]=dragEv.start.split(':').map(Number);
      const [eh,em]=(dragEv.end||dragEv.start).split(':').map(Number);
      const dur=(eh*60+em)-(sh*60+sm);
      const [nh,nm]=newStart.split(':').map(Number);
      const endMins=nh*60+nm+dur;
      const newEnd=pad(Math.floor(endMins/60))+':'+pad(endMins%60);
      const droppedDate=col.dataset.date;
        if (dragEv.isRecurring) {
            const snapEv = dragEv
            showRecurActionOverlay(
                'Reschedule Recurring Event',
                'This is one appointment in a series. What do you want to reschedule?',
                () => commitRescheduleOne(snapEv, droppedDate, newStart, newEnd),
                () => commitMove(snapEv, droppedDate, newStart, newEnd),
                'Just this one',
                'The entire series'
            )
        } else {
            commitMove(dragEv, droppedDate, newStart, newEnd)
        }
    });
  });
}
function scrollToNow() {
  const body=document.getElementById('week-body');
  if (!body) return;
  setTimeout(() => { body.scrollTop=Math.max(0,(new Date().getHours()-1)*56); }, 0);
}

// ── HTML escape helper ────────────────────────────────────────────────────────
function esc(s) { const d=document.createElement('div'); d.textContent=s; return d.innerHTML; }


// ── Mini-month calendar ───────────────────────────────────────────────────────
const MINI_MO = ['January','February','March','April','May','June',
                 'July','August','September','October','November','December'];
const MINI_DW = ['Su','Mo','Tu','We','Th','Fr','Sa'];
let miniY = today.getFullYear();
let miniM = today.getMonth();

function renderMiniCal() {
  const titleEl = document.getElementById('mini-cal-title');
  const grid    = document.getElementById('mini-cal-grid');
  if (!titleEl || !grid) return;

  titleEl.textContent = MINI_MO[miniM] + ' ' + miniY;

  // Compute highlighted week range for week view
  let hlStart = null, hlEnd = null;
  if (view === 'week') {
    hlStart = weekDays === 5 ? startOfWorkWeek(cur) : startOfWeek(cur);
    hlEnd   = addDays(hlStart, weekDays - 1);
  }

  const first = new Date(miniY, miniM, 1);
  const start = startOfWeek(first);
  const days  = Array.from({length:42}, (_, i) => addDays(start, i));

  grid.innerHTML = '';

  // DOW header — empty gutter + 7 labels
  grid.appendChild(document.createElement('div'));
  MINI_DW.forEach(d => {
    const h = document.createElement('div');
    h.className   = 'mini-dow';
    h.textContent = d;
    grid.appendChild(h);
  });

  // 6 week rows
  for (let w = 0; w < 6; w++) {
    const weekStart = new Date(days[w * 7]);

    // Left-gutter › button — click to show that week
    const wb = document.createElement('div');
    wb.className = 'mini-week-btn';
    wb.title     = 'Show this week';
    wb.innerHTML = '&#8250;';
    wb.addEventListener('click', () => {
      cur = new Date(weekStart);
      view = 'week';
      render();
    });
    grid.appendChild(wb);

    // 7 day cells
    for (let d = 0; d < 7; d++) {
      const day  = days[w * 7 + d];
      const cell = document.createElement('div');
      cell.className   = 'mini-day';
      cell.textContent = day.getDate();

      if (day.getMonth() !== miniM) cell.classList.add('mc-other');
      if (sameDay(day, today))       cell.classList.add('mc-today');
      else if (view === 'day' && sameDay(day, cur))
        cell.classList.add('mc-selected');
      else if (view === 'week' && hlStart &&
               day.getTime() >= hlStart.getTime() &&
               day.getTime() <= hlEnd.getTime())
        cell.classList.add('mc-in-week');

      cell.addEventListener('click', () => {
        cur  = new Date(day);
        view = 'day';
        render();
      });
      grid.appendChild(cell);
    }
  }
}

// Mini-cal prev / next
document.getElementById('mini-prev').addEventListener('click', () => {
  if (--miniM < 0) { miniM = 11; miniY--; }
  renderMiniCal();
});
document.getElementById('mini-next').addEventListener('click', () => {
  if (++miniM > 11) { miniM = 0; miniY++; }
  renderMiniCal();
});
// Click title → show that month in main calendar
document.getElementById('mini-cal-title').addEventListener('click', () => {
  cur  = new Date(miniY, miniM, 1);
  view = 'month';
  render();
});

// Init theme before first render
darkMode = readSavedThemePreference();
applyTheme(darkMode, false);
// ── Init ─────────────────────────────────────────────────────────────────────
renderSidebar();
render();

// ── Event hover popover (all views) ──────────────────────────────────────────
(function initEventPopover() {
  const pop = document.getElementById('ev-popover');
  let hoverTimer = null;
  const DAYS_FULL = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const MONTHS_FULL = ['January','February','March','April','May','June','July','August','September','October','November','December'];

  function formatTime12(t) {
    if (!t) return '';
    const [h,m] = t.split(':').map(Number);
    const ampm = h >= 12 ? 'PM' : 'AM';
    const h12 = h === 0 ? 12 : h > 12 ? h - 12 : h;
    return h12 + ':' + String(m).padStart(2,'0') + ' ' + ampm;
  }

  function formatDate(ds) {
    const d = new Date(ds + 'T00:00:00');
    return DAYS_FULL[d.getDay()] + ', ' + MONTHS_FULL[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
  }

  function findEvent(el) {
    const id = el.dataset.id;
    if (id) return events.find(e => String(e.id) === id);
    return null;
  }

  function getTarget(el) {
    return el.closest('.week-event, .month-event, .allday-event, .md-banner');
  }

  function showPopover(ev, anchor) {
    const clockIcon = '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>';
    const calIcon = '<svg viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>';

    let timeStr = '';
    if (ev.allDay) {
      timeStr = 'All day';
    } else if (ev.start) {
      timeStr = formatTime12(ev.start);
      if (ev.end) timeStr += ' \u2013 ' + formatTime12(ev.end);
    }

    let dateStr = formatDate(ev.date);
    if (ev.endDate && ev.endDate > ev.date) {
      dateStr += ' \u2013 ' + formatDate(ev.endDate);
    }

    pop.innerHTML =
      '<div class="ev-popover-title"><span class="ev-popover-dot" style="background:' + (eventDisplayColor(ev) || 'var(--accent)') + '"></span>' + esc(ev.title) + '</div>' +
      '<div class="ev-popover-row">' + calIcon + ' ' + dateStr + '</div>' +
      (timeStr ? '<div class="ev-popover-row">' + clockIcon + ' ' + timeStr + '</div>' : '') +
      (ev.isRecurring ? '<div class="ev-popover-row"><svg viewBox="0 0 24 24"><path d="M17 1l4 4-4 4"/><path d="M3 11V9a4 4 0 014-4h14"/><path d="M7 23l-4-4 4-4"/><path d="M21 13v2a4 4 0 01-4 4H3"/></svg>Recurring</div>' : '') +
      (ev.location ? '<div class="ev-popover-row"><svg viewBox="0 0 24 24"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg> ' + esc(ev.location) + '</div>' : '') +
      (ev.notes ? '<div class="ev-popover-notes">' + esc(ev.notes) + '</div>' : '');

    // Position near anchor
    const rect = anchor.getBoundingClientRect();
    pop.style.left = '0px';
    pop.style.top = '0px';
    pop.classList.add('visible');

    const popRect = pop.getBoundingClientRect();
    let left = rect.right + 8;
    let top = rect.top;

    // If overflows right, flip to left of anchor
    if (left + popRect.width > window.innerWidth - 8) {
      left = rect.left - popRect.width - 8;
    }
    // If overflows left, position at anchor left
    if (left < 8) left = 8;
    // If overflows bottom, shift up
    if (top + popRect.height > window.innerHeight - 8) {
      top = window.innerHeight - popRect.height - 8;
    }
    if (top < 8) top = 8;

    pop.style.left = left + 'px';
    pop.style.top = top + 'px';
  }

  function hidePopover() {
    clearTimeout(hoverTimer);
    hoverTimer = null;
    pop.classList.remove('visible');
  }

  // Delegate on document for all views
  document.addEventListener('mouseover', function(e) {
    const target = getTarget(e.target);
    if (!target) return;
    if (hoverTimer) clearTimeout(hoverTimer);
    hoverTimer = setTimeout(function() {
      const ev = findEvent(target);
      if (ev) showPopover(ev, target);
    }, 350);
  });

  document.addEventListener('mouseout', function(e) {
    const target = getTarget(e.target);
    if (!target) return;
    // Check if we moved to a child of the same target
    const related = e.relatedTarget;
    if (related && target.contains(related)) return;
    hidePopover();
  });

  // Hide on scroll or drag start
  document.addEventListener('dragstart', hidePopover);
  document.addEventListener('scroll', hidePopover, true);
})();

// ── Event resize interaction ───────────────────────────────────────────────────
document.addEventListener('mousemove', function(e) {
  if (!resizeEv || !resizeEl || !resizeCol) return;
  const colRect = resizeCol.getBoundingClientRect();
  const scrollTop = (document.getElementById('week-body') || {scrollTop:0}).scrollTop;
  // clientY relative to column top; col is absolutely positioned inside scrollable body,
  // so we use the col's bounding rect (viewport-relative) which already factors scroll.
  const y = Math.max(0, e.clientY - colRect.top);

  if (resizeEdge === 'bottom') {
    const newEnd = yToTime(y);
    const [sh, sm] = resizeOrigStart.split(':').map(Number);
    const [eh, em] = newEnd.split(':').map(Number);
    if (eh * 60 + em < sh * 60 + sm + 15) return;
    const newH = Math.max(22, timeToY(newEnd) - timeToY(resizeOrigStart));
    resizeEl.style.height = newH + 'px';
    const timeEl = resizeEl.querySelector('.ev-time');
    if (timeEl) timeEl.textContent = resizeOrigStart + ' – ' + newEnd;
  } else {
    const newStart = yToTime(y);
    const [sh, sm] = newStart.split(':').map(Number);
    const [eh, em] = resizeOrigEnd.split(':').map(Number);
    if (eh * 60 + em < sh * 60 + sm + 15) return;
    const newTop = timeToY(newStart);
    const newH = Math.max(22, timeToY(resizeOrigEnd) - newTop);
    resizeEl.style.top = newTop + 'px';
    resizeEl.style.height = newH + 'px';
    const timeEl = resizeEl.querySelector('.ev-time');
    if (timeEl) timeEl.textContent = newStart + ' – ' + resizeOrigEnd;
  }
});

document.addEventListener('mouseup', function(e) {
  if (!resizeEv || !resizeEl || !resizeCol) return;
  const colRect = resizeCol.getBoundingClientRect();
  const y = Math.max(0, e.clientY - colRect.top);

  let finalStart = resizeOrigStart;
  let finalEnd   = resizeOrigEnd;

  if (resizeEdge === 'bottom') {
    const newEnd = yToTime(y);
    const [sh, sm] = resizeOrigStart.split(':').map(Number);
    const [eh, em] = newEnd.split(':').map(Number);
    if (eh * 60 + em >= sh * 60 + sm + 15) finalEnd = newEnd;
  } else {
    const newStart = yToTime(y);
    const [sh, sm] = newStart.split(':').map(Number);
    const [eh, em] = resizeOrigEnd.split(':').map(Number);
    if (eh * 60 + em >= sh * 60 + sm + 15) finalStart = newStart;
  }

  if (finalStart !== resizeOrigStart || finalEnd !== resizeOrigEnd) {
    const snapEv = resizeEv, fs = finalStart, fe = finalEnd;
    if (snapEv.isRecurring) {
      showRecurActionOverlay(
        'Resize Recurring Event',
        'This is one appointment in a series. What do you want to resize?',
        () => commitRescheduleOne(snapEv, snapEv.date, fs, fe),
        () => commitResize(snapEv, fs, fe),
        'Just this one',
        'The entire series'
      );
    } else {
      commitResize(snapEv, fs, fe);
    }
  }

  // Clean up
  if (resizeEl) {
    resizeEl.classList.remove('resizing');
    resizeEl.draggable = true;
  }
  document.body.classList.remove('resizing-active');
  resizeEv = null; resizeEdge = null; resizeEl = null; resizeCol = null;
  resizeOrigStart = ''; resizeOrigEnd = '';
});

// ╔══════════════════════════════════════════════════════════════════╗
// ║  OOP LAYER                                                      ║
// ║  Classes own all logic.  CalendarApp.init() overwrites the old  ║
// ║  free-function names on `window` so every existing call-site    ║
// ║  (event listeners wired at parse-time, VBA callbacks, etc.)     ║
// ║  routes through the class methods transparently.                ║
// ╚══════════════════════════════════════════════════════════════════╝

// ── DateUtils ─────────────────────────────────────────────────────────────────
class DateUtils {
  static pad(n)        { return String(n).padStart(2, '0'); }
  static fmt(d)        { return `${d.getFullYear()}-${DateUtils.pad(d.getMonth()+1)}-${DateUtils.pad(d.getDate())}`; }
  static sameDay(a, b) { return a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }
  static addDays(d, n) { const r = new Date(d); r.setDate(r.getDate()+n); return r; }
  static startOfWeek(d)     { const r=new Date(d); r.setDate(r.getDate()-r.getDay()); return r; }
  static startOfWorkWeek(d) { const r=new Date(d),day=r.getDay(); r.setDate(r.getDate()-(day===0?6:day-1)); return r; }
  static timeToY(hhmm, slotH=56) { if(!hhmm) return 0; const[h,m]=hhmm.split(':').map(Number); return(h*60+m)*(slotH/60); }
  static yToTime(y, slotH=56)    { const mins=Math.round((y/slotH)*60); return `${DateUtils.pad(Math.floor(mins/60))}:${DateUtils.pad(mins%60)}`; }
}

// ── CalendarState ─────────────────────────────────────────────────────────────
class CalendarState {
  constructor() {
    this.view='week'; this.weekDays=7; this.cur=new Date(); this.today=new Date();
    this.activeCalendarId=null; this.primaryCalendarId=null; this.openCalendarIds=[]; this.darkMode=false;
    this.editingId=null; this.selColor=COLORS[0]; this.isAllDay=false;
    this.recurActive=false; this.editingRecurDate=null; this.editingRecurEv=null;
    this.dragEv=null; this.dragOffset=0;
    this.resizeEv=null; this.resizeEdge=null; this.resizeEl=null; this.resizeCol=null;
    this.resizeOrigStart=''; this.resizeOrigEnd='';
  }
}

// ── AppointmentStore ──────────────────────────────────────────────────────────
class AppointmentStore {
  constructor()          { this._items=[]; }
  load(arr)              { this._items=Array.isArray(arr)?[...arr]:[]; }
  all()                  { return this._items; }
  find(id)               { return this._items.find(e=>String(e.id)===String(id))??null; }
  forCalendar(calId)     { return calId?this._items.filter(e=>String(e.calendarId)===String(calId)):this._items; }
  add(ev)                { this._items.push(ev); }
  update(id,patch)       { const i=this._items.findIndex(e=>String(e.id)===String(id)); if(i!==-1)this._items[i]={...this._items[i],...patch}; }
  remove(id)             { this._items=this._items.filter(e=>String(e.id)!==String(id)); }
  removeByCalendar(calId){ this._items=this._items.filter(e=>String(e.calendarId)!==String(calId)); }
}

// ── CalendarStore ─────────────────────────────────────────────────────────────
class CalendarStore {
  constructor()  { this._items=[]; }
  load(arr)      { this._items=Array.isArray(arr)?[...arr]:[]; }
  all()          { return this._items; }
  find(id)       { return this._items.find(c=>String(c.id)===String(id))??null; }
  add(cal)       { this._items.push(cal); }
  remove(id)     { this._items=this._items.filter(c=>String(c.id)!==String(id)); }
  first()        { return this._items[0]??null; }
}

// ── ThemeManager ──────────────────────────────────────────────────────────────
class ThemeManager {
  constructor(state) { this._state=state; this._key='accessCalendarTheme'; }
  read()       { try{const v=localStorage.getItem(this._key);if(v==='dark')return true;if(v==='light')return false;}catch(_){} return false; }
  save(isDark) { try{localStorage.setItem(this._key,isDark?'dark':'light');}catch(_){} }
  getSaved()   { try{const v=localStorage.getItem(this._key);if(v==='dark'||v==='light')return v;}catch(_){} return this._state.darkMode?'dark':'light'; }
  apply(dark, persist=true) {
    this._state.darkMode=!!dark;
    applyTheme(dark, persist);
  }
  toggle() { this.apply(!this._state.darkMode, true); }
}

// ── VBABridge ─────────────────────────────────────────────────────────────────
class VBABridge {
  constructor(app) { this._app=app; this._pending=null; }

  loadData(json) {
    try {
      const d = JSON.parse(json);
      calendars      = d.calendars    || [];
      calendarGroups = d.calendarGroups || [];
      events         = d.appointments || [];
      const stored = getSavedThemePreference ? getSavedThemePreference() : null;
      if      (stored==='dark')                { darkMode=true;  applyTheme(true,false);  }
      else if (stored==='light')               { darkMode=false; applyTheme(false,false); }
      else if (typeof d.darkMode==='boolean')  { darkMode=d.darkMode; applyTheme(darkMode,false); }
      if (d.openCalendarIds && Array.isArray(d.openCalendarIds)) {
        openCalendarIds = d.openCalendarIds.map(String);
      }
      const incoming = d.primaryCalendarId ? String(d.primaryCalendarId) : (d.activeCalendarId ? String(d.activeCalendarId) : null);
      if (incoming && calendars.some(c=>String(c.id)===incoming)) { primaryCalendarId=incoming; }
      else if (!primaryCalendarId && calendars.length) { primaryCalendarId=String(calendars[0].id); }
      else if (primaryCalendarId && !calendars.some(c=>String(c.id)===String(primaryCalendarId))) {
        primaryCalendarId = calendars.length ? String(calendars[0].id) : null;
      }
      if (!openCalendarIds.length && primaryCalendarId) openCalendarIds = [primaryCalendarId];
      this._app.cals.load(calendars);
      this._app.appts.load(events);
      this._app.state.activeCalendarId = primaryCalendarId;
      this._app.state.primaryCalendarId = primaryCalendarId;
      this._app.state.openCalendarIds = openCalendarIds;
      this._app.state.darkMode         = darkMode;
      renderSidebar();
      render();
      setStatus('idle');
    } catch(e) { setStatus('error','Bad data from Access'); console.error(e); }
  }
  getPendingCommand()   { return pendingCommand!==null ? JSON.stringify(pendingCommand) : null; }
  clearPendingCommand() { pendingCommand=null; setStatus('idle'); }
  queueCommand(cmd)     { queueCommand(cmd); }
  setStatus(t,m)        { setStatus(t,m); }
}

// ── RecurrencePanel ───────────────────────────────────────────────────────────
class RecurrencePanel {
  constructor(app) { this.app = app; }

  get active()  { return recurActive; }
  set active(v) { recurActive = v; this.app.state.recurActive = v; }

  init() {
    // Build DOW checkbox row
    const row = document.getElementById('r-dow-row');
    row.innerHTML = '';
    DOW_LABELS.forEach((lbl, i) => {
      const lab = document.createElement('label');
      lab.style.cssText = 'display:flex;align-items:center;gap:3px;font-size:0.8rem;user-select:none;cursor:pointer';
      lab.innerHTML = `<input type="checkbox" data-dow="${i}" style="accent-color:var(--accent);cursor:pointer"> ${lbl}`;
      row.appendChild(lab);
    });

    // Wire recur-type radio listeners once
    document.querySelectorAll('input[name="recur-type"]').forEach(r =>
      r.addEventListener('change', () => { if (r.checked) this.showSubPanel(r.value); })
    );

    // Wire recur-toggle button
    document.getElementById('recur-toggle-btn').addEventListener('click', () => {
      this.toggleRecurrence();
    });
  }

  toggleRecurrence() {
    this.active = !this.active;
    document.getElementById('modal-right').style.display = this.active ? '' : 'none';
    document.getElementById('recur-toggle-btn').textContent = this.active ? 'Remove Recurrence' : 'Make Recurring';
    if (this.active) {
      const dateVal = document.getElementById('ev-date').value;
      if (dateVal) {
        const d = new Date(dateVal + 'T00:00:00');
        document.getElementById('r-range-start').value = dateVal;
        const ed = new Date(d); ed.setMonth(ed.getMonth() + 6);
        document.getElementById('r-end-date').value = fmt(ed);
        document.getElementById('r-month-day').value = d.getDate();
        document.getElementById('r-yearly-label').value =
          ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()] + ' ' + d.getDate();
        document.querySelectorAll('#r-dow-row input[type=checkbox]').forEach(cb =>
          cb.checked = parseInt(cb.dataset.dow) === d.getDay());
      }
      document.querySelector('input[name="recur-type"][value="weekly"]').checked = true;
      this.showSubPanel('weekly');
    }
  }

  showSubPanel(type) {
    ['daily', 'weekly', 'monthly', 'yearly'].forEach(t =>
      document.getElementById(`rp-${t}`).style.display = (t === type ? '' : 'none')
    );
  }

  getRecurData() {
    if (!this.active) {
      return {
        recurType: 'none', recurInterval: 1, recurDaysOfWeek: '',
        recurMonthlyMode: 'day', recurMonthDay: 1, recurMonthWeek: 1,
        recurMonthDOW: 0, recurEndType: 'never', recurEndDate: '', recurCount: 0
      };
    }
    const type    = document.querySelector('input[name="recur-type"]:checked')?.value ?? 'weekly';
    const endType = document.querySelector('input[name="r-end-type"]:checked')?.value ?? 'date';
    const data = {
      recurType: type, recurInterval: 1, recurDaysOfWeek: '',
      recurMonthlyMode: 'day', recurMonthDay: 1, recurMonthWeek: 1, recurMonthDOW: 0,
      recurEndType: endType,
      recurEndDate:    endType === 'date'  ? document.getElementById('r-end-date').value  : '',
      recurCount:      endType === 'count' ? parseInt(document.getElementById('r-end-count').value, 10) || 0 : 0,
      recurRangeStart: document.getElementById('r-range-start').value,
    };
    switch (type) {
      case 'daily':
        data.recurInterval = parseInt(document.getElementById('r-daily-int').value, 10) || 1;
        break;
      case 'weekly': {
        data.recurInterval = parseInt(document.getElementById('r-weekly-int').value, 10) || 1;
        const checked = [];
        document.querySelectorAll('#r-dow-row input[type=checkbox]:checked')
          .forEach(cb => checked.push(parseInt(cb.dataset.dow, 10)));
        data.recurDaysOfWeek = checked.join(',');
        break;
      }
      case 'monthly': {
        const mm = document.querySelector('input[name="r-monthly-mode"]:checked')?.value ?? 'day';
        data.recurMonthlyMode = mm;
        if (mm === 'day') {
          data.recurMonthDay = parseInt(document.getElementById('r-month-day').value, 10) || 1;
          data.recurInterval = parseInt(document.getElementById('r-month-int').value, 10) || 1;
        } else {
          data.recurMonthWeek = parseInt(document.getElementById('r-month-week').value, 10) || 1;
          data.recurMonthDOW  = parseInt(document.getElementById('r-month-dow').value,  10) || 0;
          data.recurInterval  = parseInt(document.getElementById('r-month-int2').value, 10) || 1;
        }
        break;
      }
      case 'yearly':
        data.recurInterval = parseInt(document.getElementById('r-yearly-int').value, 10) || 1;
        break;
    }
    return data;
  }

  setRecurData(ev) {
    if (!ev || !ev.recurType || ev.recurType === 'none') {
      this.active = false;
      document.getElementById('modal-right').style.display = 'none';
      document.getElementById('recur-toggle-btn').textContent = 'Make Recurring';
      return;
    }
    this.active = true;
    document.getElementById('modal-right').style.display = '';
    document.getElementById('recur-toggle-btn').textContent = 'Remove Recurrence';

    const rtype = document.querySelector(`input[name="recur-type"][value="${ev.recurType}"]`);
    if (rtype) rtype.checked = true;
    this.showSubPanel(ev.recurType);

    switch (ev.recurType) {
      case 'daily':
        document.getElementById('r-daily-int').value = ev.recurInterval || 1;
        break;
      case 'weekly': {
        document.getElementById('r-weekly-int').value = ev.recurInterval || 1;
        const dows = String(ev.recurDaysOfWeek).split(',').map(Number);
        document.querySelectorAll('#r-dow-row input[type=checkbox]')
          .forEach(cb => cb.checked = dows.includes(parseInt(cb.dataset.dow, 10)));
        break;
      }
      case 'monthly': {
        const mm  = ev.recurMonthlyMode || 'day';
        const mmR = document.querySelector(`input[name="r-monthly-mode"][value="${mm}"]`);
        if (mmR) mmR.checked = true;
        document.getElementById('r-month-day').value  = ev.recurMonthDay  || 1;
        document.getElementById('r-month-int').value  = ev.recurInterval  || 1;
        document.getElementById('r-month-int2').value = ev.recurInterval  || 1;
        document.getElementById('r-month-week').value = ev.recurMonthWeek || 1;
        document.getElementById('r-month-dow').value  = ev.recurMonthDOW  || 0;
        break;
      }
      case 'yearly':
        document.getElementById('r-yearly-int').value = ev.recurInterval || 1;
        break;
    }
    const rangeStart = ev.recurRangeStart || ev.date || '';
    document.getElementById('r-range-start').value = rangeStart;
    if (rangeStart) {
      const d = new Date(rangeStart + 'T00:00:00');
      document.getElementById('r-yearly-label').value =
        ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()] +
        ' ' + d.getDate();
    }
    const etR = document.querySelector(`input[name="r-end-type"][value="${ev.recurEndType || 'date'}"]`);
    if (etR) etR.checked = true;
    document.getElementById('r-end-date').value  = ev.recurEndDate || '';
    document.getElementById('r-end-count').value = ev.recurCount   || 10;
  }
}

// ── Toast notifications ─────────────────────────────────────────────────────────
function showToast(message, kind = 'info', timeoutMs = 3200) {
  const host = document.getElementById('toast-host');
  if (!host) { return; }
  const t = document.createElement('div');
  t.className = 'toast toast-' + (kind || 'info');
  t.textContent = message;
  host.appendChild(t);
  void t.offsetWidth;            // force reflow so the enter transition runs
  t.classList.add('show');
  const remove = () => {
    t.classList.remove('show');
    setTimeout(() => { if (t.parentNode) t.parentNode.removeChild(t); }, 250);
  };
  const timer = setTimeout(remove, timeoutMs);
  t.addEventListener('click', () => { clearTimeout(timer); remove(); });
}
window.showToast = showToast;

// ── AppointmentModal ──────────────────────────────────────────────────────────
class AppointmentModal {
  constructor(app) { this._app = app; }

  init() {
    const ov = document.getElementById('modal-overlay');
    document.getElementById('modal-cancel').onclick = () => this.close();
    ov.onclick = e => { if (e.target === ov) this.close(); };
    document.getElementById('modal-delete').onclick = () => this.handleDelete();
    document.getElementById('modal-save').onclick   = () => this.handleSave();
    const outlookBtn = document.getElementById('modal-add-outlook');
    if (outlookBtn) outlookBtn.onclick = () => this.handleAddToOutlook();
    const icsBtn = document.getElementById('modal-export-ics');
    if (icsBtn) icsBtn.onclick = () => this.handleExportIcs();
  }

  open(dateStr, startStr = '09:00', endStr = '10:00', ev = null) {
    // Refuse to edit a still-unsaved optimistic placeholder (tmp id): the real
    // record hasn't come back from VBA yet, so its id isn't stable. Editing/
    // resaving it would send a non-numeric id, which VBA would otherwise turn
    // into a silent duplicate (AddNew) instead of an update.
    if (ev && ev.isPending) {
      showToast('Still saving\u2026 try again in a moment.', 'info');
      return;
    }
    editingId        = ev ? (ev.recurMasterId != null ? ev.recurMasterId : ev.id) : null;
    editingRecurDate = ev && ev.isRecurring ? ev.date : null;
    editingRecurEv   = (ev && ev.isRecurring) ? ev : null;
    document.getElementById('modal-title').textContent = ev ? 'Edit Appointment' : 'New Appointment';
    const calSel = document.getElementById('ev-calendar');
    calSel.innerHTML = openCalendarIds.map(cid => {
      const c = calendars.find(cc => String(cc.id) === String(cid));
      return c ? `<option value="${c.id}">${c.name}</option>` : '';
    }).join('');
    calSel.value = ev ? String(ev.calendarId) : String(primaryCalendarId);
    document.getElementById('ev-title').value      = ev ? ev.title : '';
    document.getElementById('ev-date').value        = ev ? ev.date  : dateStr;
    document.getElementById('ev-end-date').value    = ev ? (ev.endDate || ev.date || dateStr) : dateStr;
    document.getElementById('ev-notes').value       = ev ? (ev.notes || '') : '';
    document.getElementById('ev-location').value    = ev ? (ev.location || '') : '';
    setAllDay(ev ? !!ev.allDay : false);
    document.getElementById('ev-start').value = ev ? ev.start : startStr;
    document.getElementById('ev-end').value   = ev ? ev.end   : endStr;
    selColor = ev ? ev.color : COLORS[0];
    this.renderColorRow();
    setRecurData(ev);
    document.getElementById('ev-reminder').value =
        ev ? (ev.reminderMinutes !== undefined && ev.reminderMinutes !== null && ev.reminderMinutes >= 0
              ? String(ev.reminderMinutes) : '-1')
           : '15';
    document.getElementById('modal-delete').style.display = ev ? '' : 'none';
    overlay.style.display = 'flex';
    setTimeout(() => document.getElementById('ev-title').focus(), 50);
  }

  close() {
    overlay.style.display = 'none';
    editingId        = null;
    editingRecurDate = null;
    editingRecurEv   = null;
  }

  getFormData() {
    return {
      title:    document.getElementById('ev-title').value.trim() || 'No title',
      date:     document.getElementById('ev-date').value,
      endDate:  document.getElementById('ev-end-date').value,
      start:    document.getElementById('ev-start').value,
      end:      document.getElementById('ev-end').value,
      notes:    document.getElementById('ev-notes').value.trim(),
      location: document.getElementById('ev-location').value.trim(),
      allDay:   isAllDay,
      color:    selColor,
      calendarId: document.getElementById('ev-calendar').value || primaryCalendarId,
      reminderMinutes: parseInt(document.getElementById('ev-reminder').value, 10),
      ...getRecurData()
    };
  }

  renderColorRow() {
    const row = document.getElementById('color-row');
    row.innerHTML = '';
    COLORS.forEach(c => {
      const s = document.createElement('div');
      s.className = 'color-swatch' + (c === selColor ? ' selected' : '');
      s.style.background = c;
      s.onclick = () => { selColor = c; this.renderColorRow(); };
      row.appendChild(s);
    });
  }

  handleDelete() {
    if (editingId == null) return;
    overlay.style.display = 'none';

    if (editingRecurDate !== null) {
      showRecurActionOverlay(
        'Delete Recurring Event',
        'Do you want to delete just this occurrence or all events in the series?',
        () => {
          if (!standalone)
            queueCommand({action:'deleteOccurrence', id: editingId, date: editingRecurDate});
          events = events.filter(e => !(String(e.recurMasterId) === String(editingId) && e.date === editingRecurDate));
          render();
        },
        () => {
          if (!standalone)
            queueCommand({action:'deleteSeries', id: editingId});
          events = events.filter(e => String(e.recurMasterId) !== String(editingId) && e.id != editingId);
          render();
        }
      );
    } else {
      if (standalone) {
        events = events.filter(e => e.id !== editingId);
      } else {
        queueCommand({action:'delete', id: editingId});
        events = events.filter(e => e.id != editingId);
      }
      render();
    }
  }

  handleSave() {
    const fd = this.getFormData();
    if (!fd.date) return;

    if (editingRecurDate !== null && editingId !== null && editingRecurEv !== null) {
      overlay.style.display = 'none';
      showRecurActionOverlay(
        'Edit Recurring Event',
        'Update only this occurrence, or every event in the series?',
        () => {
          commitRescheduleOne(
            editingRecurEv, fd.date, fd.start, fd.end,
            { title: fd.title, color: fd.color, notes: fd.notes, location: fd.location,
              allDay: fd.allDay, reminderMinutes: fd.reminderMinutes }
          );
        },
        () => { this._doSaveAll(fd); },
        'This occurrence',
        'All occurrences'
      );
      return;
    }
    this._doSaveAll(fd);
  }

  _doSaveAll(fd) {
    if (standalone) {
      if (editingId !== null) {
        const ev = events.find(e => e.id === editingId);
        if (ev) Object.assign(ev, {
          title: fd.title, date: fd.date, endDate: fd.endDate,
          start: fd.start, end: fd.end, notes: fd.notes,
          location: fd.location,
          allDay: fd.allDay, color: fd.color,
          recurType: fd.recurType, recurInterval: fd.recurInterval,
          recurDaysOfWeek: fd.recurDaysOfWeek, recurMonthlyMode: fd.recurMonthlyMode,
          recurMonthDay: fd.recurMonthDay, recurMonthWeek: fd.recurMonthWeek,
          recurMonthDOW: fd.recurMonthDOW, recurEndType: fd.recurEndType,
          recurEndDate: fd.recurEndDate, recurCount: fd.recurCount
        });
      } else {
        events.push({
          id: Date.now(), title: fd.title, date: fd.date, endDate: fd.endDate,
          start: fd.start, end: fd.end, notes: fd.notes, location: fd.location, allDay: fd.allDay,
          color: fd.color, calendarId: fd.calendarId,
          recurType: fd.recurType, recurInterval: fd.recurInterval,
          recurDaysOfWeek: fd.recurDaysOfWeek, recurMonthlyMode: fd.recurMonthlyMode,
          recurMonthDay: fd.recurMonthDay, recurMonthWeek: fd.recurMonthWeek,
          recurMonthDOW: fd.recurMonthDOW, recurEndType: fd.recurEndType,
          recurEndDate: fd.recurEndDate, recurCount: fd.recurCount
        });
      }
    } else {
      queueCommand({
        action:'save', id: editingId || 0, calendarId: fd.calendarId,
        title: fd.title, date: fd.date, endDate: fd.endDate,
        start: fd.start, end: fd.end, notes: fd.notes, location: fd.location, allDay: fd.allDay,
        color: fd.color, reminderMinutes: fd.reminderMinutes,
        recurType:        fd.recurType        || 'none',
        recurInterval:    fd.recurInterval    || 1,
        recurDaysOfWeek:  fd.recurDaysOfWeek  || '',
        recurMonthlyMode: fd.recurMonthlyMode || 'day',
        recurMonthDay:    fd.recurMonthDay    || 1,
        recurMonthWeek:   fd.recurMonthWeek   || 1,
        recurMonthDOW:    fd.recurMonthDOW    || 0,
        recurEndType:     fd.recurEndType     || 'never',
        recurEndDate:     fd.recurEndDate     || '',
        recurCount:       fd.recurCount       || 0,
        recurRangeStart:  fd.recurRangeStart  || fd.date
      });
      if (editingId !== null) {
        const hadRecurrence = events.some(e => String(e.recurMasterId) === String(editingId));
        events = events.filter(e => e.id !== editingId && String(e.recurMasterId) !== String(editingId));
        // For non-recurring edits OR converting recurring→single: keep event visible as placeholder
        if (!hadRecurrence || fd.recurType === 'none')
          events.push({id:`tmp${Date.now()}`, title: fd.title, date: fd.date,
            endDate: fd.endDate, start: fd.start, end: fd.end, notes: fd.notes,
            location: fd.location, allDay: fd.allDay, color: fd.color, calendarId: fd.calendarId,
            recurType:'none', isPending: true});
        // For recurring→still-recurring: no placeholder; VBA must rebuild all occurrences
      } else {
        events.push({id:`tmp${Date.now()}`, title: fd.title, date: fd.date,
          endDate: fd.endDate, start: fd.start, end: fd.end, notes: fd.notes,
          location: fd.location, allDay: fd.allDay, color: fd.color, calendarId: fd.calendarId,
          recurType: fd.recurType, recurInterval: fd.recurInterval,
          recurDaysOfWeek: fd.recurDaysOfWeek, isPending: true});
      }
    }
    overlay.style.display = 'none';
    render();
  }

  // ── Build the normalized event-field payload shared by save & Outlook ──────
  _eventFields(fd) {
    return {
      id: editingId || 0, calendarId: fd.calendarId,
      title: fd.title, date: fd.date, endDate: fd.endDate,
      start: fd.start, end: fd.end, notes: fd.notes, location: fd.location,
      allDay: fd.allDay, color: fd.color, reminderMinutes: fd.reminderMinutes,
      recurType:        fd.recurType        || 'none',
      recurInterval:    fd.recurInterval    || 1,
      recurDaysOfWeek:  fd.recurDaysOfWeek  || '',
      recurMonthlyMode: fd.recurMonthlyMode || 'day',
      recurMonthDay:    fd.recurMonthDay    || 1,
      recurMonthWeek:   fd.recurMonthWeek   || 1,
      recurMonthDOW:    fd.recurMonthDOW    || 0,
      recurEndType:     fd.recurEndType     || 'never',
      recurEndDate:     fd.recurEndDate     || '',
      recurCount:       fd.recurCount       || 0,
      recurRangeStart:  fd.recurRangeStart  || fd.date
    };
  }

  // ── Add the current appointment to the user's Outlook calendar (VBA bridge) ─
  handleAddToOutlook() {
    const fd = this.getFormData();
    if (!fd.date) { showToast('Pick a date first.', 'error'); return; }
    if (standalone) { showToast('Outlook is only available inside Access.', 'error'); return; }
    queueCommand(Object.assign({ action: 'addToOutlook' }, this._eventFields(fd)));
    this.close();
    showToast('Adding to Outlook\u2026', 'info');
  }

  // ── Export the current appointment as an .ics and email it via Outlook ─────
  handleExportIcs() {
    const fd = this.getFormData();
    if (!fd.date) { showToast('Pick a date first.', 'error'); return; }
    if (typeof CalendarIcs === 'undefined' || !CalendarIcs.buildBase64) {
      showToast('ICS export is unavailable.', 'error'); return;
    }
    const contentB64 = CalendarIcs.buildBase64(fd);
    if (!contentB64) { showToast('Could not build the .ics file.', 'error'); return; }
    const safe = (fd.title || 'event').replace(/[^a-z0-9_-]+/gi, '_').slice(0, 60) || 'event';
    if (standalone) {
      CalendarIcs.download(safe + '.ics', contentB64);
      return;
    }
    queueCommand({ action: 'exportIcs', filename: safe + '.ics',
                   subject: fd.title || 'Calendar event', contentB64: contentB64 });
    this.close();
    showToast('Preparing email with .ics attachment\u2026', 'info');
  }
}

// ── CalendarRenderer ──────────────────────────────────────────────────────────
class CalendarRenderer {
  constructor(app) { this._app = app; }

  init() { /* navigation & sidebar listeners are wired at parse-time above */ }

  // ── Main render dispatch ────────────────────────────────────────────────────
  render() {
    document.querySelectorAll('.view-tab').forEach(t =>
      t.classList.toggle('active', t.dataset.view === view));
    document.getElementById('week-toggle-wrap').style.display = view==='week' ? '' : 'none';
    document.getElementById('w5-btn').classList.toggle('active', weekDays===5);
    document.getElementById('w7-btn').classList.toggle('active', weekDays===7);
    if      (view==='month') this.renderMonth();
    else if (view==='week')  this.renderWeek();
    else                     this.renderDay();
    this.renderMiniCal();
    renderCalTabs();
  }

  // ── Sidebar ─────────────────────────────────────────────────────────────────
  renderSidebar() {
    const list = document.getElementById('cal-list');
    if (!calendars.length && !calendarGroups.length) {
      list.innerHTML = '<div class="cal-empty">No calendars yet.<br>Click + to add a group.</div>';
      return;
    }
    let html = '';
    calendarGroups.forEach(g => {
      const gCals = calendars.filter(c => String(c.groupId) === String(g.id));
      html += `<div class="cal-group" data-group-id="${g.id}">
        <div class="cal-group-header" data-group-id="${g.id}">
          <span class="cal-group-arrow">\u25BE</span>
          <span class="cal-group-name" title="${esc(g.name)}">${esc(g.name)}</span>
          <button class="cal-group-add-btn" data-group-id="${g.id}" title="Add calendar to group">+</button>
          <div class="cal-group-actions">
            <button class="grp-edit-btn" data-group-id="${g.id}" title="Rename group">&#9998;</button>
            <button class="grp-del-btn" data-group-id="${g.id}" title="Delete group">\u2715</button>
          </div>
        </div>
        <div class="cal-group-body">
          ${gCals.length ? gCals.map(c => calItemHtml(c)).join('') : '<div class="cal-empty" style="padding:8px 12px;font-size:0.72rem">Drop calendars here</div>'}
        </div>
      </div>`;
    });
    const ungrouped = calendars.filter(c => !c.groupId || !calendarGroups.some(g => String(g.id) === String(c.groupId)));
    if (ungrouped.length || !calendarGroups.length) {
      html += `<div class="cal-group" data-group-id="">
        <div class="cal-group-header" data-group-id="">
          <span class="cal-group-arrow">\u25BE</span>
          <span class="cal-group-name">Ungrouped</span>
        </div>
        <div class="cal-group-body">
          ${ungrouped.length ? ungrouped.map(c => calItemHtml(c)).join('') : '<div class="cal-empty" style="padding:8px 12px;font-size:0.72rem">No ungrouped calendars</div>'}
        </div>
      </div>`;
    }
    list.innerHTML = html;
    list.querySelectorAll('.cal-chk').forEach(chk => {
      chk.addEventListener('change', () => toggleCalendar(chk.value));
    });
    list.querySelectorAll('.cal-del-btn').forEach(btn => {
      btn.addEventListener('click', e => { e.stopPropagation(); deleteCalendar(btn.dataset.id); });
    });
    list.querySelectorAll('.cal-edit-btn').forEach(btn => {
      btn.addEventListener('click', e => { e.stopPropagation(); openCalEditModal(btn.dataset.id); });
    });
    list.querySelectorAll('.cal-item').forEach(item => {
      item.addEventListener('click', e => {
        if (e.target.closest('.cal-del-btn') || e.target.closest('.cal-edit-btn') || e.target.tagName === 'LABEL' || e.target.classList.contains('cal-chk')) return;
        setPrimaryCalendar(item.dataset.id);
      });
      item.addEventListener('dragstart', e => {
        e.dataTransfer.setData('text/cal-id', item.dataset.id);
        e.dataTransfer.effectAllowed = 'move';
        requestAnimationFrame(() => item.classList.add('dragging'));
      });
      item.addEventListener('dragend', () => item.classList.remove('dragging'));
    });
    list.querySelectorAll('.cal-group-header').forEach(hdr => {
      hdr.addEventListener('click', e => {
        if (e.target.closest('.cal-group-actions') || e.target.closest('.cal-group-add-btn')) return;
        const body = hdr.nextElementSibling;
        const collapsed = hdr.classList.toggle('collapsed');
        body.classList.toggle('collapsed', collapsed);
      });
    });
    list.querySelectorAll('.grp-edit-btn').forEach(btn => {
      btn.addEventListener('click', e => { e.stopPropagation(); openGroupModal(btn.dataset.groupId); });
    });
    list.querySelectorAll('.grp-del-btn').forEach(btn => {
      btn.addEventListener('click', e => { e.stopPropagation(); deleteGroup(btn.dataset.groupId); });
    });
    list.querySelectorAll('.cal-group-add-btn').forEach(btn => {
      btn.addEventListener('click', e => { e.stopPropagation(); openCalModal(btn.dataset.groupId); });
    });
    list.querySelectorAll('.cal-group').forEach(grp => {
      grp.addEventListener('dragover', e => {
        if (!e.dataTransfer.types.includes('text/cal-id')) return;
        e.preventDefault();
        grp.classList.add('drag-over');
      });
      grp.addEventListener('dragleave', e => {
        if (!grp.contains(e.relatedTarget)) grp.classList.remove('drag-over');
      });
      grp.addEventListener('drop', e => {
        e.preventDefault();
        grp.classList.remove('drag-over');
        const calId = e.dataTransfer.getData('text/cal-id');
        if (!calId) return;
        const groupId = grp.dataset.groupId || '';
        moveCalendarToGroup(calId, groupId);
      });
    });
  }

  // ── Month view ──────────────────────────────────────────────────────────────
  renderMonth() {
    const y=cur.getFullYear(), m=cur.getMonth();
    const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
    label.textContent = MONTHS[m]+' '+y;
    const first=new Date(y,m,1), start=startOfWeek(first);
    const days  = Array.from({length:42},(_,i)=>addDays(start,i));
    const evs   = calEvents();
    const mdEvs = evs.filter(e => e.endDate && e.endDate > e.date);
    const sdEvs = evs.filter(e => !e.endDate || e.endDate <= e.date);

    const weeksHtml = Array.from({length:6},(_,w) => {
      const wDays = days.slice(w*7, w*7+7);
      return '<div class="week-row" data-week="'+w+'">' +
        wDays.map(day => {
          const ds=fmt(day), isToday=sameDay(day,today), other=day.getMonth()!==m;
          const dayEvs = sdEvs.filter(e=>e.date===ds).sort((a,b)=>a.start.localeCompare(b.start));
          const visible=dayEvs.slice(0,3), more=dayEvs.length-3;
          return '<div class="month-cell'+(other?' other-month':'')+(isToday?' today':'')+'" data-date="'+ds+'">' +
            '<div class="day-num">'+day.getDate()+'</div>' +
            '<div class="md-spacer"></div>' +
            visible.map(ev=>'<div class="month-event'+(isEvSelected(ev)?' ev-selected':'')+'" style="background:'+eventDisplayColor(ev)+';--ev-accent:'+eventAccentColor(ev)+
              '" data-id="'+ev.id+'" draggable="true">'+(ev.isRecurring?'↻ ':'')+esc(ev.title)+'</div>').join('') +
            (more>0?'<div class="more-events">+'+more+' more</div>':'') +
            '</div>';
        }).join('') +
        '</div>';
    }).join('');

    calEl.innerHTML =
      '<div class="month-header">'+
      ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'].map(n=>'<div>'+n+'</div>').join('')+
      '</div><div class="month-grid" id="month-grid">'+weeksHtml+'</div>';

    // ── Multiday banners ───────────────────────────────────────────────────
    document.querySelectorAll('.week-row').forEach((wRow, w) => {
      const wDays     = days.slice(w*7, w*7+7);
      const wStartStr = fmt(wDays[0]), wEndStr = fmt(wDays[6]);
      const weekEvs   = mdEvs
        .filter(ev => ev.date <= wEndStr && ev.endDate >= wStartStr)
        .sort((a,b) => a.date.localeCompare(b.date));
      if (!weekEvs.length) return;

      const slotLastCol = [];
      const assigns = [];
      weekEvs.forEach(ev => {
        const segS = ev.date    < wStartStr ? wStartStr : ev.date;
        const segE = ev.endDate > wEndStr   ? wEndStr   : ev.endDate;
        const colS = wDays.findIndex(d => fmt(d) === segS);
        const colE = wDays.findIndex(d => fmt(d) === segE);
        if (colS < 0) return;
        const effColE = colE >= 0 ? colE : wDays.length - 1;
        let slot = slotLastCol.findIndex(last => last < colS);
        if (slot < 0) { slot = slotLastCol.length; slotLastCol.push(effColE); }
        else slotLastCol[slot] = effColE;
        assigns.push({ ev, colS, colE: effColE, slot,
          isPartial: ev.date < wStartStr, continues: ev.endDate > wEndStr });
      });

      const maxSlotPerCol = Array(7).fill(-1);
      assigns.forEach(({colS, colE, slot}) => {
        for (let c=colS; c<=colE; c++) maxSlotPerCol[c] = Math.max(maxSlotPerCol[c], slot);
      });
      wRow.querySelectorAll('.month-cell').forEach((cell, ci) => {
        const sp = cell.querySelector('.md-spacer');
        if (sp) sp.style.height = maxSlotPerCol[ci] >= 0 ? ((maxSlotPerCol[ci]+1)*22+2)+'px' : '0';
      });

      const pct = 100/7;
      assigns.forEach(({ev, colS, colE, slot, isPartial, continues}) => {
        const span   = colE - colS + 1;
        const banner = document.createElement('div');
        let   cls    = 'md-banner';
        if (isPartial) cls += ' md-cont-left';
        if (continues) cls += ' md-cont-right';
        if (isEvSelected(ev)) cls += ' ev-selected';
        banner.className       = cls;
        banner.style.background = eventDisplayColor(ev);
        banner.style.setProperty('--ev-accent', eventAccentColor(ev));
        banner.style.top       = (24 + slot*22)+'px';
        banner.style.left      = 'calc('+(colS*pct)+'% + 2px)';
        banner.style.width     = 'calc('+(span*pct)+'% - 4px)';
        banner.dataset.id      = String(ev.id);
        banner.textContent     = esc(ev.title);
        wRow.appendChild(banner);
      });
    });

    // ── Month event listeners ──────────────────────────────────────────────
    calEl.querySelectorAll('.month-cell').forEach(cell => {
      cell.addEventListener('click', e => {
        if (e.target.classList.contains('month-event') || e.target.classList.contains('more-events')) return;
        openModal(cell.dataset.date);
      });
      cell.addEventListener('dragover',  e => { e.preventDefault(); cell.style.background='var(--accent-hover)'; });
      cell.addEventListener('dragleave', () => { cell.style.background=''; });
      cell.addEventListener('drop', e => {
        e.preventDefault();
        cell.style.background = '';
        if (!dragEv) return;
        if (dragEv.isRecurring) {
          const snapEv = dragEv;
          const droppedDate = cell.dataset.date;
          showRecurActionOverlay(
            'Reschedule Recurring Event',
            'This is one appointment in a series. What do you want to reschedule?',
            () => commitRescheduleOne(snapEv, droppedDate, snapEv.start, snapEv.end),
            () => commitMove(snapEv, droppedDate, snapEv.start, snapEv.end),
            'Just this one',
            'The entire series'
          );
        } else {
          commitMove(dragEv, cell.dataset.date, dragEv.start, dragEv.end);
        }
      });
    });
    calEl.querySelectorAll('.month-event').forEach(el => {
      el.addEventListener('click', e => {
        e.stopPropagation();
        const ev=events.find(e2=>String(e2.id)===el.dataset.id);
        if(ev) openModal('','','',ev);
      });
      el.addEventListener('dragstart', e => {
        dragEv=events.find(e2=>String(e2.id)===el.dataset.id);
        if (dragEv && dragEv.isPending) { dragEv = null; e.preventDefault(); return; }
        e.dataTransfer.effectAllowed='move';
      });
      el.addEventListener('dragend', () => { dragEv=null; });
    });
    calEl.querySelectorAll('.md-banner').forEach(el => {
      el.addEventListener('click', e => {
        e.stopPropagation();
        const ev=events.find(e2=>String(e2.id)===el.dataset.id);
        if(ev) openModal('','','',ev);
      });
    });
  }

  // ── Work-hours helpers ──────────────────────────────────────────────────────
  _getActiveCalWorkHours() {
    const cal = calendars.find(c => String(c.id) === String(primaryCalendarId));
    if (!cal) return null;
    const ws = parseInt((cal.workStart || '').split(':')[0]);
    const we = parseInt((cal.workEnd   || '').split(':')[0]);
    if (isNaN(ws) || isNaN(we) || ws >= we) return null;
    return { start: ws, end: we };
  }

  _offHoursCls(h, wh) {
    return wh && (h >= wh.start && h < wh.end) ? ' off-hours' : '';
  }

  // ── Week view ───────────────────────────────────────────────────────────────
  renderWeek() {
    const sw = weekDays===5 ? startOfWorkWeek(cur) : startOfWeek(cur);
    const dates = Array.from({length:weekDays},(_,i)=>addDays(sw,i));
    const MONTHS=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    label.textContent = dates[0].getMonth()!==dates[dates.length-1].getMonth()
      ? `${MONTHS[dates[0].getMonth()]} \u2013 ${MONTHS[dates[dates.length-1].getMonth()]} ${dates[0].getFullYear()}`
      : `${MONTHS[dates[0].getMonth()]} ${dates[0].getFullYear()}`;
    const DAYS=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const wh = this._getActiveCalWorkHours();
    calEl.innerHTML =
      `<div class="week-container">` +
      `<div class="week-head"><div class="week-head-time"></div>${dates.map(d=>
        `<div class="week-head-day"><div class="wday">${DAYS[d.getDay()]}</div>` +
        `<div class="wdate${sameDay(d,today)?' today':''}">${d.getDate()}</div></div>`).join('')}</div>` +
      `<div class="allday-row"><div class="allday-gutter">All&#8209;day</div>${dates.map(d=>
        `<div class="allday-col" data-allday="${fmt(d)}"></div>`).join('')}</div>` +
      `<div class="week-body" id="week-body">` +
      `<div class="time-col">${Array.from({length:24},(_,h)=>`<div class="time-slot${this._offHoursCls(h,wh)}">${h===0?'':h<12?h+'am':h===12?'12pm':(h-12)+'pm'}</div>`).join('')}</div>` +
      `<div class="days-col" id="days-col">${dates.map(d=>
        `<div class="day-col" data-date="${fmt(d)}">${Array.from({length:24},(_,h)=>`<div class="hour-row${this._offHoursCls(h,wh)}" data-hour="${h}"></div>`).join('')}</div>`).join('')}</div>` +
      `</div></div>`;
    this.renderWeekEvents(dates);
    this.renderAlldayEvents(dates);
    this._app.drag.setupAlldayInteraction();
    this._app.drag.setupWeekInteraction();
    this._scrollToNow();
  }

  // ── Day view ────────────────────────────────────────────────────────────────
  renderDay() {
    const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
    const DAYS=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    label.textContent = `${DAYS[cur.getDay()]}, ${MONTHS[cur.getMonth()]} ${cur.getDate()}, ${cur.getFullYear()}`;
    const wh = this._getActiveCalWorkHours();
    calEl.innerHTML =
      `<div class="week-container">` +
      `<div class="week-head"><div class="week-head-time"></div>` +
      `<div class="week-head-day"><div class="wday">${DAYS[cur.getDay()].slice(0,3).toUpperCase()}</div>` +
      `<div class="wdate${sameDay(cur,today)?' today':''}">${cur.getDate()}</div></div></div>` +
      `<div class="allday-row"><div class="allday-gutter">All&#8209;day</div>` +
      `<div class="allday-col" data-allday="${fmt(cur)}"></div></div>` +
      `<div class="week-body" id="week-body">` +
      `<div class="time-col">${Array.from({length:24},(_,h)=>`<div class="time-slot${this._offHoursCls(h,wh)}">${h===0?'':h<12?h+'am':h===12?'12pm':(h-12)+'pm'}</div>`).join('')}</div>` +
      `<div class="days-col" id="days-col">` +
      `<div class="day-col" data-date="${fmt(cur)}">${Array.from({length:24},(_,h)=>`<div class="hour-row${this._offHoursCls(h,wh)}" data-hour="${h}"></div>`).join('')}</div>` +
      `</div></div></div>`;
    this.renderWeekEvents([new Date(cur)]);
    this.renderAlldayEvents([new Date(cur)]);
    this._app.drag.setupAlldayInteraction();
    this._app.drag.setupWeekInteraction();
    this._scrollToNow();
  }

  // ── Timed-event rendering (week & day) ──────────────────────────────────────
  renderWeekEvents(dates) {
    const daysCol = document.getElementById('days-col');
    if (!daysCol) return;
    const nowCol = daysCol.querySelector(`[data-date="${fmt(today)}"]`);
    if (nowCol) {
      const now=new Date(), line=document.createElement('div');
      line.className='now-line';
      line.style.top=(now.getHours()*60+now.getMinutes())*(56/60)+'px';
      nowCol.appendChild(line);
    }
    const evs = calEvents();
    dates.forEach(d => {
      const ds=fmt(d), col=daysCol.querySelector(`[data-date="${ds}"]`);
      if (!col) return;
      const dayEvs = evs.filter(e=>e.date===ds && !e.allDay && (!e.endDate||e.endDate<=e.date));
      const laid = layoutOverlappingEvents(dayEvs);
      laid.forEach(({ev, top, bot, col: evCol, totalCols}) => {
        const h=Math.max(bot-top,22);
        const pct = 100 / totalCols;
        const leftPct = evCol * pct;
        const el=document.createElement('div');
        el.className='week-event'+(isEvSelected(ev)?' ev-selected':'');
        el.draggable=true;
        el.style.cssText=`top:${top}px;height:${h}px;background:${eventDisplayColor(ev)};--ev-accent:${eventAccentColor(ev)};left:calc(${leftPct}% + 1px);width:calc(${pct}% - 2px)`;
        el.dataset.id=String(ev.id);
        el.innerHTML=`<div class="ev-title">${ev.isRecurring?'↻ ':''}${esc(ev.title)}</div><div class="ev-time">${ev.start}${ev.end?' \u2013 '+ev.end:''}</div>`;
        el.addEventListener('click', e => {
          e.stopPropagation();
          const found=events.find(e2=>String(e2.id)===el.dataset.id);
          if(found) openModal('','','',found);
        });
        el.addEventListener('dragstart', e => {
          if (resizeEv) { e.preventDefault(); return; }
          dragEv=events.find(e2=>String(e2.id)===el.dataset.id);
          if (dragEv && dragEv.isPending) { dragEv = null; e.preventDefault(); return; }
          dragOffset=e.clientY-el.getBoundingClientRect().top;
          e.dataTransfer.effectAllowed='move';
        });
        el.addEventListener('dragend', () => { dragEv=null; });

        // Resize handles
        const topHandle = document.createElement('div');
        topHandle.className = 'ev-resize-top';
        topHandle.title = 'Drag to change start time';
        const botHandle = document.createElement('div');
        botHandle.className = 'ev-resize-bottom';
        botHandle.title = 'Drag to change end time';

        function startResize(edge, e) {
          e.stopPropagation();
          e.preventDefault();
          const found = events.find(e2 => String(e2.id) === el.dataset.id);
          if (!found) return;
          if (found.isPending) return;
          resizeEv = found;
          resizeEdge = edge;
          resizeEl = el;
          resizeCol = col;
          resizeOrigStart = found.start;
          resizeOrigEnd = found.end || found.start;
          el.classList.add('resizing');
          document.body.classList.add('resizing-active');
          el.draggable = false;
        }
        topHandle.addEventListener('mousedown', e => startResize('top', e));
        botHandle.addEventListener('mousedown', e => startResize('bottom', e));
        el.appendChild(topHandle);
        el.appendChild(botHandle);
        col.appendChild(el);
      });
    });
  }

  // ── All-day event strip (week & day) ────────────────────────────────────────
  renderAlldayEvents(dates) {
    const evs         = calEvents();
    const visStartStr = fmt(dates[0]);
    const visEndStr   = fmt(dates[dates.length-1]);

    const singleAlldayEvs = evs.filter(e => e.allDay && (!e.endDate || e.endDate <= e.date));
    dates.forEach(d => {
      const ds  = fmt(d);
      const col = document.querySelector(`.allday-col[data-allday="${ds}"]`);
      if (!col) return;
      singleAlldayEvs.filter(e => e.date === ds).forEach(ev => {
        const el = document.createElement('div');
        el.className  = 'allday-event'+(isEvSelected(ev)?' ev-selected':'');
        el.style.background = eventDisplayColor(ev);
        el.style.setProperty('--ev-accent', eventAccentColor(ev));
        el.dataset.id = String(ev.id);
        el.textContent = ev.title;
        el.draggable = true;
        el.addEventListener('click', e => {
          e.stopPropagation();
          const found = events.find(e2 => String(e2.id) === el.dataset.id);
          if (found) openModal('', '', '', found);
        });
        el.addEventListener('dragstart', e => {
          dragEv = events.find(e2 => String(e2.id) === el.dataset.id);
          if (dragEv && dragEv.isPending) { dragEv = null; e.preventDefault(); return; }
          e.dataTransfer.effectAllowed = 'move';
          requestAnimationFrame(() => el.classList.add('dragging'));
        });
        el.addEventListener('dragend', () => {
          el.classList.remove('dragging');
          dragEv = null;
        });
        col.appendChild(el);
      });
    });

    const mdEvs = evs.filter(e => e.endDate && e.endDate > e.date);
    mdEvs.forEach(ev => {
      if (ev.date > visEndStr || ev.endDate < visStartStr) return;
      const segStartStr = ev.date    < visStartStr ? visStartStr : ev.date;
      const segEndStr   = ev.endDate > visEndStr   ? visEndStr   : ev.endDate;
      const colStartIdx = dates.findIndex(d => fmt(d) === segStartStr);
      if (colStartIdx < 0) return;
      let colEndIdx = dates.findIndex(d => fmt(d) === segEndStr);
      if (colEndIdx < 0) colEndIdx = dates.length - 1;
      const spanCount = colEndIdx - colStartIdx + 1;
      const allCols   = document.querySelectorAll('.allday-col');
      const startCol  = allCols[colStartIdx];
      if (!startCol) return;

      const el = document.createElement('div');
      el.className       = 'allday-event md-allday-span'+(isEvSelected(ev)?' ev-selected':'');
      el.style.background = eventDisplayColor(ev);
      el.style.setProperty('--ev-accent', eventAccentColor(ev));
      el.style.width     = 'calc('+spanCount+' * 100% + '+(spanCount-1)*2+'px)';
      el.dataset.id      = String(ev.id);
      el.textContent     = (ev.date < visStartStr ? '\u2039 ' : '') +
                            ev.title +
                           (ev.endDate > visEndStr ? ' \u203A' : '');
      el.addEventListener('click', e => {
        e.stopPropagation();
        const found = events.find(e2 => String(e2.id) === el.dataset.id);
        if (found) openModal('', '', '', found);
      });
      startCol.appendChild(el);
    });
  }

  _scrollToNow() {
    const body = document.getElementById('week-body');
    if (!body) return;
    setTimeout(() => { body.scrollTop = Math.max(0, (new Date().getHours()-1)*56); }, 0);
  }

  // ── Mini-month calendar ─────────────────────────────────────────────────────
  renderMiniCal() {
    const titleEl = document.getElementById('mini-cal-title');
    const grid    = document.getElementById('mini-cal-grid');
    if (!titleEl || !grid) return;

    titleEl.textContent = MINI_MO[miniM] + ' ' + miniY;

    let hlStart = null, hlEnd = null;
    if (view === 'week') {
      hlStart = weekDays === 5 ? startOfWorkWeek(cur) : startOfWeek(cur);
      hlEnd   = addDays(hlStart, weekDays - 1);
    }

    const first = new Date(miniY, miniM, 1);
    const start = startOfWeek(first);
    const days  = Array.from({length:42}, (_, i) => addDays(start, i));

    grid.innerHTML = '';
    grid.appendChild(document.createElement('div'));
    MINI_DW.forEach(d => {
      const h = document.createElement('div');
      h.className   = 'mini-dow';
      h.textContent = d;
      grid.appendChild(h);
    });

    for (let w = 0; w < 6; w++) {
      const weekStart = new Date(days[w * 7]);
      const wb = document.createElement('div');
      wb.className = 'mini-week-btn';
      wb.title     = 'Show this week';
      wb.innerHTML = '&#8250;';
      wb.addEventListener('click', () => {
        cur = new Date(weekStart);
        view = 'week';
        render();
      });
      grid.appendChild(wb);

      for (let d = 0; d < 7; d++) {
        const day  = days[w * 7 + d];
        const cell = document.createElement('div');
        cell.className   = 'mini-day';
        cell.textContent = day.getDate();
        if (day.getMonth() !== miniM) cell.classList.add('mc-other');
        if (sameDay(day, today))       cell.classList.add('mc-today');
        else if (view === 'day' && sameDay(day, cur))
          cell.classList.add('mc-selected');
        else if (view === 'week' && hlStart &&
                 day.getTime() >= hlStart.getTime() &&
                 day.getTime() <= hlEnd.getTime())
          cell.classList.add('mc-in-week');
        cell.addEventListener('click', () => {
          cur  = new Date(day);
          view = 'day';
          render();
        });
        grid.appendChild(cell);
      }
    }
  }
}

// ── DragDropManager ───────────────────────────────────────────────────────────
class DragDropManager {
  constructor(app) { this._app = app; }

  init() {
    // Resize listeners are wired at parse-time above and still reference
    // the globals (resizeEv, resizeEl, etc.).  They call commitResize /
    // commitRescheduleOne which are overridden in CalendarApp.init().
  }

  commitMove(ev, newDate, newStart, newEnd) {
    let newEndDate = '';
    if (ev.endDate && ev.endDate > ev.date) {
      const durMs = new Date(ev.endDate+'T00:00:00') - new Date(ev.date+'T00:00:00');
      newEndDate  = fmt(new Date(new Date(newDate+'T00:00:00').getTime() + durMs));
    }
    if (standalone) {
      ev.date=newDate; ev.endDate=newEndDate; ev.start=newStart; ev.end=newEnd; render();
    } else {
      queueCommand({action:'move',id:ev.id,date:newDate,endDate:newEndDate,start:newStart,end:newEnd});
      ev.date=newDate; ev.endDate=newEndDate; ev.start=newStart; ev.end=newEnd; render();
    }
  }

  commitResize(ev, newStart, newEnd) {
    if (standalone) {
      ev.start = newStart; ev.end = newEnd; render();
    } else {
      const masterId = ev.recurMasterId != null ? String(ev.recurMasterId) : '';
      queueCommand({action:'resize', id:String(ev.id), masterId:masterId, start:newStart, end:newEnd});
      ev.start = newStart; ev.end = newEnd; render();
    }
  }

  commitRescheduleOne(ev, newDate, newStart, newEnd, overrides = {}) {
    const masterId      = ev.recurMasterId != null ? ev.recurMasterId : ev.id;
    const title         = overrides.title         !== undefined ? overrides.title         : ev.title;
    const color         = overrides.color         !== undefined ? overrides.color         : (ev.color  || '#4f46e5');
    const notes         = overrides.notes         !== undefined ? overrides.notes         : (ev.notes  || '');
    const allDay        = overrides.allDay        !== undefined ? overrides.allDay        : (ev.allDay || false);
    const reminderMins  = overrides.reminderMinutes !== undefined ? overrides.reminderMinutes
                          : (ev.reminderMinutes !== undefined ? ev.reminderMinutes : -1);
    if (standalone) {
      events = events.filter(e => !(String(e.recurMasterId) === String(masterId) && e.date === ev.date));
      events.push({
        id: Date.now(), calendarId: ev.calendarId,
        title, date: newDate, endDate: '', start: newStart, end: newEnd,
        allDay, color, notes, recurType: 'none', isRecurring: false,
        reminderMinutes: reminderMins
      });
      render();
    } else {
      queueCommand({
        action: 'rescheduleOccurrence',
        masterId: String(masterId),
        originalDate: ev.date,
        newDate, start: newStart, end: newEnd,
        title, calendarId: String(ev.calendarId),
        color, notes,
        allDay: allDay ? true : false,
        reminderMinutes: reminderMins
      });
      events = events.filter(e => !(String(e.recurMasterId) === String(masterId) && e.date === ev.date));
      render();
    }
  }

  setupAlldayInteraction() {
    document.querySelectorAll('.allday-col').forEach(col => {
      col.addEventListener('dragover', e => {
        if (!dragEv || !dragEv.allDay) return;
        e.preventDefault();
        col.classList.add('drag-over');
      });
      col.addEventListener('dragleave', () => col.classList.remove('drag-over'));
      col.addEventListener('drop', e => {
        e.preventDefault();
        col.classList.remove('drag-over');
        if (!dragEv || !dragEv.allDay) return;
        const newDate = col.dataset.allday;
        if (newDate === dragEv.date) return;
        commitMove(dragEv, newDate, dragEv.start, dragEv.end);
      });
    });
  }

  setupWeekInteraction() {
    const daysCol = document.getElementById('days-col');
    if (!daysCol) return;
    daysCol.querySelectorAll('.day-col').forEach(col => {
      col.addEventListener('click', e => {
        if (e.target.closest('.week-event')) return;
        const rect=col.getBoundingClientRect(), y=e.clientY-rect.top;
        const h=Math.floor(y/56), m=y%56>=28?30:0;
        openModal(col.dataset.date, pad(h)+':'+pad(m), pad(Math.min(h+1,23))+':'+pad(m));
      });
      col.addEventListener('dragover', e => {
        if (!dragEv || dragEv.allDay) return;
        e.preventDefault(); col.style.background='var(--drag-over)';
      });
      col.addEventListener('dragleave', () => { col.style.background=''; });
      col.addEventListener('drop', e => {
        e.preventDefault(); col.style.background='';
        if (!dragEv || dragEv.allDay) return;
        const rect=col.getBoundingClientRect(), y=Math.max(0,e.clientY-rect.top-dragOffset);
        const newStart=yToTime(y);
        const [sh,sm]=dragEv.start.split(':').map(Number);
        const [eh,em]=(dragEv.end||dragEv.start).split(':').map(Number);
        const dur=(eh*60+em)-(sh*60+sm);
        const [nh,nm]=newStart.split(':').map(Number);
        const endMins=nh*60+nm+dur;
        const newEnd=pad(Math.floor(endMins/60))+':'+pad(endMins%60);
        const droppedDate=col.dataset.date;
        if (dragEv.isRecurring) {
          const snapEv = dragEv;
          showRecurActionOverlay(
            'Reschedule Recurring Event',
            'This is one appointment in a series. What do you want to reschedule?',
            () => commitRescheduleOne(snapEv, droppedDate, newStart, newEnd),
            () => commitMove(snapEv, droppedDate, newStart, newEnd),
            'Just this one',
            'The entire series'
          );
        } else {
          commitMove(dragEv, droppedDate, newStart, newEnd);
        }
      });
    });
  }
}

// ── CalendarApp ───────────────────────────────────────────────────────────────
class CalendarApp {
  constructor() {
    this.state    = new CalendarState();
    this.appts    = new AppointmentStore();
    this.cals     = new CalendarStore();
    this.theme    = new ThemeManager(this.state);
    this.bridge   = new VBABridge(this);
    this.recur    = new RecurrencePanel(this);
    this.modal    = new AppointmentModal(this);
    this.renderer = new CalendarRenderer(this);
    this.drag     = new DragDropManager(this);
  }

  init() {
    // Initialise sub-components
    this.recur.init();
    this.modal.init();
    this.renderer.init();
    this.drag.init();

    // ── Override free functions ───────────────────────────────────────────────
    // After this point every call to the old function name transparently
    // routes through the class method instead.
    window.render              = () => this.renderer.render();
    window.renderSidebar       = () => this.renderer.renderSidebar();
    window.renderMiniCal       = () => this.renderer.renderMiniCal();
    window.openModal           = (d,s,e,ev) => this.modal.open(d,s,e,ev);
    window.renderColorRow      = () => this.modal.renderColorRow();
    window.getRecurData        = () => this.recur.getRecurData();
    window.setRecurData        = (ev) => this.recur.setRecurData(ev);
    window.showRecurSubPanel   = (t) => this.recur.showSubPanel(t);
    window.commitMove          = (...a) => this.drag.commitMove(...a);
    window.commitResize        = (...a) => this.drag.commitResize(...a);
    window.commitRescheduleOne = (...a) => this.drag.commitRescheduleOne(...a);
    window.getActiveCalWorkHours = () => this.renderer._getActiveCalWorkHours();
    window.offHoursCls           = (h,wh) => this.renderer._offHoursCls(h,wh);
    window.setupWeekInteraction  = () => this.drag.setupWeekInteraction();
    window.setupAlldayInteraction = () => this.drag.setupAlldayInteraction();

    // Theme & initial render
    this.theme.apply(this.theme.read(), false);
    this.renderer.renderSidebar();
    this.renderer.render();
  }

  onDataLoaded(data) { this.bridge.loadData(JSON.stringify(data)); }
}

// ── Bootstrap ─────────────────────────────────────────────────────────────────
const app = new CalendarApp();

// Override window globals — VBA calls these by name (must not rename)
window.loadData            = json  => app.bridge.loadData(json);
window.getPendingCommand   = ()    => app.bridge.getPendingCommand();
window.clearPendingCommand = ()    => app.bridge.clearPendingCommand();

// Kick off
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => app.init());
} else {
  app.init();
}

