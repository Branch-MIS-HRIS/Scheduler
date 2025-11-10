document.addEventListener('DOMContentLoaded', function () {

  /* ===========================
     SMALL UTILS
  =========================== */
  function shadeColor(color, percent) {
    const num = parseInt(color.replace("#",""),16);
    const amt = Math.round(2.55 * percent);
    const R = (num >> 16) + amt;
    const G = ((num >> 8) & 0x00FF) + amt;
    const B = (num & 0x0000FF) + amt;
    return "#" + (
      0x1000000 +
      (R<255? (R<1?0:R) :255) * 0x10000 +
      (G<255? (G<1?0:G) :255) * 0x100 +
      (B<255? (B<1?0:B) :255)
    ).toString(16).slice(1);
  }

  function escapeHtml(str) {
    if (str == null) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // Pointer → date helper (robust to transforms/zoom)
function getDateUnderPointer() {
  const el = document.elementFromPoint(lastMouseX, lastMouseY);
  const cell = el && el.closest ? el.closest('.fc-daygrid-day, .fc-timegrid-slot') : null;
  return cell ? cell.getAttribute('data-date') : null;
}

  let multiSelectModifierActive = false;
  const hasMultiSelectModifier = evt => {
    if (evt && (typeof evt.ctrlKey === 'boolean' || typeof evt.metaKey === 'boolean')) {
      const active = !!(evt.ctrlKey || evt.metaKey);
      multiSelectModifierActive = active;
      return active;
    }
    return multiSelectModifierActive;
  };

  /* ===========================
     STATE & CONFIG
  =========================== */
  let employees = {};
  let employeeColors = JSON.parse(localStorage.getItem('employeeColors')) || {};
  const BASE_EMPLOYEE_COLORS = [
    '#3b82f6','#10b981','#f59e0b','#8b5cf6','#ec4899','#14b8a6',
    '#ef4444','#22c55e','#eab308','#0ea5e9','#6366f1','#84cc16',
    '#d946ef','#0d9488','#fb923c','#a855f7','#475569','#f97316',
    '#64748b','#60a5fa','#65a30d','#f43f5e','#059669','#7c3aed',
    '#e11d48','#9333ea','#2563eb','#9ca3af','#15803d'
  ];

  function ensureEmployeeColor(empNo) {
    if (!empNo) return null;
    if (!employeeColors) employeeColors = {};
    let color = employeeColors[empNo];
    if (!color) {
      const usedColors = Object.values(employeeColors);
      const availableColors = BASE_EMPLOYEE_COLORS.filter(c => !usedColors.includes(c));
      if (availableColors.length) {
        color = availableColors[0];
      } else {
        const key = String(empNo);
        let hash = 0;
        for (let i = 0; i < key.length; i++) { hash = ((hash << 5) - hash) + key.charCodeAt(i); hash |= 0; }
        color = BASE_EMPLOYEE_COLORS[Math.abs(hash) % BASE_EMPLOYEE_COLORS.length];
      }
      employeeColors[empNo] = color;
      try { localStorage.setItem('employeeColors', JSON.stringify(employeeColors)); } catch (e) {}
    }
    return color;
  }

  function pruneEmployeeColors(activeEmpNos = []) {
    if (!employeeColors) employeeColors = {};
    const activeSet = new Set(activeEmpNos.map(v => String(v)));
    Object.keys(employeeColors).forEach(key => { if (!activeSet.has(String(key))) delete employeeColors[key]; });
  }

  function escapeSelector(value) {
    if (typeof value === 'undefined' || value === null) return '';
    const str = String(value);
    if (typeof CSS !== 'undefined' && CSS.escape) return CSS.escape(str);
    return str.replace(/"/g, '\\"');
  }

  function findExistingShiftForEmployee(empNo, dateStr, options = {}) {
    if (!calendar || typeof calendar.getEvents !== 'function') return null;
    if (empNo == null || !dateStr) return null;
    const { exclude = [] } = options;
    const excludeObjects = new Set();
    const excludeIds = new Set();

    exclude.forEach(item => {
      if (!item) return;
      if (typeof item === 'string' || typeof item === 'number') {
        excludeIds.add(String(item));
      } else {
        excludeObjects.add(item);
        const extId = item.extendedProps && item.extendedProps.id != null ? item.extendedProps.id : null;
        const plainId = item.id != null ? item.id : null;
        if (extId != null) excludeIds.add(String(extId));
        if (plainId != null) excludeIds.add(String(plainId));
      }
    });

    const normalizedEmp = String(empNo);
    return calendar.getEvents().find(ev => {
      if (!ev) return false;
      if (excludeObjects.has(ev)) return false;
      const eventId = ev.extendedProps?.id ?? ev.id;
      if (eventId != null && excludeIds.has(String(eventId))) return false;
      return String(ev.extendedProps?.empNo) === normalizedEmp && ev.startStr === dateStr;
    }) || null;
  }

  function rgbToHsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h, s, l = (max + min) / 2;
    if (max === min) { h = s = 0; }
    else {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      switch (max) { case r: h = (g - b) / d + (g < b ? 6 : 0); break;
        case g: h = (b - r) / d + 2; break; case b: h = (r - g) / d + 4; break; }
      h /= 6;
    }
    return [Math.round(h * 360), s, l];
  }

  function getGradientFromBaseColor(hex, type = 'work') {
    if (!hex || hex[0] !== '#') return '';
    const rgb = parseInt(hex.slice(1), 16);
    const r = (rgb >> 16) & 255, g = (rgb >> 8) & 255, b = rgb & 255;
    const [h, s, l] = rgbToHsl(r, g, b);
    if (type === 'work') {
      return `linear-gradient(135deg, hsl(${h}, ${Math.round(s*100)}%, ${Math.max(25, Math.round(l*100)-10)}%), hsl(${h}, ${Math.round(s*100)}%, ${Math.min(70, Math.round(l*100)+10)}%))`;
    } else {
      return `linear-gradient(135deg, hsl(${h}, ${Math.round(s*100)}%, ${Math.min(90, Math.round(l*100)+20)}%), hsl(${h}, ${Math.round(s*100)}%, ${Math.min(95, Math.round(l*100)+25)}%))`;
    }
  }

  /* ===========================
     DOM ELEMENTS / MODALS
  =========================== */
  const calendarEl = document.getElementById('calendar');
  let employeeTableBody = document.getElementById('employee-table-body');
  const draggableCardsContainer = document.getElementById('draggable-cards-container');
  const draggablePlaceholder = document.getElementById('draggable-placeholder');
  const conflictTableBody = document.getElementById('conflict-table-body');
  const conflictsPlaceholder = document.getElementById('conflicts-placeholder');

  const shiftModal = document.getElementById('shift-modal');
  const exportModal = document.getElementById('export-modal');
  const deleteModal = document.getElementById('delete-modal');

  const shiftPresetSelect = document.getElementById('shift-preset');
  const shiftCustomInput = document.getElementById('shift-custom');
  const saveShiftBtn = document.getElementById('save-shift-btn');
  const confirmExportBtn = document.getElementById('confirm-export-btn');
  const confirmDeleteBtn = document.getElementById('confirm-delete-btn');
  const deleteModalSummary = document.getElementById('delete-modal-summary');

  // Stats
  const statsWork = document.getElementById('stats-work');
  const statsRest = document.getElementById('stats-rest');
  const statsConflicts = document.getElementById('stats-conflicts');
  const statsEmployees = document.getElementById('stats-employees');

  // Shift code presets/times
  const shiftTimes = { /* (your big shiftTimes map preserved) */ 
    "RBG-001":"9:30AM - 6:30PM","RBG-002":"10:00AM - 7:00PM","RBG-003":"11:00AM - 8:00PM",
    "RBG-004":"12:00PM - 9:00PM","RBG-005":"1:00PM - 10:00PM","RBG-006":"2:00PM - 11:00PM",
    "RBG-007":"12:30PM - 9:30PM","RBG-008":"1:30PM - 10:30PM","RBG-009":"11:30AM - 8:30PM",
    "RBG-010":"10:30AM - 7:30PM","RBG-011":"9:00AM - 6:00PM","RBG-012":"8:30AM - 5:30PM",
    "RBG-013":"3:00PM - 12:00AM","RBG-014":"3:30PM - 12:30AM","RBG-015":"2:30PM - 11:30PM",
    "RBG-016":"8:30AM - 5:30PM","RBG-017":"6:00AM - 3:00PM","RBG-018":"7:00AM - 4:00PM",
    "RBG-019":"8:00AM - 5:00PM","RBG-020":"7:30AM - 4:30PM","RBG-021":"7:00AM - 5:00PM",
    "RBG-022":"11:00AM - 7:00PM","RBG-023":"4:00PM - 1:00AM","RBG-024":"8:00PM - 5:00AM",
    "RBG-025":"9:00PM - 6:00AM","RBG-026":"6:30AM - 3:30PM","RBG-027":"6:00PM - 3:00AM",
    "RBG-028":"6:30PM - 3:30AM","RBG-029":"5:00PM - 2:00AM","RBG-030":"7:00PM - 4:00AM",
    "RBG-031":"7:00AM - 4:00PM","RBT-032":"6:00AM - 4:00PM","RBT-033":"7:00AM - 5:00PM",
    "RBT-034":"8:00AM - 6:00PM","RBT-035":"9:00AM - 7:00PM","RBT-036":"9:30AM - 7:30PM",
    "RBT-037":"9:45AM - 7:45PM","RBT-038":"10:00AM - 8:00PM","RBT-039":"11:00AM - 9:00PM",
    "RBT-040":"12:00PM - 10:00PM","RBT-041":"1:00PM - 11:00PM","RBT-042":"2:00PM - 12:00AM",
    "RBT-043":"9:00PM - 7:00AM","RBT-044":"10:00PM - 8:00AM","RBT-045":"8:30AM - 6:30PM",
    "RBT-046":"11:30AM - 9:30PM","RBT-047":"10:30AM - 8:30PM","RBT-048":"11:30PM - 9:30PM",
    "RBT-049":"12:30PM - 10:30PM","RBT-050":"7:30AM - 5:30PM","RBT-051":"6:00PM - 4:00AM",
    "RBT-052":"5:00PM - 3:00AM","RBT-053":"6:30AM - 4:30PM","RBT-054":"3:00PM - 1:00AM",
    "AASP-001":"9:30AM - 6:30PM","AASP-002":"10:00AM - 7:00PM","AASP-003":"10:30AM - 7:30PM",
    "AASP-004":"11:00AM - 8:00PM","AASP-005":"11:30AM - 8:30PM","AASP-006":"9:00AM - 7:30PM",
    "AASP-007":"10:00AM - 7:00PM","AASP-008":"9:30AM - 6:30PM","AASP-009":"7:30AM - 4:30PM",
    "AASP-010":"12:30PM - 9:30PM","AASP-011":"8:30AM - 5:30PM","AASP-012":"9:30AM - 6:30PM",
    "AASP-013":"8:00AM - 6:30PM","AASP-014":"9:00AM - 6:00PM","AASP-015":"8:00AM - 5:00PM",
    "AASP-016":"6:00AM - 3:00PM","AASP-017":"9:00AM - 6:00PM","AASP-018":"6:00AM - 3:00PM",
    "AASP-019":"12:00PM - 9:00PM","AASP-020":"6:00AM - 4:30PM","AASP-021":"8:00AM - 5:00PM",
    "AASP-021 (10-5)":"10:00AM - 5:00PM","AASP-022":"8:30AM - 7:00PM","AASP-023":"4:30PM - 12:30AM",
    "AASP-024":"9:00AM - 7:00PM","AASP-025":"6:00AM - 3:00PM","AASP-026":"8:00PM - 5:00AM",
    "AASP-027":"7:00PM - 4:00AM","AASP-028":"1:00PM - 10:00PM","AASP-029":"8:30AM - 5:30PM",
    "AASP-030":"7:00AM - 4:00PM","AASP-031":"10:00AM - 7:00PM","AASP-032":"6:00AM - 3:00PM",
    "AASP-033":"2:00PM - 11:00PM","AASP-034":"9:30AM - 7:30PM","AASP-035":"10:00AM - 8:00PM",
    "AASP-036":"10:30AM - 8:30PM","AASP-037":"11:00AM - 9:00PM","AASP-038":"11:30AM - 9:30PM",
    "AASP-039":"12:00PM - 10:00PM","AASP-040":"8:00AM - 6:00PM","AASP-041":"9:00AM - 7:00PM",
    "AASP-042":"7:30AM - 5:30PM","AASP-043":"8:30AM - 6:30PM","AASP-044":"6:00AM - 4:00PM",
    "WHSE-001":"8:00AM - 5:00PM","WHSE-002":"9:00AM - 6:00PM","WHSE-003":"10:00AM - 7:00PM",
    "WHSE-004":"12:00PM - 9:00PM","WHSE-005":"8:00AM - 12:00PM","WHSE-006":"12:00PM - 4:00PM",
    "WHSE-007":"4:00PM - 1:00AM","WHSE-008":"7:00AM - 4:00PM","WHSE-009":"7:00AM - 11:00AM",
    "WHSE-010":"5:00AM - 2:00PM","WHSE-011":"10:00AM - 2:00PM","WHSE-012":"9:00AM - 1:00PM"
  };

  function formatShiftTime(raw) {
    if (!raw) return '';
    const t = String(raw).trim();
    if (!t) return '';
    const norm = token => {
      const m = token.trim().match(/^(\d{1,2}:\d{2})\s*([AaPp][Mm])$/);
      if (!m) return token.trim().replace(/\s+/g, ' ');
      return `${m[1]} ${m[2].toUpperCase()}`;
    };
    const r = t.match(/^(.*?)\s*[-–]\s*(.*?)$/);
    if (!r) return norm(t);
    return `${norm(r[1])} – ${norm(r[2])}`;
  }

  function buildEventTooltipContent(event) {
    if (!event) return '';
    const ext = event.extendedProps || {};
    const empNo = ext.empNo;
    const employee = empNo != null && employees ? employees[empNo] : null;
    const typeLabel = ext.type === 'rest' ? 'Rest' : 'Work';
    const shiftCode = ext.shiftCode;
    const shiftTimeRaw = shiftCode && shiftTimes[shiftCode] ? shiftTimes[shiftCode] : '';
    const formattedShiftTime = formatShiftTime(shiftTimeRaw);
    const shiftCodeText = shiftCode ? escapeHtml(shiftCode) : 'N/A';
    const shiftCodeWithTime = shiftCode
      ? `${shiftCodeText}${formattedShiftTime ? `&nbsp;&nbsp;${escapeHtml(formattedShiftTime)}` : ''}`
      : 'N/A';
    const employeeNumber = employee && employee.empNo != null ? employee.empNo : (empNo ?? 'N/A');
    const position = employee && employee.position ? employee.position : 'N/A';
    return `
      <div class='text-sm space-y-1'>
        <div><strong>Employee #:</strong> ${escapeHtml(employeeNumber)}</div>
        <div><strong>Position:</strong> ${escapeHtml(position)}</div>
        <div><strong>Shift Code:</strong> ${shiftCodeWithTime}</div>
        <div><strong>Shift:</strong> ${escapeHtml(typeLabel)}</div>
      </div>
    `;
  }

  function refreshEventTooltip(event, el) {
    if (!event || typeof tippy === 'undefined') return;
    const target = el || (typeof findEventElementByEvent === 'function' ? findEventElementByEvent(event) : null);
    if (!target) return;
    const content = buildEventTooltipContent(event);
    try {
      if (target._tippy) target._tippy.setContent(content);
      else tippy(target, { content, allowHTML: true, theme: 'light-border', placement: 'top' });
    } catch (e) {}
  }

  // TomSelect init (kept)
  const shiftSelect = (typeof TomSelect !== 'undefined' && document.querySelector("#shift-preset")) ? new TomSelect("#shift-preset", {
    create: false,
    sortField: { field: "text", direction: "asc" },
    placeholder: "Search or select shift code...",
    render: {
      option: function (data, escape) {
        const code = data.value, time = shiftTimes[code] || "", ft = formatShiftTime(time);
        return `<div>${escape(code)} ${ft ? `(${escape(ft)})` : ""}</div>`;
      },
      item: function (data, escape) {
        const code = data.value, time = shiftTimes[code] || "", ft = formatShiftTime(time);
        return `<div>${escape(code)} ${ft ? `(${escape(ft)})` : ""}</div>`;
      }
    }
  }) : null;
  if (shiftSelect) {
    shiftSelect.on('change', function (value) {
      const clean = value.split(' ')[0];
      if (shiftPresetSelect) shiftPresetSelect.value = clean;
    });
  }

  // Position lists + presets
  const positionOptions = ['Branch Head','Site Supervisor','OIC','Mac Expert','Cashier'];
  const managerPositions = ['Branch Head','Site Supervisor','OIC'];
  const shiftPresets = [ /* (kept) */ 
    'RBG-001','RBG-002','RBG-003','RBG-004','RBG-005','RBG-006','RBG-007','RBG-008','RBG-009','RBG-010',
    'RBG-011','RBG-012','RBG-013','RBG-014','RBG-015','RBG-016','RBG-017','RBG-018','RBG-019','RBG-020',
    'RBG-021','RBG-022','RBG-023','RBG-024','RBG-025','RBG-026','RBG-027','RBG-028','RBG-029','RBG-030',
    'RBG-031','RBT-032','RBT-033','RBT-034','RBT-035','RBT-036','RBT-037','RBT-038','RBT-039','RBT-040',
    'RBT-041','RBT-042','RBT-043','RBT-044','RBT-045','RBT-046','RBT-047','RBT-048','RBT-049','RBT-050',
    'RBT-051','RBT-052','RBT-053','RBT-054',
    'AASP-001','AASP-002','AASP-003','AASP-004','AASP-005','AASP-006','AASP-007','AASP-008','AASP-009','AASP-010',
    'AASP-011','AASP-012','AASP-013','AASP-014','AASP-015','AASP-016','AASP-017','AASP-018','AASP-019','AASP-020',
    'AASP-021','AASP-021 (10-5)','AASP-022','AASP-023','AASP-024','AASP-025','AASP-026','AASP-027','AASP-028','AASP-029',
    'AASP-030','AASP-031','AASP-032','AASP-033','AASP-034','AASP-035','AASP-036','AASP-037','AASP-038','AASP-039',
    'AASP-040','AASP-041','AASP-042','AASP-043','AASP-044',
    'WHSE-001','WHSE-002','WHSE-003','WHSE-004','WHSE-005','WHSE-006','WHSE-007','WHSE-008','WHSE-009','WHSE-010',
    'WHSE-011','WHSE-012'
  ];

  /* ===========================
     COPY/PASTE DRAG STATE
  =========================== */
  let copiedEmployeeSchedule = null;
  let copiedEmployeeNo = null;
  let copiedEmployeeData = null;
  let calendar, draggable;
  let currentDroppingEvent = null;
  let currentDeletingEvent = null;
  let isEditingExistingEvent = false;
  let hoveredScheduleId = null;
  let pasteHistory = [];
  let copiedSchedules = [];
  let selectedSchedules = new Set();
  let elementSelectionCounter = 0;
  const selectedElementRegistry = new Map();
  const scheduleSelectionTargets = new Map();
  let selectedTargetDates = new Set();
  let lastMouseX = 0, lastMouseY = 0;
  let isDragging = false, dragGhost = null;
  let isSelectingDates = false, dateSelectStartEl = null;
  let dragAnchorDate = null;
  let dragSelectionEvents = [];

  if (!window.__multiSelectModifierStateInit) {
    window.__multiSelectModifierStateInit = true;
    const updateModifierState = active => { multiSelectModifierActive = !!active; };
    const handleModifierKeyDown = e => { if (!e) return; if (e.metaKey || e.ctrlKey || e.key === 'Meta' || e.key === 'Control') updateModifierState(true); };
    const handleModifierKeyUp = e => { if (!e) return; if (!e.metaKey && !e.ctrlKey) updateModifierState(false); };
    document.addEventListener('keydown', handleModifierKeyDown, true);
    document.addEventListener('keyup', handleModifierKeyUp, true);
    window.addEventListener('blur', () => updateModifierState(false));
    document.addEventListener('visibilitychange', () => { if (document.visibilityState !== 'visible') updateModifierState(false); });
  }

  const draggableCardsRoot = draggableCardsContainer;

  /* ===========================
     SAVE / LOAD LOCALSTORAGE
  =========================== */
  function saveToLocalStorage() {
    try {
      const allEvents = calendar ? calendar.getEvents().map(e => ({
        title: e.title,
        start: e.startStr,
        classNames: (() => {
          const classes = e?.getProp ? e.getProp('classNames') : e.classNames;
          return Array.isArray(classes) ? classes : undefined;
        })(),
        extendedProps: e.extendedProps
      })) : [];
      localStorage.setItem('employees', JSON.stringify(employees));
      localStorage.setItem('events', JSON.stringify(allEvents));
      localStorage.setItem('employeeColors', JSON.stringify(employeeColors));
    } catch (err) { console.warn('saveToLocalStorage error', err); }
  }

  function loadFromLocalStorage() {
    try {
      const storedEmployees = JSON.parse(localStorage.getItem('employees'));
      const storedEvents = JSON.parse(localStorage.getItem('events'));

      if (storedEmployees && Object.keys(storedEmployees).length > 0) {
        employees = storedEmployees;
        if (draggableCardsContainer) draggableCardsContainer.innerHTML = '';
        employeeColors = JSON.parse(localStorage.getItem('employeeColors')) || employeeColors || {};
        pruneEmployeeColors(Object.keys(employees));

        Object.values(employees).forEach(emp => {
          const color = ensureEmployeeColor(emp.empNo);
          const workGradient = getGradientFromBaseColor(color, 'work');
          const restGradient = getGradientFromBaseColor(color, 'rest');

          const workEventData = {
            title: emp.name,
            classNames: ['fc-event-pill','fc-event-work'],
            extendedProps: {
              type: 'work',
              empNo: emp.empNo,
              position: emp.position,
              forceStyle: { background: workGradient, backgroundImage: 'none', color: '#fff', border: 'none' }
            }
          };
          const restEventData = {
            title: emp.name,
            classNames: ['fc-event-pill','fc-event-rest'],
            extendedProps: {
              type: 'rest',
              empNo: emp.empNo,
              position: emp.position,
              forceStyle: { background: restGradient, backgroundImage: 'none', color, border: `2px solid ${color}` }
            }
          };

          const workStyle = `background:${workGradient}; color:#fff; border:none;`;
          const restStyle = `background:${restGradient}; color:${color}; border:2px solid ${color};`;

          const safeName = escapeHtml(emp.name);
          const safePosition = escapeHtml(emp.position);
          const cardHtml = `
            <div class="p-3 rounded-lg shadow-sm border border-gray-200" data-empno="${emp.empNo}"
                 style="border-left: 6px solid ${color}; background: linear-gradient(to right, ${color}22, ${color}08);">
              <div class="flex items-center justify-between gap-3">
                <div class="min-w-0">
                  <div class="font-medium text-gray-800 draggable-employee-name" title="${safeName}">${safeName}</div>
                  <div class="text-xs text-gray-500 draggable-employee-position" title="${safePosition}">${safePosition}</div>
                </div>
                <div class="flex space-x-2">
                  <div class='fc-event-pill fc-event-work px-3 py-1 text-xs font-medium rounded-full' data-event='${JSON.stringify(workEventData)}' style="${workStyle}">🟦 Work</div>
                  <div class='fc-event-pill fc-event-rest px-3 py-1 text-xs font-medium rounded-full' data-event='${JSON.stringify(restEventData)}' style="${restStyle}">🔴 Rest</div>
                </div>
              </div>
            </div>`;
          if (draggableCardsContainer) draggableCardsContainer.innerHTML += cardHtml;
        });

        try { localStorage.setItem('employeeColors', JSON.stringify(employeeColors)); } catch (e) {}
        if (draggablePlaceholder) draggablePlaceholder.classList.add('hidden');
        initializeDraggable();
        refreshEmployeeColors(Object.keys(employees));
      }

      if (storedEvents && storedEvents.length > 0 && calendar) {
        storedEvents.forEach(ev => {
          const addedEvent = calendar.addEvent({
            title: ev.title,
            start: ev.start,
            classNames: ev.classNames,
            extendedProps: ev.extendedProps
          });
          if (addedEvent) {
            const type = addedEvent.extendedProps?.type || ev.extendedProps?.type || 'work';
            const empNo = addedEvent.extendedProps?.empNo || ev.extendedProps?.empNo;
            const color = ensureEmployeeColor(empNo);
            const existingClasses = addedEvent?.getProp ? addedEvent.getProp('classNames') : addedEvent.classNames;
            const classSet = new Set([...(Array.isArray(existingClasses) ? existingClasses : [])]);
            classSet.add('fc-event-pill');
            classSet.add(type === 'rest' ? 'fc-event-rest' : 'fc-event-work');
            try { addedEvent.setProp('classNames', Array.from(classSet)); } catch (e) {}
            if (color && !addedEvent.extendedProps?.forceStyle) {
              const gradient = getGradientFromBaseColor(color, type === 'rest' ? 'rest' : 'work');
              const forceStyle = type === 'rest'
                ? { background: gradient, backgroundImage: 'none', color, border: `2px solid ${color}` }
                : { background: gradient, backgroundImage: 'none', color: '#fff', border: 'none' };
              try {
                addedEvent.setExtendedProp('forceStyle', forceStyle);
                addedEvent.setProp('backgroundColor', '');
                addedEvent.setProp('borderColor', type === 'rest' ? color : 'transparent');
                addedEvent.setProp('textColor', type === 'rest' ? color : '#fff');
              } catch (e) {}
            }
            decorateEventLater(addedEvent);
          }
        });
        runConflictDetection();
        updateStats();
      }
    } catch (err) { console.warn('loadFromLocalStorage error', err); }
  }

  function refreshEmployeeColors(empNos = []) {
    const targets = empNos.length ? empNos : Object.keys(employees || {});
    targets.forEach(empNo => updateEmployeeColorStyles(empNo));
    try { localStorage.setItem('employeeColors', JSON.stringify(employeeColors)); } catch (e) {}
  }

  function updateEmployeeColorStyles(empNo) {
    if (!empNo) return null;
    const color = ensureEmployeeColor(empNo); if (!color) return null;
    const workGradient = getGradientFromBaseColor(color, 'work');
    const restGradient = getGradientFromBaseColor(color, 'rest');

    if (draggableCardsContainer) {
      const selector = `[data-empno="${escapeSelector(empNo)}"]`;
      draggableCardsContainer.querySelectorAll(selector).forEach(card => {
        card.style.borderLeft = `6px solid ${color}`;
        card.style.background = `linear-gradient(to right, ${color}22, ${color}08)`;
        const workPill = card.querySelector('.fc-event-work');
        const restPill = card.querySelector('.fc-event-rest');
        if (workPill) {
          workPill.style.background = workGradient;
          workPill.style.color = '#fff';
          workPill.style.border = 'none';
          try {
            const payload = JSON.parse(workPill.getAttribute('data-event') || '{}');
            payload.extendedProps = payload.extendedProps || {};
            payload.extendedProps.forceStyle = { background: workGradient, backgroundImage: 'none', color: '#fff', border: 'none' };
            payload.extendedProps.empNo = empNo;
            const classSet = new Set([...(Array.isArray(payload.classNames) ? payload.classNames : []), 'fc-event-pill', 'fc-event-work']);
            payload.classNames = Array.from(classSet);
            workPill.setAttribute('data-event', JSON.stringify(payload));
          } catch (e) {}
        }
        if (restPill) {
          restPill.style.background = restGradient;
          restPill.style.color = color;
          restPill.style.border = `2px solid ${color}`;
          try {
            const payload = JSON.parse(restPill.getAttribute('data-event') || '{}');
            payload.extendedProps = payload.extendedProps || {};
            payload.extendedProps.forceStyle = { background: restGradient, backgroundImage: 'none', color, border: `2px solid ${color}` };
            payload.extendedProps.empNo = empNo;
            const classSet = new Set([...(Array.isArray(payload.classNames) ? payload.classNames : []), 'fc-event-pill', 'fc-event-rest']);
            payload.classNames = Array.from(classSet);
            restPill.setAttribute('data-event', JSON.stringify(payload));
          } catch (e) {}
        }
      });
    }

    if (calendar && typeof calendar.getEvents === 'function') {
      const events = calendar.getEvents().filter(ev => String(ev.extendedProps?.empNo) === String(empNo));
      events.forEach(event => {
        const type = event.extendedProps?.type === 'rest' ? 'rest' : 'work';
        const gradient = type === 'rest' ? restGradient : workGradient;
        const forceStyle = type === 'rest'
          ? { background: gradient, backgroundImage: 'none', color, border: `2px solid ${color}` }
          : { background: gradient, backgroundImage: 'none', color: '#fff', border: 'none' };
        event.setExtendedProp('forceStyle', forceStyle);
        try {
          event.setProp('backgroundColor', '');
          event.setProp('borderColor', type === 'rest' ? color : 'transparent');
          event.setProp('textColor', type === 'rest' ? color : '#fff');
        } catch (e) {}
        decorateEventLater(event);
      });
    }
    return { color, workGradient, restGradient };
  }

  /* ===========================
     CALENDAR INIT
  =========================== */
  function initializeCalendar() {
    if (!calendarEl) return;

    calendar = new FullCalendar.Calendar(calendarEl, {
      initialView: 'dayGridMonth',
      editable: true,
      droppable: true,
      selectable: true,
      dayMaxEvents: true,
      height: 'auto',
      headerToolbar: { left: 'prev,next today', center: 'title', right: 'dayGridMonth,dayGridWeek,dayGridDay' },

      /* -------------------------
         FIXED: eventReceive
      ------------------------- */
      eventReceive: function (info) {
        // FIX: reset any stale multi-select/batch state before handling first drop
        try {
          if (typeof selectedSchedules !== 'undefined' && selectedSchedules.clear) selectedSchedules.clear();
          if (typeof selectedTargetDates !== 'undefined' && selectedTargetDates.clear) selectedTargetDates.clear();
          window.__multiSelectDragActive = false;
          document.body.classList.remove('multi-dragging');
        } catch (e) {}

        const newEvent = info.event;
        // FIX: isolate props & ensure a unique ID
        try {
          const nowId = `evt_${Date.now()}_${Math.random().toString(36).slice(2,6)}`;
          const ext = JSON.parse(JSON.stringify(newEvent.extendedProps || {}));
          if (!ext.id && !newEvent.id) {
            ext.id = nowId;
            newEvent.setExtendedProp('id', nowId);
            try { newEvent.setProp('id', nowId); } catch (e) {}
          }
          // keep employee number for schedule generation
          if (ext.empNo == null && newEvent.extendedProps?.empNo != null) {
            newEvent.setExtendedProp('empNo', String(newEvent.extendedProps.empNo));
          }
          if (ext.employeeNo == null && newEvent.extendedProps?.employeeNo != null) {
            newEvent.setExtendedProp('employeeNo', String(newEvent.extendedProps.employeeNo));
          }
          // ensure type
          const safeType = ext.type || newEvent.extendedProps?.type || 'work';
          newEvent.setExtendedProp('type', safeType);
        } catch (e) { console.warn('eventReceive: cloning props failed', e); }

        const type = newEvent.extendedProps?.type || 'work';
        const empNo = newEvent.extendedProps?.empNo || newEvent.extendedProps?.employeeNo;
        let dateStr = newEvent.startStr;   // <-- make it let so we can update after pointer-fix
        // ---- FORCE DROP TO EXACT DATE UNDER THE POINTER ----
(function forceDropToPointerDate() {
  const pointedDate = (typeof getDateUnderPointer === 'function') ? getDateUnderPointer() : null;
  if (pointedDate && pointedDate !== newEvent.startStr) {
    newEvent.setStart(pointedDate);
    try { newEvent.setEnd(pointedDate); } catch (e) {}
    dateStr = newEvent.startStr; // keep duplicate checks in sync with the corrected date
  }
})();
        if (!empNo) { console.warn('eventReceive: missing empNo'); return; }

        // Style
        const color = ensureEmployeeColor(empNo);
        const grad = getGradientFromBaseColor(color, type);
        const classNames = ['fc-event-pill', `fc-event-${type}`];
        const forceStyle = (type === 'rest')
          ? { background: grad, backgroundImage: 'none', color, border: `2px solid ${color}` }
          : { background: grad, backgroundImage: 'none', color: '#fff', border: 'none' };

        try {
          newEvent.setProp('classNames', classNames);
          newEvent.setExtendedProp('forceStyle', forceStyle);
          newEvent.setProp('backgroundColor', '');
          newEvent.setProp('borderColor', type === 'rest' ? color : 'transparent');
          newEvent.setProp('textColor', type === 'rest' ? color : '#fff');
        } catch (e) {}

        // FIX: duplicate check only for work
        if (type === 'work') {
          const conflict = findExistingShiftForEmployee(empNo, dateStr, { exclude: [newEvent] });
          if (conflict) {
            const name = employees[empNo]?.name || empNo;
            showToast(`Duplicate entry blocked: ${name} already has a shift on this date.`, 'error');
            newEvent.remove();
            return;
          }
        }

        // Branch logic
        if (type === 'work') {
          currentDroppingEvent = newEvent;
          openShiftModal();
        } else {
          runConflictDetection();
          updateStats();
          saveToLocalStorage();
        }

        // FIX: force immediate DOM update (prevents first-drop "N/A")
        try { newEvent.setProp('title', newEvent.title); } catch (e) {}

        // apply decorations
        try { decorateEventLater(newEvent); } catch (e) {}
      },

      /* Drag inside calendar */
      eventDrop: function(info) {
        const eventId = info.event.extendedProps?.id || info.event.id;
        const multiDragActive = (typeof window !== 'undefined' && window.__multiSelectDragActive) || isDragging;
        if (multiDragActive || (selectedSchedules && selectedSchedules.size > 1 && selectedSchedules.has(String(eventId)))) {
          info.revert(); return;
        }
        const movedEmpNo = info.event.extendedProps?.empNo;
        if (movedEmpNo) {
          const conflict = findExistingShiftForEmployee(movedEmpNo, info.event.startStr, { exclude: [info.event] });
          if (conflict) {
            const employeeName = employees[movedEmpNo] ? employees[movedEmpNo].name : movedEmpNo;
            showToast(`Move blocked: ${employeeName} already has a shift on this date.`, 'error');
            info.revert(); return;
          }
        }
        runConflictDetection();
        updateStats();
        saveToLocalStorage();
      },

      eventRemove: function() { runConflictDetection(); updateStats(); saveToLocalStorage(); },

      eventClick: function(info) {
        if (info.jsEvent && hasMultiSelectModifier(info.jsEvent)) { info.jsEvent.preventDefault(); return; }
        openDeleteModal(info.event);
      },

      eventDidMount: function(info) {
        try {
          if (!info.event.extendedProps?.id) {
            const fallbackId = (typeof crypto !== 'undefined' && crypto.randomUUID) ? crypto.randomUUID() : (info.event.id || `evt-${Date.now()}`);
            info.event.setExtendedProp('id', fallbackId);
          }
        } catch (e) {}

        decorateSingleEventElement(info.event, info.el);

        const { type, shiftCode, isConflict, empNo } = info.event.extendedProps;
        const emp = employees[empNo];

        if (empNo) { try { info.el.setAttribute('data-empno', empNo); info.el.dataset.empno = String(empNo); } catch (e) {} }
        if (!emp) { info.el.style.display = 'none'; return; }

        const color = ensureEmployeeColor(empNo);
        if (type === 'work') {
          try {
            const grad = getGradientFromBaseColor(color, 'work');
            info.el.style.setProperty('background', grad, 'important');
            info.el.style.setProperty('color', '#fff', 'important');
            info.el.style.setProperty('border', 'none', 'important');
            const inner = info.el.querySelector('.fc-event-main-frame') || info.el;
            inner.style.setProperty('background', grad, 'important');
            inner.style.setProperty('color', '#fff', 'important');
          } catch (e) {
            info.el.style.backgroundColor = color;
            info.el.style.backgroundImage = 'none';
            info.el.style.color = '#fff';
            info.el.style.border = 'none';
            const inner = info.el.querySelector('.fc-event-main-frame') || info.el;
            inner.style.backgroundColor = color; inner.style.backgroundImage = 'none'; inner.style.color = '#fff';
          }
        } else if (type === 'rest') {
          try {
            const grad = getGradientFromBaseColor(color, 'rest');
            info.el.style.setProperty('background', grad, 'important');
            info.el.style.setProperty('color', color, 'important');
            info.el.style.setProperty('border', `2px solid ${color}`, 'important');
            const inner = info.el.querySelector('.fc-event-main-frame') || info.el;
            inner.style.setProperty('background', grad, 'important');
            inner.style.setProperty('color', color, 'important');
            inner.style.setProperty('border', `2px solid ${color}`, 'important');
          } catch (e) {
            info.el.style.backgroundColor = '#fff';
            info.el.style.backgroundImage = 'none';
            info.el.style.border = `2px solid ${color}`;
            info.el.style.color = color;
            const inner = info.el.querySelector('.fc-event-main-frame') || info.el;
            inner.style.backgroundColor = '#fff';
            inner.style.backgroundImage = 'none';
            inner.style.border = `2px solid ${color}`;
            inner.style.color = color;
          }
        }

        refreshEventTooltip(info.event, info.el);
        info.el.classList.add(type === 'work' ? 'fc-event-work' : 'fc-event-rest');
        if (isConflict) info.el.classList.add('fc-event-conflict'); else info.el.classList.remove('fc-event-conflict');

        const shiftInfo = shiftCode ? ` (${shiftCode})` : '';
        const title = `${info.event.title} (${emp.position})\nType: ${type.charAt(0).toUpperCase()+type.slice(1)}${shiftInfo}`;
        info.el.title = title;
      },

      eventContent: function(arg) {
        const { shiftCode } = arg.event.extendedProps;
        const shiftHtml = shiftCode ? `<span class="pill-shift">${shiftCode}</span>` : '';
        return { html: `<div class="pill-content"><span class="pill-name">${arg.event.title}</span>${shiftHtml}</div>` };
      },

      dayCellClassNames: function(arg) {
        if (arg.date.getDay() === 0 || arg.date.getDay() === 6) return 'fc-day-sat-sun';
        return null;
      }
    });

    // FIX: prevent transforms during native event drag
    calendar.on('eventDragStart', function(){ document.body.classList.add('no-transform-during-drag'); });
    calendar.on('eventDragStop', function(){ document.body.classList.remove('no-transform-during-drag'); });

    calendar.render();
  }

  /* ===========================
     DRAGGABLE (SIDEBAR) INIT
  =========================== */
  function initializeDraggable() {
    if (!draggableCardsContainer) return;
    if (draggable) { try { draggable.destroy(); } catch (e) {} }

    try {
      draggable = new FullCalendar.Draggable(draggableCardsContainer, {
        itemSelector: '.fc-event-pill',
        eventData: function (eventEl) {
          try {
            if (!eventEl._fcEventDataTemplate) {
              const raw = eventEl.getAttribute('data-event');
              eventEl._fcEventDataTemplate = raw ? JSON.parse(raw) : null;
            }
            const template = eventEl._fcEventDataTemplate;
            if (!template) return null;

            // FIX: deep clone for every drop to isolate objects
            const parsed = JSON.parse(JSON.stringify(template));
            // assign a unique temp id to prevent shared id
            const uid = `evt_${Date.now()}_${Math.random().toString(36).slice(2,6)}`;
            parsed.id = parsed.id || uid;
            if (parsed.extendedProps) parsed.extendedProps.id = parsed.extendedProps.id || uid;

            const t = parsed.extendedProps && parsed.extendedProps.type ? parsed.extendedProps.type : '';
            const classSet = new Set(parsed.classNames || []);
            classSet.add('fc-event-pill');
            if (t === 'rest') classSet.add('fc-event-rest'); else classSet.add('fc-event-work');
            parsed.classNames = Array.from(classSet);
            return parsed;
          } catch (err) {
            console.error('Invalid event data on draggable element', err);
            return null;
          }
        }
      });
    } catch (err) { console.warn('initializeDraggable error', err); }

    // Cosmetic lift while dragging inside calendar
    if (calendar && typeof calendar.on === 'function') {
      calendar.on('eventDragStart', function(info) {
        info.el.style.transform = 'translateY(-4px)';
      });
      calendar.on('eventDragStop', function(info) {
        info.el.style.transform = 'translateY(0)';
      });
    }

    // Also guard against transforms during external HTML5 drag
    draggableCardsRoot?.addEventListener('dragstart', () => { document.body.classList.add('no-transform-during-drag'); });
    draggableCardsRoot?.addEventListener('dragend',   () => { document.body.classList.remove('no-transform-during-drag'); });
  }

  /* ===========================
     EMPLOYEE TABLE UI (kept)
  =========================== */
  function addEmployeeRow() {
    if (!employeeTableBody) {
      console.warn('employee-table-body not found in DOM. Creating fallback table.');
      const sidebar = document.querySelector('aside') || document.body;
      const wrapper = document.createElement('div');
      wrapper.innerHTML = `
        <div class="p-4 bg-white border border-gray-200 rounded-xl shadow-sm">
          <div class="overflow-x-auto">
            <table class="w-full text-sm">
              <thead class="text-left text-gray-500">
                <tr>
                  <th class="p-2">Name</th><th class="p-2">Employee No.</th><th class="p-2">Position</th><th class="p-2 w-10"></th>
                </tr>
              </thead>
              <tbody id="employee-table-body"></tbody>
            </table>
          </div>
        </div>`;
      sidebar.insertBefore(wrapper, sidebar.firstChild);
      employeeTableBody = document.getElementById('employee-table-body');
    }
    const tbody = document.getElementById('employee-table-body');
    if (!tbody) { console.error('Unable to locate or create employee table body.'); return; }

    const tr = document.createElement('tr');
    tr.className = 'align-top';
    const positionDropdown = positionOptions.map(pos => `<option value="${pos}">${pos}</option>`).join('');
    tr.innerHTML = `
      <td class="p-1"><input type="text" class="emp-name w-full text-sm border-gray-300 rounded-md shadow-sm focus:border-blue-500 focus:ring-blue-500" placeholder="J. Dela Cruz"></td>
      <td class="p-1"><input type="text" class="emp-no w-full text-sm border-gray-300 rounded-md shadow-sm focus:border-blue-500 focus:ring-blue-500" placeholder="12345"></td>
      <td class="p-1">
        <select class="emp-pos w-full text-sm border-gray-300 rounded-md shadow-sm focus:border-blue-500 focus:ring-blue-500">
          <option value="">Select...</option>${positionDropdown}
        </select>
      </td>
      <td class="p-1 text-center">
        <button class="remove-row-btn p-1 text-red-400 hover:text-red-600 rounded-full" title="Remove row">
          <svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" viewBox="0 0 20 20" fill="currentColor">
            <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clip-rule="evenodd" />
          </svg>
        </button>
      </td>`;
    tbody.appendChild(tr);
  }

  function removeEmployeeRow(buttonEl) {
    if (!employeeTableBody) return;
    if (employeeTableBody.rows.length > 1) {
      const tr = buttonEl.closest('tr'); if (tr) tr.remove();
    } else { showToast('At least one employee row is required.', 'warn'); }
  }

  function saveAndGenerate() {
    employees = {};
    let isValid = true;

    if (draggableCardsContainer) draggableCardsContainer.innerHTML = '';
    if (employeeTableBody) {
      document.querySelectorAll('#employee-table-body input, #employee-table-body select').forEach(el => el.classList.remove('validation-error'));
    }

    const rows = employeeTableBody ? employeeTableBody.querySelectorAll('tr') : [];
    rows.forEach(row => {
      const nameInput = row.querySelector('.emp-name');
      const noInput   = row.querySelector('.emp-no');
      const posInput  = row.querySelector('.emp-pos');
      const name = nameInput ? nameInput.value.trim() : '';
      const empNo = noInput ? noInput.value.trim() : '';
      const position = posInput ? posInput.value : '';

      if (!name && nameInput) { nameInput.classList.add('validation-error'); isValid = false; }
      if (!empNo && noInput) { noInput.classList.add('validation-error'); isValid = false; }
      if (!position && posInput) { posInput.classList.add('validation-error'); isValid = false; }
      if (empNo && employees[empNo]) { if (noInput) noInput.classList.add('validation-error'); isValid = false; showToast(`Duplicate Employee No: ${empNo}.`, 'error'); }
      if (isValid && name && empNo && position) { employees[empNo] = { name, empNo, position }; }
    });

    if (!isValid) { showToast('Please fix validation errors in the employee table.', 'error'); return; }
    if (Object.keys(employees).length === 0) { showToast('No valid employee data found. Please add at least one employee.', 'warn'); return; }

    pruneEmployeeColors(Object.keys(employees));

    Object.values(employees).forEach(emp => {
      const color = ensureEmployeeColor(emp.empNo);
      const workGradient = getGradientFromBaseColor(color, 'work');
      const restGradient = getGradientFromBaseColor(color, 'rest');

      const workEventData = {
        title: emp.name,
        classNames: ['fc-event-pill','fc-event-work'],
        extendedProps: {
          type: 'work',
          empNo: emp.empNo,
          position: emp.position,
          forceStyle: { background: workGradient, backgroundImage: 'none', color: '#fff', border: 'none' }
        }
      };
      const restEventData = {
        title: emp.name,
        classNames: ['fc-event-pill','fc-event-rest'],
        extendedProps: {
          type: 'rest',
          empNo: emp.empNo,
          position: emp.position,
          forceStyle: { background: restGradient, backgroundImage: 'none', color, border: `2px solid ${color}` }
        }
      };

      const safeName = escapeHtml(emp.name);
      const safePosition = escapeHtml(emp.position);
      const cardHtml = `
        <div class="p-3 rounded-lg shadow-sm border border-gray-200" data-empno="${emp.empNo}"
             style="border-left: 6px solid ${color}; background: linear-gradient(to right, ${color}22, ${color}08);">
          <div class="flex items-center justify-between gap-3">
            <div class="min-w-0">
              <div class="font-medium text-gray-800 draggable-employee-name" title="${safeName}">${safeName}</div>
              <div class="text-xs text-gray-500 draggable-employee-position" title="${safePosition}">${safePosition}</div>
            </div>
            <div class="flex space-x-2">
              <div class="fc-event-pill fc-event-work px-3 py-1 text-xs font-medium rounded-full cursor-pointer select-none shadow-sm"
                   data-event='${JSON.stringify(workEventData)}'
                   style="background:${workGradient}; color:#fff; border:none;">🟦 Work</div>
              <div class="fc-event-pill fc-event-rest px-3 py-1 text-xs font-medium rounded-full cursor-pointer select-none shadow-sm"
                   data-event='${JSON.stringify(restEventData)}'
                   style="background:${restGradient}; color:${color}; border:2px solid ${color};">🔴 Rest</div>
            </div>
          </div>
        </div>`;
      if (draggableCardsContainer) draggableCardsContainer.innerHTML += cardHtml;
    });

    if (draggablePlaceholder) draggablePlaceholder.classList.add('hidden');
    initializeDraggable();
    updateStats();
    refreshEmployeeColors(Object.keys(employees));
    showToast('Employees saved and cards generated. You can now drag schedules.', 'success');
    saveToLocalStorage();

    document.querySelectorAll('#draggable-cards-container > div').forEach(card => {
      const nameEl = card.querySelector('.font-medium');
      const name = nameEl ? nameEl.textContent.trim() : '';
      const emp = Object.values(employees).find(e => e.name === name);
      if (emp) card.dataset.empno = emp.empNo;
    });
  }

  /* ===========================
     SHIFT MODAL & SAVE
  =========================== */
  function openShiftModal(options = {}) {
    if (!shiftModal) return;
    const { initialShiftCode = '', useLastUsedWhenEmpty = true } = options;
    const normalizedInitial = initialShiftCode ? String(initialShiftCode).trim() : '';

    if (shiftSelect) { try { shiftSelect.clear(); } catch (e) {} }
    if (shiftPresetSelect) shiftPresetSelect.value = '';
    if (shiftCustomInput) shiftCustomInput.value = '';

    if (shiftPresetSelect && shiftPresetSelect.options.length <= 1) {
      shiftPresets.forEach(code => { const option = new Option(code, code); shiftPresetSelect.add(option); });
    }

    if (normalizedInitial) {
      const cleanInitial = normalizedInitial.split(' ')[0];
      let presetMatched = false;
      if (shiftPresetSelect) {
        const optionsArr = Array.from(shiftPresetSelect.options);
        presetMatched = optionsArr.some(opt => opt.value === cleanInitial);
        if (presetMatched) shiftPresetSelect.value = cleanInitial;
      }
      if (shiftSelect && presetMatched) { try { shiftSelect.setValue(cleanInitial); } catch (e) {} }
      if (shiftCustomInput) shiftCustomInput.value = presetMatched ? '' : cleanInitial;
    } else if (useLastUsedWhenEmpty && window.lastUsedShiftCode) {
      const last = String(window.lastUsedShiftCode).split(' ')[0];
      if (shiftSelect) { try { shiftSelect.setValue(last); } catch (e) { if (shiftPresetSelect) shiftPresetSelect.value = last; } }
      else if (shiftPresetSelect) shiftPresetSelect.value = last;
    }

    shiftModal.classList.remove('hidden');
    setTimeout(() => {
      const tomInput = document.querySelector('.ts-dropdown .ts-input input, .ts-control input');
      if (tomInput) tomInput.focus();
    }, 200);
  }

  const shiftSearchInput = document.getElementById('shift-search');
  if (shiftSearchInput && shiftPresetSelect) {
    shiftSearchInput.addEventListener('input', function() {
      const search = this.value.toLowerCase();
      for (const option of shiftPresetSelect.options) {
        const match = option.text.toLowerCase().includes(search);
        option.style.display = match ? '' : 'none';
      }
    });
  }

  function handleSaveShift() {
    let preset = '';
    try { preset = shiftSelect ? (shiftSelect.getValue() || shiftPresetSelect.value) : (shiftPresetSelect ? shiftPresetSelect.value : ''); }
    catch (e) { preset = shiftPresetSelect ? shiftPresetSelect.value : ''; }

    const custom = shiftCustomInput ? shiftCustomInput.value.trim() : '';
    if (!preset && !custom && window.lastUsedShiftCode) preset = window.lastUsedShiftCode;

    const shiftCode = custom || preset;
    if (!shiftCode) { showToast('Please select a preset or enter a custom shift code.', 'warn'); return; }
    window.lastUsedShiftCode = shiftCode;

    const cleanCode = String(shiftCode).split(' ')[0];
    if (currentDroppingEvent) {
      currentDroppingEvent.setExtendedProp('shiftCode', cleanCode);
      // FIX: immediate repaint so pill shows code on 1st drop
      try { currentDroppingEvent.setProp('title', currentDroppingEvent.title); } catch (e) {}
      const dropType = currentDroppingEvent.extendedProps?.type || 'work';
      const existingClassNames = currentDroppingEvent?.getProp ? currentDroppingEvent.getProp('classNames') : currentDroppingEvent.classNames;
      const classSet = new Set([...(Array.isArray(existingClassNames) ? existingClassNames : []), 'fc-event-pill', `fc-event-${dropType}`]);
      try { currentDroppingEvent.setProp('classNames', Array.from(classSet)); } catch (e) {}
      decorateEventLater(currentDroppingEvent);
      refreshEventTooltip(currentDroppingEvent);
      if (calendar && typeof calendar.rerenderEvents === 'function') { calendar.rerenderEvents(); }
      runConflictDetection(); updateStats(); saveToLocalStorage();
    }
    if (shiftModal) closeModal(shiftModal);
    currentDroppingEvent = null; isEditingExistingEvent = false;
  }

  function openDeleteModal(event) {
    if (!deleteModal || !deleteModalSummary) return;
    currentDeletingEvent = event;
    const { type } = event.extendedProps;
    deleteModalSummary.textContent = `${event.title} - ${type.toUpperCase()} on ${event.start.toLocaleDateString()}`;
    deleteModal.classList.remove('hidden');
  }

  function handleConfirmDelete() {
    if (currentDeletingEvent) { currentDeletingEvent.remove(); showToast('Schedule entry removed.', 'info'); saveToLocalStorage(); }
    if (deleteModal) closeModal(deleteModal);
    currentDeletingEvent = null;
  }

  function openExportModal() {
    if (!exportModal || !calendar) return;
    const conflicts = getConflicts();
    const allEvents = calendar.getEvents();
    const workCount = allEvents.filter(e => e.extendedProps.type === 'work').length;
    const restCount = allEvents.filter(e => e.extendedProps.type === 'rest').length;
    const exportConflictWarning = document.getElementById('export-conflict-warning');
    const exportConflictCount    = document.getElementById('export-conflict-count');
    const exportWorkCount        = document.getElementById('export-work-count');
    const exportRestCount        = document.getElementById('export-rest-count');

    if (exportWorkCount) exportWorkCount.textContent = workCount;
    if (exportRestCount) exportRestCount.textContent = restCount;
    if (exportConflictWarning && exportConflictCount) {
      if (conflicts.length > 0) { exportConflictCount.textContent = conflicts.length; exportConflictWarning.classList.remove('hidden'); }
      else exportConflictWarning.classList.add('hidden');
    }
    exportModal.classList.remove('hidden');
  }

  function closeModal(modalEl) { if (!modalEl) return; modalEl.classList.add('hidden'); }

  function handleCloseModal(modalId) {
    const modalEl = document.getElementById(modalId);
    closeModal(modalEl);
    if (modalId === 'shift-modal' && currentDroppingEvent && !isEditingExistingEvent) {
      currentDroppingEvent.remove(); showToast('Schedule add canceled.', 'info'); saveToLocalStorage();
    }
    if (modalId === 'shift-modal') { currentDroppingEvent = null; isEditingExistingEvent = false; }
    if (modalId === 'delete-modal') { currentDeletingEvent = null; }
  }

  function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container') || document.body;
    const toast = document.createElement('div');
    let bgColor, textColor, borderColor;
    switch (type) {
      case 'success': bgColor='bg-green-50'; textColor='text-green-800'; borderColor='border-green-200'; break;
      case 'error':   bgColor='bg-red-50';   textColor='text-red-800';   borderColor='border-red-200';   break;
      case 'warn':    bgColor='bg-yellow-50';textColor='text-yellow-800';borderColor='border-yellow-200';break;
      default:        bgColor='bg-blue-50';  textColor='text-blue-800';  borderColor='border-blue-200';
    }
    toast.className = `toast p-4 rounded-lg shadow-lg border ${bgColor} ${textColor} ${borderColor}`;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => toast.classList.add('show'), 10);
    setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 500); }, 4000);
  }

  /* ===========================
     CONFLICT DETECTION (kept)
  =========================== */
  function getGroupedEvents() {
    const allEvents = calendar ? calendar.getEvents() : [];
    const eventsByDate = {}, eventsByEmp = {};
    allEvents.forEach(event => {
      const dateStr = event.startStr; const { empNo } = event.extendedProps;
      if (!eventsByDate[dateStr]) eventsByDate[dateStr] = []; eventsByDate[dateStr].push(event);
      if (!eventsByEmp[empNo]) eventsByEmp[empNo] = []; eventsByEmp[empNo].push(event);
    });
    return { allEvents, eventsByDate, eventsByEmp };
  }

  function getConflicts() {
    const { eventsByDate, eventsByEmp } = getGroupedEvents();
    let conflicts = [];

    for (const date in eventsByDate) {
      const dailyEvents = eventsByDate[date];
      const empEventsOnDate = {};
      dailyEvents.forEach(event => {
        const { empNo } = event.extendedProps;
        if (!empEventsOnDate[empNo]) empEventsOnDate[empNo] = [];
        empEventsOnDate[empNo].push(event);
      });
      for (const empNo in empEventsOnDate) {
        const events = empEventsOnDate[empNo];
        if (events.length > 1) {
          const hasWork = events.some(e => e.extendedProps.type === 'work');
          const hasRest = events.some(e => e.extendedProps.type === 'rest');
          if (hasWork && hasRest) {
            events.forEach(event => { conflicts.push({ empNo, date, rule: 'Work & Rest on Same Day', event }); });
          }
        }
      }
    }

    for (const date in eventsByDate) {
      const managerRestEvents = eventsByDate[date].filter(event => {
        const emp = employees[event.extendedProps.empNo];
        return event.extendedProps.type === 'rest' && emp && managerPositions.includes(emp.position);
      });
      if (managerRestEvents.length > 1) {
        managerRestEvents.forEach(event => { conflicts.push({ empNo: event.extendedProps.empNo, date, rule: 'Multiple Managers Resting', event }); });
      }
    }

    for (const empNo in eventsByEmp) {
      const weekendRestEvents = eventsByEmp[empNo].filter(event => {
        const day = event.start.getDay();
        return event.extendedProps.type === 'rest' && (day === 0 || day === 6);
      });
      if (weekendRestEvents.length > 2) {
        weekendRestEvents.forEach(event => { conflicts.push({ empNo, date: event.startStr, rule: 'Exceeds 2 Weekend Rest Days', event }); });
      }
    }
    return conflicts;
  }

  function runConflictDetection() {
    const { allEvents } = getGroupedEvents();
    if (allEvents && allEvents.forEach) allEvents.forEach(event => event.setExtendedProp('isConflict', false));
    if (conflictTableBody) conflictTableBody.innerHTML = '';

    const conflicts = getConflicts();
    const conflictEvents = new Set();
    const conflictTableEntries = {};
    conflicts.forEach(conflict => {
      conflict.event.setExtendedProp('isConflict', true);
      conflictEvents.add(conflict.event.extendedProps.id);
      const key = `${conflict.empNo}-${conflict.rule}`;
      if (!conflictTableEntries[key]) conflictTableEntries[key] = { empNo: conflict.empNo, rule: conflict.rule, dates: new Set() };
      conflictTableEntries[key].dates.add(conflict.date);
    });

    if (Object.keys(conflictTableEntries).length > 0 && conflictTableBody) {
      if (conflictsPlaceholder) conflictsPlaceholder.classList.add('hidden');
      Object.values(conflictTableEntries).forEach(entry => {
        const emp = employees[entry.empNo] || { name: entry.empNo, empNo: entry.empNo };
        const tr = document.createElement('tr');
        tr.className = 'bg-yellow-50 border-b border-yellow-200';
        tr.innerHTML = `
          <td class="p-3 font-medium text-yellow-900">${emp.name}</td>
          <td class="p-3 text-yellow-800">${emp.empNo}</td>
          <td class="p-3 text-yellow-800">${entry.rule}</td>
          <td class="p-3 text-yellow-800">${[...entry.dates].join(', ')}</td>`;
        conflictTableBody.appendChild(tr);
      });
    } else { if (conflictsPlaceholder) conflictsPlaceholder.classList.remove('hidden'); }

    updateStats(conflictEvents.size);
  }

  /* ===========================
     STATS/EXPORT/RESET
  =========================== */
  function updateStats(conflictCount = null) {
    const allEvents = calendar ? calendar.getEvents() : [];
    const workCount = allEvents.filter(e => e.extendedProps.type === 'work').length;
    const restCount = allEvents.filter(e => e.extendedProps.type === 'rest').length;
    if (conflictCount === null) conflictCount = allEvents.filter(e => e.extendedProps.isConflict).length;
    const empCount = Object.keys(employees).length;
    if (statsWork) statsWork.textContent = workCount;
    if (statsRest) statsRest.textContent = restCount;
    if (statsConflicts) statsConflicts.textContent = conflictCount;
    if (statsEmployees) statsEmployees.textContent = empCount;
  }

  function exportToExcel() {
    if (!calendar) return;

    const conflicts = typeof getConflicts === 'function' ? getConflicts() : [];
    const allEvents = calendar.getEvents();
    const daysOfWeek = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const wb = XLSX.utils.book_new();

    // --- Group events by employee number ---
    const byEmp = {};
    allEvents.forEach(e => {
      const empNo = e?.extendedProps?.empNo;
      if (!empNo) return;
      if (!byEmp[empNo]) byEmp[empNo] = { work: [], rest: [] };
      const t = e?.extendedProps?.type;
      if (t === 'work') byEmp[empNo].work.push(e);
      else if (t === 'rest') byEmp[empNo].rest.push(e);
    });

    const getEmp = (empNo) => (employees && employees[empNo])
      ? employees[empNo]
      : { name: 'N/A', empNo: empNo, position: 'N/A' };

    const empOrder = Object.keys(byEmp).sort((a,b) => {
      const ea = getEmp(a);
      const eb = getEmp(b);
      // Sort by employee name, then by employee number as tie-breaker
      return (ea.name || '').localeCompare(eb.name || '') || String(a).localeCompare(String(b));
    });

    // ===== Work sheet =====
    const workData = [['NAME','EMPLOYEE NO.','POSITION','WORK DATE','SHIFT CODE','DAY OF WEEK']];
    empOrder.forEach(empNo => {
      const emp = getEmp(empNo);
      byEmp[empNo].work.sort((a,b) => a.start - b.start);
      byEmp[empNo].work.forEach(e => {
        const shift = (e?.extendedProps?.shiftCode) || 'N/A';
        const day = daysOfWeek[e.start.getDay()];
        workData.push([emp.name, emp.empNo, emp.position || 'N/A', e.startStr, shift, day]);
      });
    });
    const wsWork = XLSX.utils.aoa_to_sheet(workData);
    try {
      for (let c = 0; c < workData[0].length; c++) {
        const addr = XLSX.utils.encode_cell({ c, r: 0 });
        if (!wsWork[addr]) continue;
        wsWork[addr].s = Object.assign({}, wsWork[addr].s || {}, { font: { bold: true } });
      }
    } catch (_) {}
    XLSX.utils.book_append_sheet(wb, wsWork, 'Work Schedule');

    // ===== Rest sheet =====
    const restData = [['NAME','EMPLOYEE NO.','POSITION','REST DAY DATE','DAY OF WEEK']];
    empOrder.forEach(empNo => {
      const emp = getEmp(empNo);
      byEmp[empNo].rest.sort((a,b) => a.start - b.start);
      byEmp[empNo].rest.forEach(e => {
        const day = daysOfWeek[e.start.getDay()];
        restData.push([emp.name, emp.empNo, emp.position || 'N/A', e.startStr, day]);
      });
    });
    const wsRest = XLSX.utils.aoa_to_sheet(restData);
    try {
      for (let c = 0; c < restData[0].length; c++) {
        const addr = XLSX.utils.encode_cell({ c, r: 0 });
        if (!wsRest[addr]) continue;
        wsRest[addr].s = Object.assign({}, wsRest[addr].s || {}, { font: { bold: true } });
      }
    } catch (_) {}
    XLSX.utils.book_append_sheet(wb, wsRest, 'Rest Schedule');

    // ===== Conflict Summary (unchanged grouping) =====
    if (conflicts && conflicts.length > 0) {
      const conflictData = [['EMPLOYEE NAME','EMPLOYEE NO.','POLICY VIOLATED','DATES INVOLVED']];
      const map = {};
      conflicts.forEach(c => {
        const key = `${c.empNo}-${c.rule}`;
        if (!map[key]) map[key] = { empNo: c.empNo, rule: c.rule, dates: new Set() };
        map[key].dates.add(c.date);
      });
      Object.values(map).forEach(entry => {
        const emp = getEmp(entry.empNo);
        conflictData.push([emp.name, emp.empNo, entry.rule, [...entry.dates].join(', ')]);
      });
      const wsCon = XLSX.utils.aoa_to_sheet(conflictData);
      try {
        for (let c = 0; c < conflictData[0].length; c++) {
          const addr = XLSX.utils.encode_cell({ c, r: 0 });
          if (!wsCon[addr]) continue;
          wsCon[addr].s = Object.assign({}, wsCon[addr].s || {}, { font: { bold: true } });
        }
      } catch (_) {}
      XLSX.utils.book_append_sheet(wb, wsCon, 'Conflict Summary');
    }

    XLSX.writeFile(wb, 'Branch_Schedule_Report.xlsx');
    if (typeof closeModal === 'function' && typeof exportModal !== 'undefined' && exportModal) closeModal(exportModal);
    if (typeof showToast === 'function') showToast('Excel report downloaded successfully!', 'success');
  }

  function resetAll() {
    if (!deleteModal || !deleteModalSummary || !confirmDeleteBtn) return;
    currentDeletingEvent = null;
    deleteModalSummary.textContent = "This will clear all employees, cards, and calendar data. This action cannot be undone.";
    deleteModal.querySelector('h3').textContent = "Reset All Data?";
    confirmDeleteBtn.textContent = "Confirm Reset";
    confirmDeleteBtn.classList.remove('bg-red-600','hover:bg-red-700','focus:ring-red-500');
    confirmDeleteBtn.classList.add('bg-orange-600','hover:bg-orange-700','focus:ring-orange-500');

    deleteModal.classList.remove('hidden');
    localStorage.clear();

    confirmDeleteBtn.onclick = () => {
      employees = {};
      if (calendar) calendar.removeAllEvents();
      if (employeeTableBody) employeeTableBody.innerHTML = '';
      addEmployeeRow();
      if (draggableCardsContainer) draggableCardsContainer.innerHTML = '';
      if (draggablePlaceholder) draggablePlaceholder.classList.remove('hidden');
      if (draggable) try { draggable.destroy(); } catch (e) {}
      if (conflictTableBody) conflictTableBody.innerHTML = '';
      if (conflictsPlaceholder) conflictsPlaceholder.classList.remove('hidden');
      updateStats(0);

      closeModal(deleteModal);
      deleteModal.querySelector('h3').textContent = "Delete Schedule Entry";
      confirmDeleteBtn.textContent = "Delete";
      confirmDeleteBtn.classList.add('bg-red-600','hover:bg-red-700','focus:ring-red-500');
      confirmDeleteBtn.classList.remove('bg-orange-600','hover:bg-orange-700','focus:ring-orange-500');
      confirmDeleteBtn.onclick = handleConfirmDelete;
      showToast('Scheduler has been reset.', 'success');
    };
  }

  /* ===========================
     EVENT DECORATION HELPERS
  =========================== */
  function getScheduleIdFromElement(el) {
    if (!el) return null;
    const idCarrier = el.closest('[data-id], [data-event-id], [data-schedule-id], [data-sched-id]');
    if (!idCarrier) return null;
    return ( idCarrier.dataset?.id ||
             idCarrier.getAttribute('data-event-id') ||
             idCarrier.getAttribute('data-schedule-id') ||
             idCarrier.getAttribute('data-sched-id') ||
             idCarrier.getAttribute('data-id') || null );
  }

  function findEventElementByEvent(event) {
    if (!event) return null;
    const candidates = [];
    if (event.id != null) candidates.push(`[data-event-id="${event.id}"]`);
    const extId = event.extendedProps?.id || event.id;
    if (extId) candidates.push(`[data-id="${extId}"]`);
    if (event.extendedProps?.empNo) candidates.push(`.fc-event[data-empno="${event.extendedProps.empNo}"]`);
    for (const selector of candidates) {
      const el = document.querySelector(selector);
      if (el) return el;
    }
    return null;
  }

  function ensurePillStructure(event, el) {
    if (!event || !el) return;
    const mainFrame = el.querySelector('.fc-event-main-frame') || el.querySelector('.fc-event-main');
    if (!mainFrame) return;

    let pillContent = mainFrame.querySelector('.pill-content');
    if (!pillContent) {
      mainFrame.innerHTML = '';
      pillContent = document.createElement('div');
      pillContent.className = 'pill-content';
      mainFrame.appendChild(pillContent);
    }

    let nameSpan = pillContent.querySelector('.pill-name');
    if (!nameSpan) { nameSpan = document.createElement('span'); nameSpan.className = 'pill-name'; pillContent.insertBefore(nameSpan, pillContent.firstChild); }
    nameSpan.textContent = event.title || '';
    nameSpan.title = event.title || '';

    const shiftCode = event.extendedProps?.shiftCode;
    let shiftSpan = pillContent.querySelector('.pill-shift');
    if (shiftCode) {
      if (!shiftSpan) { shiftSpan = document.createElement('span'); shiftSpan.className = 'pill-shift'; pillContent.appendChild(shiftSpan); }
      shiftSpan.textContent = shiftCode;
    } else if (shiftSpan) { shiftSpan.remove(); }

    refreshEventTooltip(event, el);
  }

  function decorateSingleEventElement(event, el) {
    if (!event || !el) return;
    const extId = event.extendedProps?.id || event.id;
    try {
      if (event.id != null) { el.setAttribute('data-event-id', event.id); el.dataset.eventId = String(event.id); }
      if (extId) { el.setAttribute('data-id', extId); el.dataset.id = String(extId); }
      if (event.extendedProps?.empNo) { el.setAttribute('data-empno', event.extendedProps.empNo); el.dataset.empno = String(event.extendedProps.empNo); }
    } catch (e) {}

    el.classList.add('schedule-pill','fc-event-pill');
    el.classList.remove('fc-event-work','fc-event-rest');
    if (event.extendedProps?.type === 'work') el.classList.add('fc-event-work');
    else if (event.extendedProps?.type === 'rest') el.classList.add('fc-event-rest');

    if (event.extendedProps?.forceStyle) {
      try { Object.entries(event.extendedProps.forceStyle).forEach(([k, v]) => { el.style.setProperty(k, v, 'important'); }); } catch (e) {}
    } else {
      try { ['background','backgroundImage','color','border'].forEach(prop => { el.style.removeProperty(prop); }); } catch (e) {}
    }

    ensurePillStructure(event, el);

    if (!el.__scheduleSelectionHandler) {
      const handler = e => {
        if (hasMultiSelectModifier(e)) { e.preventDefault(); e.stopPropagation(); toggleScheduleSelection(el, true); return; }
        toggleScheduleSelection(el, false);
      };
      el.addEventListener('click', handler);
      Object.defineProperty(el, '__scheduleSelectionHandler', { value: handler, enumerable: false, configurable: true, writable: false });
    }

    el.onmouseenter = () => { const id = getScheduleIdFromElement(el) || extId; if (id != null) hoveredScheduleId = String(id); };
    el.onmouseleave = () => { const id = getScheduleIdFromElement(el) || extId; if (id != null && hoveredScheduleId === String(id)) hoveredScheduleId = null; };
  }

  function decorateEventLater(event, attempt = 0) {
    const scheduler = () => {
      const el = findEventElementByEvent(event);
      if (!el) { if (attempt < 5) decorateEventLater(event, attempt + 1); return; }
      decorateSingleEventElement(event, el);
    };
    if (typeof requestAnimationFrame === 'function') requestAnimationFrame(scheduler); else setTimeout(scheduler, 0);
  }

  function decorateCalendarEvents() {
    if (!calendar) return;
    const allEvents = calendar.getEvents();
    allEvents.forEach(ev => {
      const el = findEventElementByEvent(ev);
      if (!el) return;
      decorateSingleEventElement(ev, el);
    });
  }

  /* ===========================
     CONTEXT MENU / COPY / PASTE / MULTI-DRAG
  =========================== */
  // (Kept logic) — includes: selection registry, copy/paste, context menu, keyboard, drag selection etc.
  // NOTE: we only add two FIXES below: ghost alignment & no-transform-during-drag toggles.

  const qAll = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const q = (sel, root = document) => root.querySelector(sel);

  const cssEscape = (typeof window !== 'undefined' && window.CSS?.escape) ? window.CSS.escape : str => String(str).replace(/"/g, '');

  function ensureSelectionKey(el) {
    if (!el) return null;
    try {
      if (!el.dataset) return null;
      let key = el.dataset.selectionKey;
      if (!key) { elementSelectionCounter += 1; key = `sel-${elementSelectionCounter}`; el.dataset.selectionKey = key; }
      return key;
    } catch { return null; }
  }

  function applySelectionStyles(el, active) {
    if (!el || !el.classList) return;
    const method = active ? 'add' : 'remove';
    el.classList[method]('selected','ring-2','ring-blue-500','shadow-lg');
    if (!active) el.classList.remove('drag-preview');
  }

  function unregisterSelectionByKey(selectionKey, { removeId = true } = {}) {
    if (!selectionKey) return;
    const entry = selectedElementRegistry.get(selectionKey);
    if (!entry) return;
    const { id, el } = entry;
    applySelectionStyles(el, false);
    if (el?.classList) el.classList.remove('drag-preview');
    selectedElementRegistry.delete(selectionKey);
    if (id != null) {
      const normalizedId = String(id);
      if (scheduleSelectionTargets.get(normalizedId) === selectionKey) {
        scheduleSelectionTargets.delete(normalizedId);
        if (removeId) selectedSchedules.delete(normalizedId);
      }
    }
  }

  function registerSelectionElement(el, id) {
    if (!el) return null;
    const normalizedId = id != null ? String(id) : null;
    const selectionKey = ensureSelectionKey(el);
    if (!selectionKey) return null;

    if (normalizedId) {
      const existingKey = scheduleSelectionTargets.get(normalizedId);
      if (existingKey && existingKey !== selectionKey) unregisterSelectionByKey(existingKey);
      scheduleSelectionTargets.set(normalizedId, selectionKey);
      selectedSchedules.add(normalizedId);
    }

    selectedElementRegistry.set(selectionKey, { id: normalizedId, el });
    applySelectionStyles(el, true);
    return selectionKey;
  }

  function getScheduleElementsById(id, root = document) {
    const escaped = cssEscape(String(id));
    return qAll(`.schedule-pill[data-id="${escaped}"], .fc-event-pill[data-id="${escaped}"]`, root);
  }

  function findCalendarEventByScheduleId(id) {
    if (!calendar) return null;
    const events = calendar.getEvents();
    return events.find(ev => String(ev.extendedProps?.id) === String(id) || String(ev.id) === String(id)) || null;
  }

  function highlightSelectionByIds(ids, attempt = 0) {
    const missing = [];
    ids.forEach(id => {
      const normalizedId = String(id);
      const existingKey = scheduleSelectionTargets.get(normalizedId);
      const existingEntry = existingKey ? selectedElementRegistry.get(existingKey) : null;
      if (existingEntry?.el && document.contains(existingEntry.el)) { applySelectionStyles(existingEntry.el, true); return; }
      const els = getScheduleElementsById(normalizedId);
      if (!els.length) { missing.push(normalizedId); return; }
      registerSelectionElement(els[0], normalizedId);
    });
    if (missing.length && attempt < 5) {
      const retry = () => highlightSelectionByIds(missing, attempt + 1);
      if (typeof requestAnimationFrame === 'function') requestAnimationFrame(retry); else setTimeout(retry, 16);
    }
  }

  function applyDragPreviewToSelection(active) {
    const method = active ? 'add' : 'remove';
    selectedElementRegistry.forEach(({ el }) => {
      if (!el || !el.classList) return;
      if (!document.contains(el)) return;
      el.classList[method]('drag-preview');
    });
    if (!active) document.body.classList.remove('multi-dragging');
  }

  function queueSelectionReset(ids) {
    const normalized = Array.from(new Set(ids.map(id => String(id))));
    clearScheduleSelection();
    normalized.forEach(id => selectedSchedules.add(id));
    if (normalized.length) {
      const highlight = () => highlightSelectionByIds(normalized);
      if (typeof requestAnimationFrame === 'function') requestAnimationFrame(highlight); else setTimeout(highlight, 16);
    }
  }

  function showToastWrapper(msg, type = 'info') { showToast?.(msg, type); }

  function clearScheduleSelection() {
    Array.from(selectedElementRegistry.keys()).forEach(key => unregisterSelectionByKey(key));
    qAll('.schedule-pill.selected, .fc-event-pill.selected, .schedule-pill.drag-preview, .fc-event-pill.drag-preview')
      .forEach(el => el.classList.remove('selected','ring-2','ring-blue-500','shadow-lg','drag-preview'));
    selectedElementRegistry.clear();
    scheduleSelectionTargets.clear();
    selectedSchedules.clear();
  }

  function clearTargetDateSelection() {
    selectedTargetDates.clear();
    qAll('.fc-daygrid-day.selected-date, .fc-timegrid-slot.selected-date')
      .forEach(el => el.classList.remove('selected-date'));
  }

  function toggleScheduleSelection(el, keepOthers = false) {
    if (!el) return;
    const rawId = getScheduleIdFromElement(el) || el.dataset?.id || null;
    const normalizedId = rawId != null ? String(rawId) : null;
    const selectionKey = ensureSelectionKey(el);
    if (!selectionKey) return;
    const wasSelected = selectedElementRegistry.has(selectionKey);

    if (!keepOthers) {
      if (wasSelected) { clearScheduleSelection(); return; }
      clearScheduleSelection();
    }
    if (keepOthers && wasSelected) { unregisterSelectionByKey(selectionKey); return; }
    registerSelectionElement(el, normalizedId);
    if (!document.contains(el) && normalizedId != null) highlightSelectionByIds([normalizedId]);
  }

  function toggleTargetDateSelection(el, keepOthers = false) {
    if (!el) return;
    const dateStr = el.getAttribute('data-date'); if (!dateStr) return;
    if (!keepOthers) clearTargetDateSelection();
    if (selectedTargetDates.has(dateStr)) { selectedTargetDates.delete(dateStr); el.classList.remove('selected-date'); }
    else { selectedTargetDates.add(dateStr); el.classList.add('selected-date'); }
  }

  document.addEventListener('contextmenu', e => {
    const pill = e.target.closest('.schedule-pill');
    const dateEl = e.target.closest('.fc-daygrid-day, .fc-timegrid-slot');
    const keepOthers = hasMultiSelectModifier(e);

    if (pill) {
      if (keepOthers) toggleScheduleSelection(pill, true);
      else if (!pill.classList.contains('selected')) toggleScheduleSelection(pill, false);
    }

    e.preventDefault();
    q('#context-menu')?.remove();

    const x = e.pageX, y = e.pageY;
    const menu = document.createElement('div');
    menu.id = 'context-menu';
    menu.className = 'absolute bg-white border border-gray-300 rounded shadow-lg z-50';
    menu.style.left = `${x}px`; menu.style.top = `${y}px`; menu.style.minWidth = '180px';

    const scheduleId = pill ? getScheduleIdFromElement(pill) : null;
    const contextEvent = scheduleId ? findCalendarEventByScheduleId(String(scheduleId)) : null;

    if (contextEvent) {
      const editBtn = document.createElement('button');
      editBtn.className = 'block w-full text-left px-4 py-2 hover:bg-gray-100';
      editBtn.innerText = '✏️ Edit Shift Code';
      editBtn.onclick = () => {
        currentDroppingEvent = contextEvent;
        isEditingExistingEvent = true;
        const existingCode = contextEvent.extendedProps?.shiftCode ?? '';
        openShiftModal({ initialShiftCode: existingCode, useLastUsedWhenEmpty: false });
        q('#context-menu')?.remove();
      };
      menu.appendChild(editBtn);
    }

    if (selectedSchedules.size) {
      const btn = document.createElement('button');
      btn.className = 'block w-full text-left px-4 py-2 hover:bg-gray-100';
      btn.innerText = `📋 Copy Selected Schedule${selectedSchedules.size>1?'s':''}`;
      btn.onclick = () => { copySelectedSchedules(); q('#context-menu')?.remove(); };
      menu.appendChild(btn);
    }

    if (dateEl && copiedSchedules.length) {
      const targetStr = dateEl.getAttribute('data-date');
      const pasteBtn = document.createElement('button');
      pasteBtn.className = 'block w-full text-left px-4 py-2 hover:bg-gray-100';
      pasteBtn.innerText = `📅 Paste Here (${targetStr})`;
      pasteBtn.onclick = () => { pasteSchedulesToDates([targetStr]); q('#context-menu')?.remove(); };
      menu.appendChild(pasteBtn);

      if (selectedTargetDates.size) {
        const pasteMultiBtn = document.createElement('button');
        pasteMultiBtn.className = 'block w-full text-left px-4 py-2 hover:bg-gray-100';
        pasteMultiBtn.innerText = `📅 Paste to ${selectedTargetDates.size} selected date(s)`;
        pasteMultiBtn.onclick = () => { pasteSchedulesToDates(Array.from(selectedTargetDates)); q('#context-menu')?.remove(); };
        menu.appendChild(pasteMultiBtn);
      }
    }

    if (!menu.hasChildNodes()) {
      const hint = document.createElement('div');
      hint.className = 'px-4 py-2 text-sm text-gray-500';
      hint.innerText = 'No actions available';
      menu.appendChild(hint);
    }

    document.body.appendChild(menu);
    document.addEventListener('click', () => q('#context-menu')?.remove(), { once: true });
    document.addEventListener('keydown', ev => { if (ev.key === 'Escape') q('#context-menu')?.remove(); }, { once: true });
  });

  document.addEventListener('keydown', e => {
    const tag = e.target?.tagName?.toLowerCase();
    if (tag === 'input' || tag === 'textarea' || e.target?.isContentEditable) return;
    if (!hasMultiSelectModifier(e) || e.repeat) return;
    const key = (e.key || '').toLowerCase();
    if (key === 'c') { e.preventDefault(); copySelectedSchedules(); }
    if (key === 'v') {
      e.preventDefault();
      const el = document.elementFromPoint(lastMouseX, lastMouseY);
      const dateEl = el?.closest('.fc-daygrid-day, .fc-timegrid-slot');
      const dateStr = dateEl?.getAttribute('data-date') || null;
      const hasTargetSelection = selectedTargetDates.size > 0;
      if (dateStr && (!hasTargetSelection || !selectedTargetDates.has(dateStr))) pasteSchedulesToDates([dateStr]);
      else if (selectedTargetDates.size > 1) pasteSchedulesToDates(Array.from(selectedTargetDates));
      else if (selectedTargetDates.size === 1) { const [single] = Array.from(selectedTargetDates); pasteSchedulesToDates([single]); }
      else if (dateStr) pasteSchedulesToDates([dateStr]);
      else showToastWrapper('Hover over a date or select target dates to paste.', 'warn');
    }
    if (key === 'z') { e.preventDefault(); undoLastPaste(); }
  });

  function copySelectedSchedules() {
    const pointerPill = document.elementFromPoint(lastMouseX, lastMouseY)?.closest('.schedule-pill');
    const pointerId = getScheduleIdFromElement(pointerPill);
    const effectiveHoverId = pointerId || hoveredScheduleId;
    let idsToCopy = [];
    if (selectedSchedules.size > 1) idsToCopy = Array.from(selectedSchedules);
    else if (pointerId) idsToCopy = [pointerId];
    else if (selectedSchedules.size === 1) idsToCopy = Array.from(selectedSchedules);
    else if (effectiveHoverId) idsToCopy = [effectiveHoverId];

    if (!idsToCopy.length) return showToastWrapper('No schedule selected.', 'warn');
    if (!calendar) return showToastWrapper('Calendar not ready.', 'error');
    const allEvents = calendar.getEvents();
    copiedSchedules = idsToCopy.map(id =>
      allEvents.find(e => String(e.extendedProps?.id) === String(id) || String(e.id) === String(id))
    ).filter(Boolean).map(ev => ({
      empNo: ev.extendedProps.empNo,
      shiftCode: ev.extendedProps.shiftCode ?? null,
      type: ev.extendedProps.type ?? null,
      date: ev.startStr
    }));
    showToastWrapper(`Copied ${copiedSchedules.length} schedule${copiedSchedules.length > 1 ? 's' : ''}.`, 'info');
  }

  function pasteSchedulesToDates(datesArray) {
    if (!copiedSchedules.length) return showToastWrapper('Nothing copied.', 'warn');
    if (!calendar) return showToastWrapper('Calendar not ready.', 'error');

    const uniqueDates = Array.from(new Set(datesArray)).sort();
    const baseDate = new Date(copiedSchedules[0].date);
    let added = 0, skipped = 0;
    const createdIdsForBatch = [];

    uniqueDates.forEach(targetStr => {
      const targetDate = new Date(targetStr);
      const offsetDays = Math.round((targetDate - baseDate) / (1000*60*60*24));
      copiedSchedules.forEach(src => {
        const newDate = new Date(src.date);
        newDate.setDate(newDate.getDate() + offsetDays);
        const newDateStr = newDate.toISOString().split('T')[0];
        const conflict = findExistingShiftForEmployee(src.empNo, newDateStr);
        if (conflict) { skipped++; return; }
        const newId = crypto?.randomUUID?.() ?? `sched-${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
        const pasteType = src.type || 'work';
        const color = ensureEmployeeColor(src.empNo);
        const gradient = getGradientFromBaseColor(color, pasteType === 'rest' ? 'rest' : 'work');
        const forceStyle = pasteType === 'rest'
          ? { background: gradient, backgroundImage: 'none', color, border: `2px solid ${color}` }
          : { background: gradient, backgroundImage: 'none', color: '#fff', border: 'none' };
        const classNames = ['fc-event-pill', pasteType === 'rest' ? 'fc-event-rest' : 'fc-event-work'];
        const newEvent = calendar.addEvent({
          id: newId, title: employees?.[src.empNo]?.name ?? src.empNo, start: newDateStr, classNames,
          extendedProps: { id: newId, empNo: src.empNo, shiftCode: src.shiftCode, type: pasteType, forceStyle }
        });
        if (newEvent) {
          try {
            newEvent.setProp('classNames', classNames);
            newEvent.setExtendedProp('forceStyle', forceStyle);
            newEvent.setProp('backgroundColor','');
            newEvent.setProp('borderColor', pasteType === 'rest' ? color : 'transparent');
            newEvent.setProp('textColor', pasteType === 'rest' ? color : '#fff');
          } catch (e) {}
          const createdId = newEvent.extendedProps?.id || newEvent.id || newId;
          createdIdsForBatch.push(createdId);
          decorateEventLater(newEvent);
        }
        added++;
      });
    });

    clearScheduleSelection();
    if (createdIdsForBatch.length) { pasteHistory.push(createdIdsForBatch); if (pasteHistory.length > 20) pasteHistory.shift(); }
    runConflictDetection(); updateStats(); saveToLocalStorage();
    showToastWrapper(`Pasted ${added} schedule${added !== 1 ? 's' : ''} (${skipped} skipped).`, added ? 'success' : 'warn');
    clearTargetDateSelection();
  }

  function undoLastPaste() {
    if (!calendar) return showToastWrapper('Calendar not ready.', 'error');
    if (!pasteHistory.length) return showToastWrapper('Nothing to undo.', 'warn');

    const lastBatch = pasteHistory.pop();
    let removed = 0;
    const events = calendar.getEvents();
    lastBatch.forEach(id => {
      const match = events.find(e => String(e.extendedProps?.id) === String(id) || String(e.id) === String(id));
      if (match) { match.remove(); removed++; }
    });
    if (!removed) { showToastWrapper('Nothing to undo.', 'warn'); return; }
    runConflictDetection(); updateStats(); saveToLocalStorage();
    showToastWrapper(`Undid paste (${removed} schedule${removed !== 1 ? 's' : ''} removed).`, 'info');
  }

  // Mouse tracking for copy/paste hover
  const updateMouse = e => {
    lastMouseX = e.clientX; lastMouseY = e.clientY;
    const pill = e.target.closest?.('.schedule-pill');
    if (pill) {
      const id = getScheduleIdFromElement(pill);
      if (id) hoveredScheduleId = id;
    } else if (!document.elementFromPoint(lastMouseX,lastMouseY)?.closest('.schedule-pill')) {
      hoveredScheduleId = null;
    }
  };
  document.addEventListener('mousemove', updateMouse);

  function clearTargetDateSelectionIfNeeded() { /* helper noop */ }

  // DRAG SELECT start
  document.addEventListener('mousedown', e => {
    if (e.button !== 0) return;
    updateMouse(e);
    const pill = e.target.closest('.schedule-pill');
    const day  = e.target.closest('.fc-daygrid-day, .fc-timegrid-slot');

    if (pill) {
      const pillId = getScheduleIdFromElement(pill) || pill.dataset.id;
      const normalizedId = pillId ? String(pillId) : null;
      const isModifier = hasMultiSelectModifier(e);
      if (isModifier) return;

      const canDragGroup = normalizedId && selectedSchedules.size > 1 && selectedSchedules.has(normalizedId);
      if (canDragGroup) {
        e.preventDefault(); e.stopPropagation();
        buildCopiedFromSelected({ silent: true });
        dragSelectionEvents = Array.from(selectedSchedules).map(id => findCalendarEventByScheduleId(String(id))).filter(Boolean);
        const anchorEvent = findCalendarEventByScheduleId(normalizedId) || dragSelectionEvents[0] || null;
        dragAnchorDate = anchorEvent ? anchorEvent.startStr : null;
        if (!dragSelectionEvents.length || !dragAnchorDate) { dragSelectionEvents = []; dragAnchorDate = null; return; }
        isDragging = true; window.__multiSelectDragActive = true;
        createDragGhost(selectedSchedules.size, e.clientX, e.clientY);
        document.body.classList.add('no-select', 'no-transform-during-drag'); // FIX: lock transforms during ghost drag
        return;
      }
    }

    if (day) {
      isSelectingDates = true; dateSelectStartEl = day;
      if (!hasMultiSelectModifier(e)) clearTargetDateSelection();
      toggleTargetDateSelection(day, true);
      document.body.classList.add('no-select');
    }
  });

  // FIX: ghost follows viewport coords exactly
  document.addEventListener('mousemove', e => {
    if (isDragging && dragGhost) {
      dragGhost.style.left = `${e.clientX}px`;
      dragGhost.style.top  = `${e.clientY}px`;
      dragGhost.style.transform = 'translate(-50%, -50%)';
    } else if (isSelectingDates && dateSelectStartEl) {
      const currentEl = document.elementFromPoint(e.clientX, e.clientY);
      if (currentEl && currentEl.closest('.fc-daygrid-day')) {
        const currentDayEl = currentEl.closest('.fc-daygrid-day');
        const allDays = Array.from(document.querySelectorAll('.fc-daygrid-day'));
        const startIdx = allDays.indexOf(dateSelectStartEl);
        const endIdx   = allDays.indexOf(currentDayEl);
        allDays.forEach(day => day.classList.remove('range-selecting'));
        if (startIdx !== -1 && endIdx !== -1) {
          const [minIdx, maxIdx] = [Math.min(startIdx, endIdx), Math.max(startIdx, endIdx)];
          for (let i = minIdx; i <= maxIdx; i++) allDays[i].classList.add('range-selecting');
        }
      }
    }
  });

  document.addEventListener('mouseup', e => {
    if (isDragging) {
      const el = document.elementFromPoint(e.clientX, e.clientY);
      const dateEl = el?.closest('.fc-daygrid-day, .fc-timegrid-slot');
      const targetDateStr = dateEl?.getAttribute('data-date');
      let moved = false, attempted = false;
      if (targetDateStr) { attempted = true; moved = moveDraggedSchedules(targetDateStr); }
      else if (selectedTargetDates.size) {
        const [singleTarget] = Array.from(selectedTargetDates);
        if (singleTarget) { attempted = true; moved = moveDraggedSchedules(singleTarget); }
      }
      if (!moved && !attempted) showToastWrapper('Drop target not valid.', 'warn');
      removeDragGhost();
      isDragging = false;
      window.__multiSelectDragActive = false;
      document.body.classList.remove('no-select', 'no-transform-during-drag'); // FIX: cleanup
      buildCopiedFromSelected({ silent: true });
      return;
    }
    if (isSelectingDates) { isSelectingDates = false; dateSelectStartEl = null; document.body.classList.remove('no-select'); }
  });

  function moveDraggedSchedules(targetDateStr) {
    if (!calendar) return false;
    if (!dragSelectionEvents.length || !dragAnchorDate) return false;
    const targetDate = new Date(targetDateStr); const anchorDate = new Date(dragAnchorDate);
    if (Number.isNaN(targetDate.getTime())) { showToastWrapper('Drop target not valid.', 'warn'); return false; }
    if (Number.isNaN(anchorDate.getTime())) return false;

    const msPerDay = 24*60*60*1000;
    const offsetDays = Math.round((targetDate - anchorDate) / msPerDay);

    const plannedMoves = dragSelectionEvents.map(event => {
      const original = new Date(event.startStr);
      original.setDate(original.getDate() + offsetDays);
      const newDateStr = original.toISOString().split('T')[0];
      return { event, newDateStr };
    });

    const eventsToIgnore = new Set(dragSelectionEvents);
    const ignoredEvents = Array.from(eventsToIgnore);
    const hasConflict = plannedMoves.some(({ event, newDateStr }) => {
      const empNo = event.extendedProps?.empNo; if (!empNo) return false;
      const conflict = findExistingShiftForEmployee(empNo, newDateStr, { exclude: [...ignoredEvents, event] });
      return !!conflict;
    });
    if (hasConflict) { showToastWrapper('Move cancelled: a selected employee already has a shift on that date.', 'error'); return false; }

    plannedMoves.forEach(({ event, newDateStr }) => {
      event.setStart(newDateStr); try { event.setEnd(newDateStr); } catch (e) {}
      decorateEventLater(event);
    });

    queueSelectionReset(dragSelectionEvents.map(ev => ev.extendedProps?.id || ev.id));
    clearTargetDateSelection();
    dragAnchorDate = targetDateStr;
    runConflictDetection(); updateStats(); saveToLocalStorage();
    showToastWrapper(`Moved ${plannedMoves.length} schedule${plannedMoves.length !== 1 ? 's' : ''}.`, 'success');
    return true;
  }

  function createDragGhost(count, x = null, y = null) {
    removeDragGhost({ keepPreview: true });
    dragGhost = document.createElement('div');
    dragGhost.className = 'drag-ghost fixed z-50 p-2 rounded border border-gray-300 bg-white shadow';
    dragGhost.style.pointerEvents = 'none';
    const left = (typeof x === 'number') ? x : lastMouseX;
    const top  = (typeof y === 'number') ? y : lastMouseY;
    dragGhost.style.left = `${left}px`;
    dragGhost.style.top = `${top}px`;
    dragGhost.style.transform = 'translate(-50%,-50%)';
    dragGhost.style.minWidth = '120px';
    dragGhost.innerHTML = `
      <svg class="drag-ghost-icon" viewBox="0 0 20 20" fill="currentColor" aria-hidden="true">
        <path d="M7 4a1 1 0 110-2 1 1 0 010 2zm6-2a1 1 0 100 2 1 1 0 000-2zM7 11a1 1 0 110-2 1 1 0 010 2zm6-2a1 1 0 100 2 1 1 0 000-2zM7 18a1 1 0 110-2 1 1 0 010 2zm6-2a1 1 0 100 2 1 1 0 000-2z"/>
      </svg>
      <strong>${count}</strong> schedule${count>1?'s':''}
      <span class="drag-ghost-hint">Release to drop</span>
    `.trim();
    document.body.appendChild(dragGhost);
    document.body.classList.add('multi-dragging');
    applyDragPreviewToSelection(true);
  }

  function removeDragGhost(options = {}) {
    const { keepPreview = false } = options;
    if (dragGhost) { dragGhost.remove(); dragGhost = null; }
    if (!keepPreview) applyDragPreviewToSelection(false);
    dragSelectionEvents = []; dragAnchorDate = null;
  }

  function buildCopiedFromSelected(options = {}) {
    if (!calendar) return;
    const { silent = false } = options;
    const allEvents = calendar.getEvents();
    copiedSchedules = Array.from(selectedSchedules).map(id =>
      allEvents.find(e => String(e.extendedProps?.id) === String(id) || String(e.id) === String(id))
    ).filter(Boolean).map(ev => ({
      empNo: ev.extendedProps.empNo,
      shiftCode: ev.extendedProps?.shiftCode ?? null,
      type: ev.extendedProps?.type ?? null,
      date: ev.startStr
    }));
  }

  /* ===========================
     GLOBAL LISTENERS / STARTUP
  =========================== */
  const addRowBtnEl = document.getElementById('add-row-btn');
  if (addRowBtnEl) addRowBtnEl.addEventListener('click', addEmployeeRow);
  document.addEventListener('click', function (e) {
    const btn = e.target.closest && e.target.closest('#add-row-btn');
    if (btn) { try { addEmployeeRow(); } catch (err) { console.error('addEmployeeRow error (fallback):', err); } }
  });
  try { window.addEmployeeRow = addEmployeeRow; } catch(e){}

  const saveGenerateBtnEl = document.getElementById('save-generate-btn');
  if (saveGenerateBtnEl) saveGenerateBtnEl.addEventListener('click', saveAndGenerate);
  if (employeeTableBody) {
    employeeTableBody.addEventListener('click', function(e) {
      const removeBtn = e.target.closest('.remove-row-btn');
      if (removeBtn) removeEmployeeRow(removeBtn);
    });
  }

  const exportBtnEl = document.getElementById('export-btn');
  if (exportBtnEl) exportBtnEl.addEventListener('click', openExportModal);
  const resetBtnEl = document.getElementById('reset-btn');
  if (resetBtnEl) resetBtnEl.addEventListener('click', resetAll);

  if (saveShiftBtn) saveShiftBtn.addEventListener('click', handleSaveShift);
  if (confirmExportBtn) confirmExportBtn.addEventListener('click', exportToExcel);
  if (confirmDeleteBtn) confirmDeleteBtn.addEventListener('click', handleConfirmDelete);

  document.querySelectorAll('[data-modal-close]').forEach(btn => { btn.addEventListener('click', () => handleCloseModal(btn.dataset.modalClose)); });
  document.querySelectorAll('.modal-backdrop').forEach(backdrop => {
    backdrop.addEventListener('click', function(e) { if (e.target === this) handleCloseModal(this.id); });
  });

  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      if (shiftModal && !shiftModal.classList.contains('hidden')) handleCloseModal('shift-modal');
      if (exportModal && !exportModal.classList.contains('hidden')) handleCloseModal('export-modal');
      if (deleteModal && !deleteModal.classList.contains('hidden')) handleCloseModal('delete-modal');
    }
  });

  function startScheduler() {
    try { initializeCalendar(); } catch (err) { console.error('initializeCalendar failed', err); }
    try { loadFromLocalStorage(); } catch (err) { console.error('loadFromLocalStorage failed', err); }
    try { if (!employeeTableBody || employeeTableBody.querySelectorAll('tr').length === 0) { addEmployeeRow(); } } catch (err) {}
    try { decorateCalendarEvents(); calendar?.on('eventDidMount', decorateCalendarEvents); } catch (err) { console.error('decorateCalendarEvents failed', err); }
  }

  startScheduler();
});
