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
  let pendingImportRows = [];
  let importMode = 'preview';
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

  const gradientCache = new Map();

  function getGradientFromBaseColor(hex, type = 'work') {
    if (!hex || hex[0] !== '#') return '';
    const cacheKey = `${hex}|${type}`;
    const cached = gradientCache.get(cacheKey);
    if (cached) return cached;

    const rgb = parseInt(hex.slice(1), 16);
    const r = (rgb >> 16) & 255, g = (rgb >> 8) & 255, b = rgb & 255;
    const [h, s, l] = rgbToHsl(r, g, b);
    const gradient = (type === 'work')
      ? `linear-gradient(135deg, hsl(${h}, ${Math.round(s*100)}%, ${Math.max(25, Math.round(l*100)-10)}%), hsl(${h}, ${Math.round(s*100)}%, ${Math.min(70, Math.round(l*100)+10)}%))`
      : `linear-gradient(135deg, hsl(${h}, ${Math.round(s*100)}%, ${Math.min(90, Math.round(l*100)+20)}%), hsl(${h}, ${Math.round(s*100)}%, ${Math.min(95, Math.round(l*100)+25)}%))`;

    gradientCache.set(cacheKey, gradient);
    return gradient;
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
  const importModal = document.getElementById('import-modal');

  const shiftPresetSelect = document.getElementById('shift-preset');
  const shiftCustomInput = document.getElementById('shift-custom');
  const saveShiftBtn = document.getElementById('save-shift-btn');
  const confirmExportBtn = document.getElementById('confirm-export-btn');
  const confirmDeleteBtn = document.getElementById('confirm-delete-btn');
  const deleteModalSummary = document.getElementById('delete-modal-summary');
  const importFileInput = document.getElementById('import-file-input');
  const importPreviewBody = document.getElementById('import-preview-body');
  const importPreviewNote = document.getElementById('import-preview-note');
  const importSummaryPanel = document.getElementById('import-summary-panel');
  const importSummaryCounts = document.getElementById('import-summary-counts');
  const importSummaryList = document.getElementById('import-summary-list');
  const confirmImportBtn = document.getElementById('confirm-import-btn');
  const cancelImportBtn = document.getElementById('cancel-import-btn');

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

  const tooltipCache = new WeakMap();
  let tooltipDelegateInstance = null;

  const isDragPerfSuppressed = () => !!window.__globalDragActive;

  function ensureTooltipDelegate() {
    if (typeof tippy === 'undefined') return;
    if (tooltipDelegateInstance) return;
    tooltipDelegateInstance = tippy.delegate(document.body, {
      target: '.schedule-pill[data-tooltip-content]',
      allowHTML: true,
      theme: 'light-border',
      placement: 'top',
      delay: [150, 50],
      onShow(instance) {
        const content = instance?.reference?.dataset?.tooltipContent;
        if (content && instance.props.content !== content) instance.setContent(content);
      }
    });
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

  function getTooltipCacheKey(event) {
    if (!event) return '';
    const ext = event.extendedProps || {};
    const empNo = ext.empNo ?? '';
    const type = ext.type ?? '';
    const shiftCode = ext.shiftCode ?? '';
    const title = event.title ?? '';
    const employee = employees?.[empNo] || {};
    const position = employee.position || ext.position || '';
    const shiftTime = shiftCode && shiftTimes[shiftCode] ? shiftTimes[shiftCode] : '';
    return [empNo, type, shiftCode, title, position, shiftTime].join('|');
  }

  function getCachedTooltipContent(event) {
    const key = getTooltipCacheKey(event);
    const cached = tooltipCache.get(event);
    if (cached && cached.key === key) return cached.content;

    const content = buildEventTooltipContent(event);
    tooltipCache.set(event, { key, content });
    return content;
  }

  function refreshEventTooltip(event, el) {
    if (!event || typeof tippy === 'undefined' || isDragPerfSuppressed()) return;
    const target = el || (typeof findEventElementByEvent === 'function' ? findEventElementByEvent(event) : null);
    if (!target) return;
    const content = getCachedTooltipContent(event);
    try {
      if (target._tippy && target._tippy.destroy) { target._tippy.destroy(); }
      if (target.dataset.tooltipContent !== content) target.dataset.tooltipContent = content;
      target.setAttribute('data-tooltip-content', content || '');
      ensureTooltipDelegate();
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
  const positionOptions = ['Branch Head','Site Supervisor','OIC','Mac Expert','Cashier','Inventory Analyst','Engineer','Parts Management Analyst','Customer Service Officer','Service Specialist','Parts Management Specialist','Customer Care Representative','Asst. Supervisor','Trainer','Customer Relations Officer'];
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
  let calendar;
  let __sidebarDragListenersWired = false;
  let __calendarDragLiftWired = false;
  let __bodyDragGuardsWired = false;
  let __isSidebarDragging = false;
  const wiredSidebarCards = new WeakSet();
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
  let calendarDragGhost = null;
  let calendarDragMoveHandler = null;
  let sidebarDragImageEl = null;
  let calendarMaintenanceTimer = null;
  const setGlobalDragFlag = active => { window.__globalDragActive = !!active; };
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

  function runWithCalendarBatch(callback) {
    if (calendar && typeof calendar.batchRendering === 'function') {
      calendar.batchRendering(callback);
    } else {
      callback();
    }
  }

  function runCalendarMaintenance(options = {}) {
    const { conflicts = true, stats = true, save = true } = options;
    if (conflicts) {
      runConflictDetection();
    } else if (stats) {
      updateStats();
    }
    if (save) saveToLocalStorage();
  }

  function scheduleCalendarMaintenance(options = {}) {
    const {
      delay = 40,
      conflicts = true,
      stats = true,
      save = true
    } = options;

    if (calendarMaintenanceTimer) clearTimeout(calendarMaintenanceTimer);
    calendarMaintenanceTimer = setTimeout(() => {
      calendarMaintenanceTimer = null;
      const runner = () => runCalendarMaintenance({ conflicts, stats, save });
      if (!conflicts && typeof requestIdleCallback === 'function') {
        requestIdleCallback(runner, { timeout: 600 });
      } else {
        runner();
      }
    }, delay);
  }

  function ensureConflictPlaceholderRow() {
    if (!conflictTableBody) return null;
    let row = document.getElementById('conflicts-placeholder');
    if (!row) {
      row = document.createElement('tr');
      row.id = 'conflicts-placeholder';
      row.innerHTML = `
        <td colspan="4" class="p-6 text-center text-gray-500">
          No conflicts detected.
        </td>
      `;
    }
    return row;
  }

  function setEventConflictFlag(event, active) {
    if (!event) return;
    if (event.extendedProps) event.extendedProps.isConflict = !!active;
    const el = findEventElementByEvent(event);
    if (el) el.classList.toggle('fc-event-conflict', !!active);
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
                  <div class='fc-event-pill fc-event-work px-3 py-1 text-xs font-medium rounded-full' data-event='${JSON.stringify(workEventData)}' style="${workStyle}">Work</div>
                  <div class='fc-event-pill fc-event-rest px-3 py-1 text-xs font-medium rounded-full' data-event='${JSON.stringify(restEventData)}' style="${restStyle}">Rest</div>
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
        scheduleCalendarMaintenance({ delay: 40, save: false });
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

  function moveCalendarEventDragPreview(x, y) {
    if (!calendarDragGhost || typeof x !== 'number' || typeof y !== 'number') return;
    calendarDragGhost.style.left = `${x}px`;
    calendarDragGhost.style.top = `${y}px`;
  }

  function stopCalendarEventDragPreview() {
    if (calendarDragMoveHandler) {
      document.removeEventListener('mousemove', calendarDragMoveHandler, true);
      document.removeEventListener('pointermove', calendarDragMoveHandler, true);
      calendarDragMoveHandler = null;
    }
    if (calendarDragGhost) {
      calendarDragGhost.remove();
      calendarDragGhost = null;
    }
    document.body.classList.remove('calendar-event-dragging');
  }

  function startCalendarEventDragPreview(info) {
    if (!info || !info.event || window.__multiSelectDragActive) return;
    stopCalendarEventDragPreview();

    const event = info.event;
    const type = event.extendedProps?.type === 'rest' ? 'rest' : 'work';
    const shiftCode = event.extendedProps?.shiftCode || '';
    const empNo = event.extendedProps?.empNo;
    const color = empNo ? ensureEmployeeColor(empNo) : null;
    const ghost = document.createElement('div');
    ghost.className = `calendar-drag-ghost fc-event-pill schedule-pill fc-event-${type}`;
    ghost.innerHTML = `<div class="pill-content"><span class="pill-name">${escapeHtml(event.title || '')}</span>${shiftCode ? `<span class="pill-shift">${escapeHtml(shiftCode)}</span>` : ''}</div>`;

    if (color) {
      const gradient = getGradientFromBaseColor(color, type);
      if (gradient) ghost.style.setProperty('background', gradient, 'important');
      if (type === 'rest') {
        ghost.style.setProperty('color', color, 'important');
        ghost.style.setProperty('border', `2px solid ${color}`, 'important');
      } else {
        ghost.style.setProperty('color', '#fff', 'important');
        ghost.style.setProperty('border', 'none', 'important');
      }
    }

    document.body.appendChild(ghost);
    calendarDragGhost = ghost;
    document.body.classList.add('calendar-event-dragging', 'no-transform-during-drag');

    const startX = info.jsEvent?.clientX ?? lastMouseX;
    const startY = info.jsEvent?.clientY ?? lastMouseY;
    moveCalendarEventDragPreview(startX, startY);

    calendarDragMoveHandler = ev => {
      lastMouseX = ev.clientX;
      lastMouseY = ev.clientY;
      moveCalendarEventDragPreview(ev.clientX, ev.clientY);
    };
    document.addEventListener('mousemove', calendarDragMoveHandler, true);
    document.addEventListener('pointermove', calendarDragMoveHandler, true);
  }

/* ===========================
   CALENDAR INIT (improved)
=========================== */
function initializeCalendar() {
  if (!calendarEl) return;

  // Make sidebar items draggable for FullCalendar
  const draggableContainer = document.getElementById('draggable-cards-container');
  if (draggableContainer && FullCalendar && FullCalendar.Draggable) {
    new FullCalendar.Draggable(draggableContainer, {
      itemSelector: '.fc-event-pill, .draggable-pill',
      eventData: (el) => {
        const ds = el.dataset || {};
        const titleFromDom =
          ds.title ||
          el.querySelector('.pill-name')?.textContent?.trim() ||
          el.textContent?.trim() ||
          'Shift';

        const pillClasses = (el.className || '')
          .split(/\s+/)
          .filter(Boolean);

        return {
          title: titleFromDom,
          allDay: true,
          classNames: Array.from(new Set(['fc-event-pill', ...pillClasses])),
          extendedProps: {
            type: ds.type || 'work',
            empNo: ds.empNo || ds.empno || ds.employeeNo || null,
            employeeNo: ds.employeeNo || ds.empNo || ds.empno || null,
            position: ds.position || null,
          },
        };
      },
    });
  }

  // Build the calendar
  calendar = new FullCalendar.Calendar(calendarEl, {
    initialView: 'dayGridMonth',
    headerToolbar: {
      left: 'prev,next today',
      center: 'title',
      right: 'dayGridMonth,dayGridWeek,dayGridDay',
    },

    // Core interaction settings
    editable: true,
    droppable: true,
    dropAccept: '.fc-event-pill, .draggable-pill',
    eventStartEditable: true,
    eventDurationEditable: false,
    dragScroll: false,
    selectable: true,
    dayMaxEvents: true,
    progressiveEventRendering: true,
    eventOrderStrict: false,
    height: 'auto',

    eventDragStart(info) {
      setGlobalDragFlag(true);
      document.body.classList.add('no-transform-during-drag');
      startCalendarEventDragPreview(info);
    },
eventDragStop(info)  {
      setGlobalDragFlag(false);
      document.body.classList.remove('no-transform-during-drag');
      stopCalendarEventDragPreview();
    },

eventResizeStart(info) { setGlobalDragFlag(true); document.body.classList.add('no-transform-during-drag'); },
eventResizeStop(info)  { setGlobalDragFlag(false); document.body.classList.remove('no-transform-during-drag'); },

    /* -------------------------
       External item received
    ------------------------- */
    eventReceive(info) {
      try {
        if (typeof selectedSchedules !== 'undefined' && selectedSchedules?.clear) selectedSchedules.clear();
        if (typeof selectedTargetDates !== 'undefined' && selectedTargetDates?.clear) selectedTargetDates.clear();
        window.__multiSelectDragActive = false;
        document.body.classList.remove('multi-dragging');
      } catch (e) {}

      const newEvent = info.event;
      try {
        const nowId = `evt_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
        const ext = JSON.parse(JSON.stringify(newEvent.extendedProps || {}));

        if (!ext.id && !newEvent.id) {
          ext.id = nowId;
          newEvent.setExtendedProp('id', nowId);
          try { newEvent.setProp('id', nowId); } catch (_) {}
        }

        if (ext.empNo == null && newEvent.extendedProps?.empNo != null) {
          newEvent.setExtendedProp('empNo', String(newEvent.extendedProps.empNo));
        }
        if (ext.employeeNo == null && newEvent.extendedProps?.employeeNo != null) {
          newEvent.setExtendedProp('employeeNo', String(newEvent.extendedProps.employeeNo));
        }

        const safeType = ext.type || newEvent.extendedProps?.type || 'work';
        newEvent.setExtendedProp('type', safeType);

        if (newEvent.allDay !== true) {
          try { newEvent.setAllDay(true); } catch (_) {}
        }

        const baseClasses = new Set([...(newEvent.classNames || []), 'fc-event-pill']);
        newEvent.setProp('classNames', Array.from(baseClasses));
      } catch (e) {
        console.warn('eventReceive: cloning/normalize failed', e);
      }

      finalizeSidebarEventDrop(newEvent);
    },

    /* -------------------------
       Plain DOM drop fallback
    ------------------------- */
    drop(info) {
      // FullCalendar also calls eventReceive for these draggable sidebar pills.
      // Keeping this as a no-op prevents duplicate add/finalize work on every drop.
      if (info?.jsEvent?.preventDefault) info.jsEvent.preventDefault();
    },

    /* -------------------------
       Dragging existing events
    ------------------------- */
    eventDrop(info) {
      const eventId = info.event.extendedProps?.id || info.event.id;
      const multiDragActive =
        (typeof window !== 'undefined' && window.__multiSelectDragActive) ||
        (typeof isDragging !== 'undefined' && !!isDragging);

      if (
        multiDragActive ||
        (typeof selectedSchedules !== 'undefined' &&
          selectedSchedules?.size > 1 &&
          selectedSchedules.has(String(eventId)))
      ) {
        info.revert();
        return;
      }

      const movedEmpNo = info.event.extendedProps?.empNo;
      if (movedEmpNo) {
        const conflict = findExistingShiftForEmployee(
          movedEmpNo,
          info.event.startStr,
          { exclude: [info.event] }
        );
        if (conflict) {
          const employeeName = (typeof employees !== 'undefined' && employees?.[movedEmpNo]?.name) || movedEmpNo;
          const existType =
            (conflict.extendedProps && conflict.extendedProps.type === 'rest')
              ? 'rest day'
              : 'work shift';
          const movingType =
            (info.event.extendedProps && info.event.extendedProps.type === 'rest')
              ? 'rest day'
              : 'work shift';

          showToast(
            `Move blocked: ${employeeName} already has a ${existType} on this date. Cannot overlap with another ${movingType}.`,
            'error'
          );
          info.revert();
          return;
        }
      }

      scheduleCalendarMaintenance();
    },

    eventRemove() { scheduleCalendarMaintenance(); },
    eventAdd() { scheduleCalendarMaintenance({ delay: 20, save: false }); },
    eventChange() { scheduleCalendarMaintenance({ delay: 20 }); },

    eventClick(info) {
      if (info.jsEvent && hasMultiSelectModifier(info.jsEvent)) { info.jsEvent.preventDefault(); return; }
      openDeleteModal(info.event);
    },

    eventDidMount(info) {
      try {
        if (!info.event.extendedProps?.id) {
          const fallbackId = (typeof crypto !== 'undefined' && crypto.randomUUID)
            ? crypto.randomUUID()
            : (info.event.id || `evt-${Date.now()}`);
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

    eventContent(arg) {
  const { shiftCode, type } = arg.event.extendedProps;
  const shiftHtml = shiftCode ? `<span class="pill-shift">${shiftCode}</span>` : '';
  const restHtml  = type === 'rest' ? `<span class="pill-rest">REST</span>` : ''; // <-- add this
  return {
    html: `<div class="pill-content"><span class="pill-name">${arg.event.title}</span>${shiftHtml}${restHtml}</div>`
  };
},

    dayCellClassNames(arg) {
      if (arg.date.getDay() === 0 || arg.date.getDay() === 6) return 'fc-day-sat-sun';
      return null;
    },
  });

  calendar.render();
  try { window.calendar = calendar; } catch (e) {}
}

  if (calendarEl && !calendarEl.__sidebarDragOverBound) {
    calendarEl.addEventListener('dragover', (e) => {
      if (__isSidebarDragging) {
        e.preventDefault();
      }
    });
    calendarEl.__sidebarDragOverBound = true;
  }

/* ===========================
   DRAGGABLE (SIDEBAR) INIT
=========================== */
function initializeDraggable() {
  if (!draggableCardsContainer) return;

  draggableCardsContainer.classList.remove('is-dragging');

  // Wire sidebar drag/scroll listeners ONCE (not inside per-card wiring)
  if (!__sidebarDragListenersWired) {
    // When dragging a pill from the sidebar, stop wheel/touchmove bubbling so the calendar doesn't hijack scroll
    draggableCardsContainer.addEventListener('wheel', (e) => {
      if (__isSidebarDragging) e.stopPropagation();
    }, { passive: true });

    draggableCardsContainer.addEventListener('touchmove', (e) => {
      if (__isSidebarDragging) e.stopPropagation();
    }, { passive: true });

    __sidebarDragListenersWired = true;
  }

  // Wire each pill for native HTML5 drag only once
  const sidebarPills = draggableCardsContainer.querySelectorAll('.fc-event-pill');
  sidebarPills.forEach(card => wireSidebarCardDrag(card));

  // Guard against body transforms during external HTML5 drag — wire once
  if (!__bodyDragGuardsWired) {
    draggableCardsRoot?.addEventListener('dragstart', () => {
      document.body.classList.add('no-transform-during-drag');
    });
    draggableCardsRoot?.addEventListener('dragend', () => {
      document.body.classList.remove('no-transform-during-drag');
    });
    __bodyDragGuardsWired = true;
  }
}

function wireSidebarCardDrag(card) {
  if (!card || wiredSidebarCards.has(card)) return;
  card.setAttribute('draggable', 'true');
  card.addEventListener('dragstart', handleSidebarCardDragStart);
  card.addEventListener('dragend', handleSidebarCardDragEnd);
  wiredSidebarCards.add(card);
}

function prepareSidebarEventPayload(eventEl) {
  if (!eventEl) return null;
  try {
    if (!eventEl._fcEventDataTemplate) {
      const raw = eventEl.getAttribute('data-event');
      eventEl._fcEventDataTemplate = raw ? JSON.parse(raw) : null;
    }
    const template = eventEl._fcEventDataTemplate;
    if (!template) return null;

    const parsed = JSON.parse(JSON.stringify(template));
    const uid = `evt_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
    parsed.id = parsed.id || uid;
    parsed.extendedProps = parsed.extendedProps || {};
    parsed.extendedProps.id = parsed.extendedProps.id || uid;

    const type = parsed.extendedProps?.type === 'rest' ? 'rest' : 'work';
    parsed.extendedProps.type = type;
    const classSet = new Set(parsed.classNames || []);
    classSet.add('fc-event-pill');
    classSet.add(type === 'rest' ? 'fc-event-rest' : 'fc-event-work');
    parsed.classNames = Array.from(classSet);

    return parsed;
  } catch (err) {
    console.error('Invalid event data on draggable element', err);
    return null;
  }
}

function clearSidebarDragImage() {
  if (!sidebarDragImageEl) return;
  try {
    sidebarDragImageEl.remove();
  } catch (e) {
    if (sidebarDragImageEl.parentNode) sidebarDragImageEl.parentNode.removeChild(sidebarDragImageEl);
  }
  sidebarDragImageEl = null;
}

function setSidebarDragImage(ev, card) {
  if (!ev?.dataTransfer || !card) return;
  clearSidebarDragImage();

  const rect = card.getBoundingClientRect();
  const baseWidth = card.offsetWidth || rect.width || 0;
  const baseHeight = card.offsetHeight || rect.height || 0;
  const scaleX = baseWidth ? rect.width / baseWidth : 1;
  const scaleY = baseHeight ? rect.height / baseHeight : 1;
  const clone = card.cloneNode(true);
  clone.style.position = 'fixed';
  clone.style.top = '-1000px';
  clone.style.left = '-1000px';
  clone.style.margin = '0';
  clone.style.transformOrigin = 'top left';
  clone.style.transform = `scale(${scaleX}, ${scaleY})`;
  clone.style.pointerEvents = 'none';
  clone.style.boxSizing = 'border-box';
  clone.style.width = `${baseWidth || rect.width}px`;
  clone.style.height = `${baseHeight || rect.height}px`;
  document.body.appendChild(clone);
  sidebarDragImageEl = clone;

  const rawOffsetX = typeof ev.clientX === 'number' ? ev.clientX - rect.left : rect.width / 2;
  const rawOffsetY = typeof ev.clientY === 'number' ? ev.clientY - rect.top : rect.height / 2;
  const normalizedOffsetX = Number.isFinite(rawOffsetX) ? rawOffsetX / (scaleX || 1) : (baseWidth || rect.width) / 2;
  const normalizedOffsetY = Number.isFinite(rawOffsetY) ? rawOffsetY / (scaleY || 1) : (baseHeight || rect.height) / 2;
  const maxOffsetX = baseWidth || rect.width;
  const maxOffsetY = baseHeight || rect.height;
  const offsetX = Math.min(maxOffsetX, Math.max(0, normalizedOffsetX));
  const offsetY = Math.min(maxOffsetY, Math.max(0, normalizedOffsetY));

  try { ev.dataTransfer.setDragImage(clone, offsetX, offsetY); } catch (e) {}
}

function handleSidebarCardDragStart(ev) {
  const card = ev.currentTarget;
  const payload = prepareSidebarEventPayload(card);
  if (!payload) {
    ev.preventDefault();
    return;
  }

  const payloadJson = JSON.stringify(payload);
  card.setAttribute('data-drag-payload', payloadJson);
  __isSidebarDragging = true;
  draggableCardsContainer?.classList.add('is-dragging');
  setGlobalDragFlag(true);
  document.body.classList.add('no-transform-during-drag');
  setSidebarDragImage(ev, card);

  if (ev.dataTransfer) {
    try { ev.dataTransfer.effectAllowed = 'copyMove'; } catch (e) {}
    try { ev.dataTransfer.setData('application/json', payloadJson); } catch (e) {}
    try { ev.dataTransfer.setData('text/plain', payloadJson); } catch (e) {}
  }
}

function handleSidebarCardDragEnd(ev) {
  const card = ev.currentTarget;
  __isSidebarDragging = false;
  draggableCardsContainer?.classList.remove('is-dragging');
  setGlobalDragFlag(false);
  document.body.classList.remove('no-transform-during-drag');
  clearSidebarDragImage();
  if (card) card.removeAttribute('data-drag-payload');
}

function parseSidebarDragPayload(draggedEl, dataTransfer) {
  let raw = draggedEl?.getAttribute('data-drag-payload');
  if (!raw) raw = draggedEl?.getAttribute('data-event');
  if (!raw && dataTransfer) {
    raw = dataTransfer.getData('application/json') || dataTransfer.getData('text/plain') || dataTransfer.getData('text');
  }
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (err) {
    console.warn('Unable to parse sidebar drag payload', err);
    return null;
  }
}

function normalizeSidebarPayload(payload) {
  if (!payload) return null;
  const clone = JSON.parse(JSON.stringify(payload));
  clone.extendedProps = clone.extendedProps || {};
  const type = clone.extendedProps?.type === 'rest' ? 'rest' : 'work';
  clone.extendedProps.type = type;
  const classSet = new Set(clone.classNames || []);
  classSet.add('fc-event-pill');
  classSet.add(type === 'rest' ? 'fc-event-rest' : 'fc-event-work');
  clone.classNames = Array.from(classSet);
  if (!clone.id) {
    const uid = `evt_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
    clone.id = uid;
    clone.extendedProps.id = clone.extendedProps.id || uid;
  } else if (!clone.extendedProps.id) {
    clone.extendedProps.id = clone.id;
  }
  return clone;
}

function finalizeSidebarEventDrop(newEvent) {
  if (!newEvent) return;

  let dateStr = newEvent.startStr;

  (function forceDropToPointerDate() {
    const pointedDate = (typeof getDateUnderPointer === 'function') ? getDateUnderPointer() : null;
    if (pointedDate && pointedDate !== newEvent.startStr) {
      newEvent.setStart(pointedDate);
      try { newEvent.setEnd(pointedDate); } catch (e) {}
      dateStr = newEvent.startStr;
    }
  })();

  const type = newEvent.extendedProps?.type || 'work';
  const empNo = newEvent.extendedProps?.empNo || newEvent.extendedProps?.employeeNo;
  if (!empNo) { console.warn('eventReceive: missing empNo'); return; }

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

  const conflict = findExistingShiftForEmployee(empNo, dateStr, { exclude: [newEvent] });
  if (conflict) {
    const name = employees[empNo]?.name || empNo;
    const existType = (conflict.extendedProps && conflict.extendedProps.type === 'rest') ? 'rest day' : 'work shift';
    const newType = (type === 'rest') ? 'rest day' : 'work shift';
    showToast(`Duplicate entry blocked: ${name} already has a ${existType} on this date. Cannot add another ${newType}.`, 'error');
    newEvent.remove();
    return;
  }

  if (type === 'work') {
    currentDroppingEvent = newEvent;
    openShiftModal();
  } else {
    scheduleCalendarMaintenance();
  }

  try { newEvent.setProp('title', newEvent.title); } catch (e) {}
  try { decorateEventLater(newEvent); } catch (e) {}
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

  function addEmployeeRowWithValues({ name = '', empNo = '', position = '' }) {
    addEmployeeRow();
    const rows = employeeTableBody ? employeeTableBody.querySelectorAll('tr') : [];
    const lastRow = rows[rows.length - 1];
    if (!lastRow) return;
    const nameInput = lastRow.querySelector('.emp-name');
    const noInput = lastRow.querySelector('.emp-no');
    const posInput = lastRow.querySelector('.emp-pos');
    if (nameInput) nameInput.value = name;
    if (noInput) noInput.value = empNo;
    if (posInput) posInput.value = position;
  }

  function findPositionMatch(position) {
    const normalized = String(position || '').trim().toLowerCase();
    if (!normalized) return '';
    const match = positionOptions.find(option => option.toLowerCase() === normalized);
    return match || '';
  }

  function ensurePositionOption(position) {
    const trimmed = String(position || '').trim();
    if (!trimmed) return '';
    const match = findPositionMatch(trimmed);
    if (match) return match;
    positionOptions.push(trimmed);
    document.querySelectorAll('.emp-pos').forEach(select => {
      const exists = Array.from(select.options).some(option => option.value.toLowerCase() === trimmed.toLowerCase());
      if (!exists) {
        const option = document.createElement('option');
        option.value = trimmed;
        option.textContent = trimmed;
        select.appendChild(option);
      }
    });
    return trimmed;
  }

  function normalizeHeaderValue(value) {
    if (value == null) return '';
    return String(value).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  function findHeaderIndex(headers, candidates) {
    const normalized = headers.map(normalizeHeaderValue);
    return normalized.findIndex(item => candidates.includes(item));
  }

  function resetImportModalState() {
    pendingImportRows = [];
    importMode = 'preview';
    if (importPreviewBody) {
      importPreviewBody.innerHTML = `
        <tr>
          <td colspan="3" class="p-6 text-center text-gray-400">
            Select a file to preview the first rows before importing.
          </td>
        </tr>`;
    }
    if (importPreviewNote) importPreviewNote.textContent = '';
    if (importSummaryPanel) importSummaryPanel.classList.add('hidden');
    if (importSummaryCounts) importSummaryCounts.textContent = '';
    if (importSummaryList) importSummaryList.innerHTML = '';
    if (confirmImportBtn) confirmImportBtn.textContent = 'Import Employees';
    if (confirmImportBtn) {
      confirmImportBtn.disabled = false;
      confirmImportBtn.classList.remove('opacity-50');
      confirmImportBtn.classList.remove('hidden');
    }
    if (cancelImportBtn) {
      cancelImportBtn.textContent = 'Cancel';
      cancelImportBtn.classList.remove('hidden');
    }
  }

  function renderImportPreview(rows) {
    if (!importPreviewBody) return;
    const previewLimit = 15;
    const previewRows = rows.slice(0, previewLimit);
    if (previewRows.length === 0) {
      importPreviewBody.innerHTML = `
        <tr>
          <td colspan="3" class="p-6 text-center text-gray-400">No employee rows found to preview.</td>
        </tr>`;
      if (importPreviewNote) importPreviewNote.textContent = '';
      return;
    }
    importPreviewBody.innerHTML = previewRows.map(row => `
      <tr>
        <td class="p-3 text-gray-700">${escapeHtml(row.empNo)}</td>
        <td class="p-3 text-gray-700">${escapeHtml(row.name)}</td>
        <td class="p-3 text-gray-700">${escapeHtml(row.position)}</td>
      </tr>`).join('');
    if (importPreviewNote) {
      const total = rows.length;
      const note = total > previewLimit
        ? `Showing ${previewLimit} of ${total} rows.`
        : `Showing ${total} row${total === 1 ? '' : 's'}.`;
      importPreviewNote.textContent = note;
    }
  }

  function renderImportSummary(summary) {
    if (!importSummaryPanel || !importSummaryCounts || !importSummaryList) return;
    importSummaryPanel.classList.remove('hidden');
    importSummaryCounts.textContent = `Imported ${summary.imported} row${summary.imported === 1 ? '' : 's'}, skipped ${summary.skipped} row${summary.skipped === 1 ? '' : 's'}, failed ${summary.failed} row${summary.failed === 1 ? '' : 's'}.`;
    const lines = [];
    summary.details.forEach(item => {
      lines.push(`<li>${escapeHtml(item)}</li>`);
    });
    importSummaryList.innerHTML = lines.length ? lines.join('') : '<li>No issues detected.</li>';
  }

  function extractEmployeeRowsFromSheet(sheet) {
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const headerRowIndex = rawRows.findIndex(row => Array.isArray(row) && row.some(cell => String(cell).trim() !== ''));
    if (headerRowIndex === -1) {
      return { error: 'No header row detected in the selected file.' };
    }
    const headers = rawRows[headerRowIndex] || [];
    const requiredHeaderOptions = {
      empNo: ['employeenumber', 'employeeno', 'empno', 'employeeid', 'employeenum', 'employeenumber'],
      name: ['employeename', 'name', 'fullname'],
      position: ['position', 'role', 'title']
    };
    const empNoIndex = findHeaderIndex(headers, requiredHeaderOptions.empNo);
    const nameIndex = findHeaderIndex(headers, requiredHeaderOptions.name);
    const positionIndex = findHeaderIndex(headers, requiredHeaderOptions.position);
    const missing = [];
    if (empNoIndex === -1) missing.push('Employee Number');
    if (nameIndex === -1) missing.push('Employee Name');
    if (positionIndex === -1) missing.push('Position');
    if (missing.length) {
      return { error: `Missing required column(s): ${missing.join(', ')}.` };
    }
    const dataRows = rawRows.slice(headerRowIndex + 1).filter(row => Array.isArray(row) && row.some(cell => String(cell).trim() !== ''));
    const mapped = dataRows.map((row, idx) => ({
      empNo: row[empNoIndex] != null ? String(row[empNoIndex]).trim() : '',
      name: row[nameIndex] != null ? String(row[nameIndex]).trim() : '',
      position: row[positionIndex] != null ? String(row[positionIndex]).trim() : '',
      sourceRow: headerRowIndex + 2 + idx
    }));
    return { rows: mapped };
  }

  function applyImportedEmployees(rows) {
    const summary = { imported: 0, skipped: 0, failed: 0, details: [] };
    const existingEmpNos = new Set();
    const tableRows = employeeTableBody ? employeeTableBody.querySelectorAll('tr') : [];
    tableRows.forEach(row => {
      const noInput = row.querySelector('.emp-no');
      const empNo = noInput ? noInput.value.trim() : '';
      if (empNo) existingEmpNos.add(empNo);
    });
    const seenInImport = new Set();
    rows.forEach(row => {
      const empNo = row.empNo ? row.empNo.trim() : '';
      const name = row.name ? row.name.trim() : '';
      const rawPosition = row.position ? row.position.trim() : '';
      const position = ensurePositionOption(rawPosition);

      if (!empNo || !name || !rawPosition) {
        summary.failed += 1;
        summary.details.push(`Row ${row.sourceRow}: missing required fields.`);
        return;
      }
      if (seenInImport.has(empNo)) {
        summary.skipped += 1;
        summary.details.push(`Row ${row.sourceRow}: duplicate Employee Number ${empNo} in import file.`);
        return;
      }
      if (existingEmpNos.has(empNo)) {
        summary.skipped += 1;
        summary.details.push(`Row ${row.sourceRow}: Employee Number ${empNo} already exists.`);
        return;
      }
      seenInImport.add(empNo);
      existingEmpNos.add(empNo);
      addEmployeeRowWithValues({ name, empNo, position });
      summary.imported += 1;
    });
    return summary;
  }

  function handleImportFile(file) {
    if (!file) return;
    const fileName = file.name || '';
    const ext = fileName.split('.').pop().toLowerCase();
    const allowedExtensions = ['xlsx', 'numbers'];
    if (!allowedExtensions.includes(ext)) {
      showToast('Please upload an Excel (.xlsx) or Apple Numbers (.numbers) file.', 'error');
      return;
    }
    if (typeof XLSX === 'undefined') {
      showToast('Import requires the XLSX library, which is unavailable.', 'error');
      return;
    }
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        if (!sheetName) {
          showToast('No worksheet found in the uploaded file.', 'error');
          return;
        }
        const result = extractEmployeeRowsFromSheet(workbook.Sheets[sheetName]);
        if (result.error) {
          showToast(result.error, 'error');
          return;
        }
        pendingImportRows = result.rows || [];
        renderImportPreview(pendingImportRows);
        importMode = 'preview';
        if (importSummaryPanel) importSummaryPanel.classList.add('hidden');
        if (confirmImportBtn) {
          confirmImportBtn.textContent = 'Import Employees';
          confirmImportBtn.disabled = pendingImportRows.length === 0;
          confirmImportBtn.classList.toggle('opacity-50', confirmImportBtn.disabled);
        }
        if (importModal) importModal.classList.remove('hidden');
      } catch (error) {
        console.error('Import parse error:', error);
        showToast('Unable to read that file. If using Numbers, export to .xlsx first.', 'error');
      }
    };
    reader.onerror = () => {
      showToast('Unable to read the selected file.', 'error');
    };
    reader.readAsArrayBuffer(file);
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
                   style="background:${workGradient}; color:#fff; border:none;">Work</div>
              <div class="fc-event-pill fc-event-rest px-3 py-1 text-xs font-medium rounded-full cursor-pointer select-none shadow-sm"
                   data-event='${JSON.stringify(restEventData)}'
                   style="background:${restGradient}; color:${color}; border:2px solid ${color};">Rest</div>
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
      scheduleCalendarMaintenance();
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
    if (currentDeletingEvent) {
      currentDeletingEvent.remove();
      showToast('Schedule entry removed.', 'info');
      scheduleCalendarMaintenance();
    }
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
      currentDroppingEvent.remove();
      showToast('Schedule add canceled.', 'info');
      scheduleCalendarMaintenance();
    }
    if (modalId === 'shift-modal') { currentDroppingEvent = null; isEditingExistingEvent = false; }
    if (modalId === 'delete-modal') { currentDeletingEvent = null; }
    if (modalId === 'import-modal') {
      resetImportModalState();
      if (importFileInput) importFileInput.value = '';
    }
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

// Weekend RD rule copied from the schedule checker app.
const WEEKEND_REST_MONTH_LIMIT = 4;
const WEEKEND_REST_DAYS = ['Friday', 'Saturday', 'Sunday'];

// ISO week key (YYYY-Www). Week starts Monday (ISO-8601).
function isoWeekKeyFromDate(d) {
  // clone to avoid mutating
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  // Thursday in current week decides the year
  date.setUTCDate(date.getUTCDate() + 4 - (date.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  const yyyy = date.getUTCFullYear();
  return `${yyyy}-W${String(weekNo).padStart(2, '0')}`;
}

// Anchor any Fri/Sat/Sun to the FRIDAY of that same weekend (LOCAL time)
function weekendWeekKeyLocal(d) {
  const dt = new Date(d.getFullYear(), d.getMonth(), d.getDate()); // local midnight
  const dow = dt.getDay(); // Sun=0 .. Sat=6
  if (!(dow === 5 || dow === 6 || dow === 0)) return null; // only Fri/Sat/Sun

// distance TO Friday (Fri=0, Sat=-1, Sun=-2)
const deltaToFri = (dow === 0) ? -2 : (5 - dow);
const fri = new Date(dt);
fri.setDate(dt.getDate() + deltaToFri);

  // Use Friday date string as the unique weekend key
  const y = fri.getFullYear();
  const m = String(fri.getMonth() + 1).padStart(2, '0');
  const d2 = String(fri.getDate()).padStart(2, '0');
  return `${y}-${m}-${d2}`;
}

function getLocalDateKey(dateLike) {
  const d = dateLike instanceof Date ? dateLike : new Date(dateLike);
  if (Number.isNaN(d.getTime())) return '';
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function getWeekStartLocalKey(dateLike) {
  const d = dateLike instanceof Date ? new Date(dateLike) : new Date(dateLike);
  if (Number.isNaN(d.getTime())) return '';
  const day = d.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + mondayOffset);
  return getLocalDateKey(d);
}

function getDayNameLocal(dateLike) {
  const d = dateLike instanceof Date ? dateLike : new Date(dateLike);
  if (Number.isNaN(d.getTime())) return '';
  return ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][d.getDay()] || '';
}

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
  const allEvents = calendar ? calendar.getEvents() : [];

  // Same-day Work & Rest conflict
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

  // Duplicate dates within the same schedule type, matching app-51-fixed.js.
  ['work', 'rest'].forEach(type => {
    const seen = new Map();
    const label = type === 'work' ? 'Work Schedule' : 'Rest Day Schedule';
    allEvents.forEach(event => {
      if (event.extendedProps?.type !== type) return;
      const empNo = event.extendedProps?.empNo;
      const date = event.startStr;
      if (!empNo || !date) return;
      const key = `${empNo}-${date}`;
      if (seen.has(key)) {
        const other = seen.get(key);
        conflicts.push({ empNo, date, rule: `Duplicate date in ${label}`, event: other });
        conflicts.push({ empNo, date, rule: `Duplicate date in ${label}`, event });
      } else {
        seen.set(key, event);
      }
    });
  });

  // RD employee must also have at least one WS row/event.
  const workEmpNos = new Set(allEvents
    .filter(event => event.extendedProps?.type === 'work' && event.extendedProps?.empNo)
    .map(event => String(event.extendedProps.empNo)));

  allEvents.forEach(event => {
    const empNo = event.extendedProps?.empNo;
    if (event.extendedProps?.type === 'rest' && empNo && !workEmpNos.has(String(empNo))) {
      conflicts.push({
        empNo,
        date: event.startStr,
        rule: 'Employee not found in Work Schedule',
        event
      });
    }
  });

  // Multiple managers resting same day
  for (const date in eventsByDate) {
    const managerRestEvents = eventsByDate[date].filter(event => {
      const emp = employees[event.extendedProps.empNo];
      return event.extendedProps.type === 'rest' && emp && managerPositions.includes(emp.position);
    });
    if (managerRestEvents.length > 1) {
      managerRestEvents.forEach(event => { conflicts.push({ empNo: event.extendedProps.empNo, date, rule: 'Multiple Managers Resting', event }); });
    }
  }

  // Weekend RD validation copied from app-51-fixed.js:
  // Friday/Saturday/Sunday count as weekend groups per employee/month,
  // Saturday-Sunday consecutive RD is not allowed, and max is 4 groups/month.
  const employeeMonthWeekMap = {};
  allEvents.forEach(event => {
    if (event.extendedProps?.type !== 'rest') return;
    const empNo = event.extendedProps?.empNo;
    if (!empNo || !event.start) return;

    const dayName = getDayNameLocal(event.start);
    if (!WEEKEND_REST_DAYS.includes(dayName)) return;

    const year = event.start.getFullYear();
    const month = event.start.getMonth() + 1;
    const empKey = `${empNo}-${year}-${month}`;
    const weekKey = getWeekStartLocalKey(event.start);
    if (!weekKey) return;

    if (!employeeMonthWeekMap[empKey]) employeeMonthWeekMap[empKey] = {};
    if (!employeeMonthWeekMap[empKey][weekKey]) employeeMonthWeekMap[empKey][weekKey] = [];
    employeeMonthWeekMap[empKey][weekKey].push({ event, empNo, dayName });
  });

  Object.values(employeeMonthWeekMap).forEach(weeks => {
    let weekendGroupCount = 0;

    Object.values(weeks).forEach(entries => {
      const days = entries.map(e => e.dayName);
      const hasFriday = days.includes('Friday');
      const hasSaturday = days.includes('Saturday');
      const hasSunday = days.includes('Sunday');

      if (hasSaturday && hasSunday) {
        entries
          .filter(e => e.dayName === 'Saturday' || e.dayName === 'Sunday')
          .forEach(e => {
            conflicts.push({
              empNo: e.empNo,
              date: e.event.startStr,
              rule: 'Saturday-Sunday consecutive Rest Days are not allowed.',
              event: e.event
            });
          });
      }

      if (hasFriday || hasSaturday || hasSunday) weekendGroupCount++;
    });

    if (weekendGroupCount > WEEKEND_REST_MONTH_LIMIT) {
      Object.values(weeks).flat().forEach(e => {
        conflicts.push({
          empNo: e.empNo,
          date: e.event.startStr,
          rule: `Maximum weekend RD groups exceeded. Allowed maximum is ${WEEKEND_REST_MONTH_LIMIT} per month.`,
          event: e.event
        });
      });
    }
  });

  // Weekly requirement: employees marked mustHavePerWeek must have at least one event in each ISO week within the current range
  {
    let rangeStart = calendar?.view?.activeStart;
    let rangeEnd = calendar?.view?.activeEnd;
    if (!rangeStart || !rangeEnd) {
      allEvents.forEach(ev => {
        if (!rangeStart || ev.start < rangeStart) rangeStart = ev.start;
        if (!rangeEnd || ev.start > rangeEnd) rangeEnd = ev.start;
      });
    }

    const iterEnd = rangeEnd ? new Date(rangeEnd) : null;
    if (rangeStart && iterEnd) {
      if (iterEnd.getTime() === rangeStart.getTime()) iterEnd.setDate(iterEnd.getDate() + 1);
      const weekKeysInRange = new Set();
      const cursor = new Date(rangeStart);
      while (cursor < iterEnd) {
        weekKeysInRange.add(isoWeekKeyFromDate(cursor));
        cursor.setDate(cursor.getDate() + 1);
      }

      for (const empNo in employees) {
        if (employees[empNo]?.mustHavePerWeek !== true) continue;
        const empEvents = eventsByEmp[empNo] || [];
        const empWeekKeys = new Set();
        empEvents.forEach(ev => { empWeekKeys.add(isoWeekKeyFromDate(ev.start)); });

        weekKeysInRange.forEach(weekKey => {
          if (!empWeekKeys.has(weekKey)) {
            conflicts.push({
              empNo,
              date: weekKey,
              rule: 'No Schedule This Week',
              event: null,
              virtualId: `missing-${empNo}-${weekKey}`
            });
          }
        });
      }
    }
  }

  return conflicts;
}

  function runConflictDetection() {
    runWithCalendarBatch(() => {
      const { allEvents } = getGroupedEvents();
      if (allEvents && allEvents.forEach) {
        allEvents.forEach(event => {
          if (event.extendedProps?.isConflict) setEventConflictFlag(event, false);
        });
      }
      const placeholderRow = ensureConflictPlaceholderRow();
      if (conflictTableBody) conflictTableBody.innerHTML = '';

      const conflicts = getConflicts();
      const conflictEvents = new Set();
      const conflictTableEntries = {};
      conflicts.forEach(conflict => {
        if (conflict.event) {
          if (!conflict.event.extendedProps?.isConflict) setEventConflictFlag(conflict.event, true);
          conflictEvents.add(conflict.event.extendedProps.id);
        } else if (conflict.virtualId) {
          conflictEvents.add(conflict.virtualId);
        }
        const key = `${conflict.empNo}-${conflict.rule}`;
        if (!conflictTableEntries[key]) conflictTableEntries[key] = { empNo: conflict.empNo, rule: conflict.rule, dates: new Set() };
        conflictTableEntries[key].dates.add(conflict.date);
      });

      if (Object.keys(conflictTableEntries).length > 0 && conflictTableBody) {
        if (conflictsPlaceholder) conflictsPlaceholder.classList.add('hidden');
        const fragment = document.createDocumentFragment();
        Object.values(conflictTableEntries).forEach(entry => {
          const emp = employees[entry.empNo] || { name: entry.empNo, empNo: entry.empNo };
          const tr = document.createElement('tr');
          tr.className = 'bg-yellow-50 border-b border-yellow-200';
          tr.innerHTML = `
            <td class="p-3 font-medium text-yellow-900">${emp.name}</td>
            <td class="p-3 text-yellow-800">${emp.empNo}</td>
            <td class="p-3 text-yellow-800">${entry.rule}</td>
            <td class="p-3 text-yellow-800">${[...entry.dates].join(', ')}</td>`;
          fragment.appendChild(tr);
        });
        conflictTableBody.appendChild(fragment);
      } else if (conflictTableBody && placeholderRow) {
        placeholderRow.classList.remove('hidden');
        conflictTableBody.appendChild(placeholderRow);
      }

      updateStats(conflictEvents.size);
    });
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

    // --- NEW: gather required report fields ---
const brCodeEl = document.getElementById('export-br-code');
const branchNameEl = document.getElementById('export-branch-name');
const monthEl = document.getElementById('export-month');

const brCode = (brCodeEl?.value || '').trim();
const branchName = (branchNameEl?.value || '').trim();
const monthLabel = (monthEl?.value || '').trim();

if (!brCode || !branchName || !monthLabel) {
  if (typeof showToast === 'function') {
    showToast('Please fill in BR-CODE, BRANCH NAME, and MONTH before exporting.', 'warn');
  }
  return;
}

// safe filename pieces (avoid illegal filename chars)
const clean = s => String(s).replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim();
const fileName = `${clean(brCode)}_${clean(branchName)}_${clean(monthLabel)}_Schedule.xlsx`;

// we’ll also add a small “Report Info” sheet as the first tab

    const conflicts = typeof getConflicts === 'function' ? getConflicts() : [];
    const allEvents = calendar.getEvents();
    const daysOfWeek = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const wb = XLSX.utils.book_new();

    const infoData = [
  ['BR-CODE', brCode],
  ['BRANCH NAME', branchName],
  ['MONTH', monthLabel],
  ['Generated', new Date().toLocaleString()]
];
const wsInfo = XLSX.utils.aoa_to_sheet(infoData);
try {
  // make left column a bit wider
  wsInfo['!cols'] = [{ wpx: 140 }, { wpx: 420 }];
  // bold left keys
  infoData.forEach((row, r) => {
    const addr = XLSX.utils.encode_cell({ c: 0, r });
    if (wsInfo[addr]) wsInfo[addr].s = Object.assign({}, wsInfo[addr].s || {}, { font: { bold: true } });
  });
} catch (_) {}
XLSX.utils.book_append_sheet(wb, wsInfo, 'Report Info');

    // Helper: pixel-based auto-fit (works better than 'wch' across Excel versions)
    function autoFitColumnsWpx(aoa, floorsPx, maxPx = 1200, padPx = 18) {
      const widths = [];
      for (let r = 0; r < aoa.length; r++) {
        const row = aoa[r] || [];
        for (let c = 0; c < row.length; c++) {
          const v = row[c];
          const s = (v === null || typeof v === 'undefined') ? '' : String(v);
          const maxLineLen = s.split(/\r?\n/).reduce((m, seg) => Math.max(m, seg.length), 0);
          const px = (maxLineLen * 8) + padPx; // ~8px per character baseline
          widths[c] = Math.max(widths[c] || 0, px);
        }
      }
      return widths.map((px, i) => ({ wpx: Math.min(maxPx, Math.max((floorsPx && floorsPx[i]) || 0, px || 0)) }));
    }

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
      // Bold header
      for (let c = 0; c < workData[0].length; c++) {
        const addr = XLSX.utils.encode_cell({ c, r: 0 });
        if (!wsWork[addr]) continue;
        wsWork[addr].s = Object.assign({}, wsWork[addr].s || {}, { font: { bold: true } });
      }
      // Auto-size all columns (NAME & POSITION floors raised to avoid clipping)
      wsWork['!cols'] = autoFitColumnsWpx(workData, [200, 130, 260, 140, 120, 140], 1000, 22);
      // Add AutoFilter on header (optional quality-of-life)
      wsWork['!autofilter'] = { ref: 'A1:F' + workData.length };
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
      wsRest['!cols'] = autoFitColumnsWpx(restData, [200, 130, 260, 140, 140], 1000, 22);
      wsRest['!autofilter'] = { ref: 'A1:E' + restData.length };
    } catch (_) {}
    XLSX.utils.book_append_sheet(wb, wsRest, 'Rest Schedule');

    // ===== Conflict Summary =====
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
        wsCon['!cols'] = autoFitColumnsWpx(conflictData, [240, 130, 300, 340], 1100, 22);
        wsCon['!autofilter'] = { ref: 'A1:D' + conflictData.length };
      } catch (_) {}
      XLSX.utils.book_append_sheet(wb, wsCon, 'Conflict Summary');
    }

    // Write the file and close
    XLSX.writeFile(wb, fileName);
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
      initializeDraggable();
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

  const pendingEventDecorations = new Set();
  let decorationScheduled = false;

  function flushPendingDecorations() {
    decorationScheduled = false;
    pendingEventDecorations.forEach(ev => decorateEventLater(ev));
    pendingEventDecorations.clear();
  }

  function queueEventDecoration(event) {
    if (!event) return;
    pendingEventDecorations.add(event);
    if (decorationScheduled) return;
    decorationScheduled = true;
    if (typeof requestAnimationFrame === 'function') requestAnimationFrame(flushPendingDecorations);
    else setTimeout(flushPendingDecorations, 0);
  }

  function decorateCalendarEvents(targetEvents = null) {
    const events = targetEvents
      ? (Array.isArray(targetEvents) ? targetEvents : [targetEvents])
      : (calendar ? calendar.getEvents() : []);
    events.forEach(queueEventDecoration);
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
    scheduleCalendarMaintenance();
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
    scheduleCalendarMaintenance();
    showToastWrapper(`Undid paste (${removed} schedule${removed !== 1 ? 's' : ''} removed).`, 'info');
  }

  // Mouse tracking for copy/paste hover
  let mouseUpdateQueued = false;
  const updateMouse = e => {
    lastMouseX = e.clientX; lastMouseY = e.clientY;
    if (mouseUpdateQueued) return;
    mouseUpdateQueued = true;
    const runner = () => {
      mouseUpdateQueued = false;
      if (isDragPerfSuppressed()) return;
      const pointerTarget = document.elementFromPoint(lastMouseX,lastMouseY);
      const pill = pointerTarget?.closest('.schedule-pill');
      if (pill) {
        const id = getScheduleIdFromElement(pill);
        if (id) hoveredScheduleId = id;
      } else if (!pointerTarget?.closest('.schedule-pill')) {
        hoveredScheduleId = null;
      }
    };
    if (typeof requestAnimationFrame === 'function') requestAnimationFrame(runner); else setTimeout(runner, 16);
  };
  document.addEventListener('mousemove', updateMouse, { passive: true });

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
        isDragging = true; window.__multiSelectDragActive = true; setGlobalDragFlag(true);
        createDragGhost(selectedSchedules.size, e.clientX, e.clientY);
        document.body.classList.add('no-select', 'no-transform-during-drag'); // FIX: lock transforms during ghost drag
        return;
      }
    }

    if (day) {
      isSelectingDates = true; dateSelectStartEl = day;
      if (!hasMultiSelectModifier(e)) clearTargetDateSelection();
      toggleTargetDateSelection(day, true);
      rangeDayCells = Array.from(document.querySelectorAll('.fc-daygrid-day'));
      rangeDayCells.forEach(dayEl => dayEl.classList.remove('range-selecting'));
      lastRangeSelection = null;
      document.body.classList.add('no-select');
    }
  });

  let rangeDayCells = null;
  let lastRangeSelection = null;
  let dragFrameQueued = false;
  let pendingDragPoint = null;

  function updateRangeSelection(startIdx, endIdx) {
    if (!rangeDayCells) return;
    if (startIdx === -1 || endIdx === -1) {
      if (lastRangeSelection) {
        const { min, max } = lastRangeSelection;
        for (let i = min; i <= max; i++) rangeDayCells[i]?.classList.remove('range-selecting');
        lastRangeSelection = null;
      }
      return;
    }

    const minIdx = Math.min(startIdx, endIdx);
    const maxIdx = Math.max(startIdx, endIdx);

    if (lastRangeSelection && lastRangeSelection.min === minIdx && lastRangeSelection.max === maxIdx) return;

    if (lastRangeSelection) {
      const { min: prevMin, max: prevMax } = lastRangeSelection;
      for (let i = prevMin; i <= prevMax; i++) {
        if (i < minIdx || i > maxIdx) rangeDayCells[i]?.classList.remove('range-selecting');
      }
    }

    for (let i = minIdx; i <= maxIdx; i++) {
      if (!lastRangeSelection || i < lastRangeSelection.min || i > lastRangeSelection.max) {
        rangeDayCells[i]?.classList.add('range-selecting');
      }
    }

    lastRangeSelection = { min: minIdx, max: maxIdx };
  }

  function processDragFrame() {
    dragFrameQueued = false;
    if (!pendingDragPoint) return;
    const { x, y } = pendingDragPoint;

    if (isDragging && dragGhost) {
      dragGhost.style.left = `${x}px`;
      dragGhost.style.top = `${y}px`;
      dragGhost.style.transform = 'translate(-50%, -50%)';
      return;
    }

    if (isSelectingDates && dateSelectStartEl) {
      if (!rangeDayCells) rangeDayCells = Array.from(document.querySelectorAll('.fc-daygrid-day'));
      const currentEl = document.elementFromPoint(x, y);
      const currentDayEl = currentEl?.closest?.('.fc-daygrid-day');
      const startIdx = rangeDayCells.indexOf(dateSelectStartEl);
      const endIdx = currentDayEl ? rangeDayCells.indexOf(currentDayEl) : -1;
      updateRangeSelection(startIdx, endIdx);
    }
  }

  // FIX: ghost follows viewport coords exactly (throttled)
  document.addEventListener('mousemove', e => {
    pendingDragPoint = { x: e.clientX, y: e.clientY };
    if (dragFrameQueued) return;
    dragFrameQueued = true;
    if (typeof requestAnimationFrame === 'function') requestAnimationFrame(processDragFrame);
    else setTimeout(processDragFrame, 16);
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
      setGlobalDragFlag(false);
      document.body.classList.remove('no-select', 'no-transform-during-drag'); // FIX: cleanup
      buildCopiedFromSelected({ silent: true });
      return;
    }
    if (isSelectingDates) {
      isSelectingDates = false;
      dateSelectStartEl = null;
      rangeDayCells = null;
      lastRangeSelection = null;
      document.body.classList.remove('no-select');
    }
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
    scheduleCalendarMaintenance();
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

  const importBtnEl = document.getElementById('import-btn');
  if (importBtnEl && importFileInput) {
    importBtnEl.addEventListener('click', () => importFileInput.click());
  }
  if (importFileInput) {
    importFileInput.addEventListener('change', (event) => {
      const file = event.target.files && event.target.files[0];
      handleImportFile(file);
      importFileInput.value = '';
    });
  }
  if (confirmImportBtn) {
    confirmImportBtn.addEventListener('click', () => {
      if (importMode === 'preview') {
        if (!pendingImportRows.length) {
          showToast('No employee rows available to import.', 'warn');
          return;
        }
        const summary = applyImportedEmployees(pendingImportRows);
        renderImportSummary(summary);
        importMode = 'summary';
        confirmImportBtn.textContent = 'Close';
        confirmImportBtn.disabled = false;
        confirmImportBtn.classList.remove('opacity-50');
        confirmImportBtn.classList.add('hidden');
        if (cancelImportBtn) {
          cancelImportBtn.textContent = 'Close';
          cancelImportBtn.classList.remove('hidden');
        }
        const message = `Import complete: ${summary.imported} added, ${summary.skipped} skipped, ${summary.failed} failed.`;
        showToast(message, summary.failed ? 'warn' : 'success');
        return;
      }
      handleCloseModal('import-modal');
    });
  }

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
      if (importModal && !importModal.classList.contains('hidden')) handleCloseModal('import-modal');
    }
  });

  function startScheduler() {
    try { initializeCalendar(); } catch (err) { console.error('initializeCalendar failed', err); }
    try { loadFromLocalStorage(); } catch (err) { console.error('loadFromLocalStorage failed', err); }
    try { if (!employeeTableBody || employeeTableBody.querySelectorAll('tr').length === 0) { addEmployeeRow(); } } catch (err) {}
    try {
      decorateCalendarEvents();
      calendar?.on('eventDidMount', info => queueEventDecoration(info?.event));
    } catch (err) { console.error('decorateCalendarEvents failed', err); }
  }

  startScheduler();
});