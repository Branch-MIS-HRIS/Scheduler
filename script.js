document.addEventListener('DOMContentLoaded', function () {

  // ---------- small utility: shadeColor ----------
// color: hex string like "#3b82f6", percent: negative to darken, positive to lighten
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
            
            // --- STATE & CONFIG ---
            
            // In-memory 'database' for employees
            let employees = {};
let employeeColors = JSON.parse(localStorage.getItem('employeeColors')) || {};

/* ====== MOVED HELPERS: make gradient/color helpers top-level so other code can use them ====== */
function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h, s, l = (max + min) / 2;
  if (max === min) { h = s = 0; }
  else {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r: h = (g - b) / d + (g < b ? 6 : 0); break;
      case g: h = (b - r) / d + 2; break;
      case b: h = (r - g) / d + 4; break;
    }
    h /= 6;
  }
  return [Math.round(h * 360), s, l];
}

function getGradientFromBaseColor(hex, type = 'work') {
  if (!hex || hex[0] !== '#') return '';
  // Convert hex to RGB
  const rgb = parseInt(hex.slice(1), 16);
  const r = (rgb >> 16) & 255;
  const g = (rgb >> 8) & 255;
  const b = rgb & 255;
  const hsl = rgbToHsl(r, g, b);
  const [h, s, l] = hsl;

  if (type === 'work') {
    return `linear-gradient(135deg, hsl(${h}, ${Math.round(s * 100)}%, ${Math.max(25, Math.round(l * 100) - 10)}%), hsl(${h}, ${Math.round(s * 100)}%, ${Math.min(70, Math.round(l * 100) + 10)}%))`;
  } else {
    return `linear-gradient(135deg, hsl(${h}, ${Math.round(s * 100)}%, ${Math.min(90, Math.round(l * 100) + 20)}%), hsl(${h}, ${Math.round(s * 100)}%, ${Math.min(95, Math.round(l * 100) + 25)}%))`;
  }
}

            // --- COPY / PASTE SUPPORT ---
let copiedEmployeeSchedule = null;
let copiedEmployeeNo = null;
let copiedEmployeeData = null; // previously undeclared
            
            // FullCalendar instances
            let calendar;
            let draggable;
            
            // State for modals
            let currentDroppingEvent = null;
            let currentDeletingEvent = null;
            // track last mouse position for reliable key-based copy/paste target detection
            let lastMouseX = 0, lastMouseY = 0;
            
            // Configs
            const positionOptions = ['Branch Head', 'Site Supervisor', 'OIC', 'Mac Expert', 'Cashier'];
            const managerPositions = ['Branch Head', 'Site Supervisor', 'OIC'];
            const shiftPresets = [
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

            // --- DOM Elements ---
            const calendarEl = document.getElementById('calendar');
            let employeeTableBody = document.getElementById('employee-table-body');
            const draggableCardsContainer = document.getElementById('draggable-cards-container');
            const draggablePlaceholder = document.getElementById('draggable-placeholder');
            const conflictTableBody = document.getElementById('conflict-table-body');
            const conflictsPlaceholder = document.getElementById('conflicts-placeholder');
            
            // Modals
            const shiftModal = document.getElementById('shift-modal');
            const exportModal = document.getElementById('export-modal');
            const deleteModal = document.getElementById('delete-modal');
            
            // Modal Inputs/Buttons
            const shiftPresetSelect = document.getElementById('shift-preset');
            const shiftCustomInput = document.getElementById('shift-custom');
            const saveShiftBtn = document.getElementById('save-shift-btn');
            const confirmExportBtn = document.getElementById('confirm-export-btn');
            const confirmDeleteBtn = document.getElementById('confirm-delete-btn');
            const deleteModalSummary = document.getElementById('delete-modal-summary');

// === Shift Code Display With Time ===
const shiftTimes = {
  "RBG-001": "9:30AM - 6:30PM",
  "RBG-002": "10:00AM - 7:00PM",
  "RBG-003": "11:00AM - 8:00PM",
  "RBG-004": "12:00PM - 9:00PM",
  "RBG-005": "1:00PM - 10:00PM",
  "RBG-006": "2:00PM - 11:00PM",
  "RBG-007": "12:30PM - 9:30PM",
  "RBG-008": "1:30PM - 10:30PM",
  "RBG-009": "11:30AM - 8:30PM",
  "RBG-010": "10:30AM - 7:30PM",
  "RBG-011": "9:00AM - 6:00PM",
  "RBG-012": "8:30AM - 5:30PM",
  "RBG-013": "3:00PM - 12:00AM",
  "RBG-014": "3:30PM - 12:30AM",
  "RBG-015": "2:30PM - 11:30PM",
  "RBG-016": "8:30AM - 5:30PM",
  "RBG-017": "6:00AM - 3:00PM",
  "RBG-018": "7:00AM - 4:00PM",
  "RBG-019": "8:00AM - 5:00PM",
  "RBG-020": "7:30AM - 4:30PM",
  "RBG-021": "7:00AM - 5:00PM",
  "RBG-022": "11:00AM - 7:00PM",
  "RBG-023": "4:00PM - 1:00AM",
  "RBG-024": "8:00PM - 5:00AM",
  "RBG-025": "9:00PM - 6:00AM",
  "RBG-026": "6:30AM - 3:30PM",
  "RBG-027": "6:00PM - 3:00AM",
  "RBG-028": "6:30PM - 3:30AM",
  "RBG-029": "5:00PM - 2:00AM",
  "RBG-030": "7:00PM - 4:00AM",
  "RBG-031": "7:00AM - 4:00PM",
  "RBT-032": "6:00AM - 4:00PM",
  "RBT-033": "7:00AM - 5:00PM",
  "RBT-034": "8:00AM - 6:00PM",
  "RBT-035": "9:00AM - 7:00PM",
  "RBT-036": "9:30AM - 7:30PM",
  "RBT-037": "9:45AM - 7:45PM",
  "RBT-038": "10:00AM - 8:00PM",
  "RBT-039": "11:00AM - 9:00PM",
  "RBT-040": "12:00PM - 10:00PM",
  "RBT-041": "1:00PM - 11:00PM",
  "RBT-042": "2:00PM - 12:00AM",
  "RBT-043": "9:00PM - 7:00AM",
  "RBT-044": "10:00PM - 8:00AM",
  "RBT-045": "8:30AM - 6:30PM",
  "RBT-046": "11:30AM - 9:30PM",
  "RBT-047": "10:30AM - 8:30PM",
  "RBT-048": "11:30PM - 9:30PM",
  "RBT-049": "12:30PM - 10:30PM",
  "RBT-050": "7:30AM - 5:30PM",
  "RBT-051": "6:00PM - 4:00AM",
  "RBT-052": "5:00PM - 3:00AM",
  "RBT-053": "6:30AM - 4:30PM",
  "RBT-054": "3:00PM - 1:00AM",
  "AASP-001": "9:30AM - 6:30PM",
  "AASP-002": "10:00AM - 7:00PM",
  "AASP-003": "10:30AM - 7:30PM",
  "AASP-004": "11:00AM - 8:00PM",
  "AASP-005": "11:30AM - 8:30PM",
  "AASP-006": "9:00AM - 7:30PM",
  "AASP-007": "10:00AM - 7:00PM",
  "AASP-008": "9:30AM - 6:30PM",
  "AASP-009": "7:30AM - 4:30PM",
  "AASP-010": "12:30PM - 9:30PM",
  "AASP-011": "8:30AM - 5:30PM",
  "AASP-012": "9:30AM - 6:30PM",
  "AASP-013": "8:00AM - 6:30PM",
  "AASP-014": "9:00AM - 6:00PM",
  "AASP-015": "8:00AM - 5:00PM",
  "AASP-016": "6:00AM - 3:00PM",
  "AASP-017": "9:00AM - 6:00PM",
  "AASP-018": "6:00AM - 3:00PM",
  "AASP-019": "12:00PM - 9:00PM",
  "AASP-020": "6:00AM - 4:30PM",
  "AASP-021": "8:00AM - 5:00PM",
  "AASP-021 (10-5)": "10:00AM - 5:00PM",
  "AASP-022": "8:30AM - 7:00PM",
  "AASP-023": "4:30PM - 12:30AM",
  "AASP-024": "9:00AM - 7:00PM",
  "AASP-025": "6:00AM - 3:00PM",
  "AASP-026": "8:00PM - 5:00AM",
  "AASP-027": "7:00PM - 4:00AM",
  "AASP-028": "1:00PM - 10:00PM",
  "AASP-029": "8:30AM - 5:30PM",
  "AASP-030": "7:00AM - 4:00PM",
  "AASP-031": "10:00AM - 7:00PM",
  "AASP-032": "6:00AM - 3:00PM",
  "AASP-033": "2:00PM - 11:00PM",
  "AASP-034": "9:30AM - 7:30PM",
  "AASP-035": "10:00AM - 8:00PM",
  "AASP-036": "10:30AM - 8:30PM",
  "AASP-037": "11:00AM - 9:00PM",
  "AASP-038": "11:30AM - 9:30PM",
  "AASP-039": "12:00PM - 10:00PM",
  "AASP-040": "8:00AM - 6:00PM",
  "AASP-041": "9:00AM - 7:00PM",
  "AASP-042": "7:30AM - 5:30PM",
  "AASP-043": "8:30AM - 6:30PM",
  "AASP-044": "6:00AM - 4:00PM",
  "WHSE-001": "8:00AM - 5:00PM",
  "WHSE-002": "9:00AM - 6:00PM",
  "WHSE-003": "10:00AM - 7:00PM",
  "WHSE-004": "12:00PM - 9:00PM",
  "WHSE-005": "8:00AM - 12:00PM",
  "WHSE-006": "12:00PM - 4:00PM",
  "WHSE-007": "4:00PM - 1:00AM",
  "WHSE-008": "7:00AM - 4:00PM",
  "WHSE-009": "7:00AM - 11:00AM",
  "WHSE-010": "5:00AM - 2:00PM",
  "WHSE-011": "10:00AM - 2:00PM",
  "WHSE-012": "9:00AM - 1:00PM"
};

const shiftSelect = (typeof TomSelect !== 'undefined' && document.querySelector("#shift-preset")) ? new TomSelect("#shift-preset", {
  create: false,
  sortField: { field: "text", direction: "asc" },
  placeholder: "Search or select shift code...",

  render: {
    option: function (data, escape) {
      const code = data.value;
      const time = shiftTimes[code] || "";
      return `<div>${escape(code)} ${time ? `(${escape(time)})` : ""}</div>`;
    },
    item: function (data, escape) {
      const code = data.value;
      const time = shiftTimes[code] || "";
      return `<div>${escape(code)} ${time ? `(${escape(time)})` : ""}</div>`;
    }
  }
}) : null;

// Keep only the code value when saving/exporting
if (shiftSelect) {
  shiftSelect.on('change', function (value) {
    const cleanCode = value.split(' ')[0];
    if (shiftPresetSelect) shiftPresetSelect.value = cleanCode;
  });
}

            // Stats Bar
            const statsWork = document.getElementById('stats-work');
            const statsRest = document.getElementById('stats-rest');
            const statsConflicts = document.getElementById('stats-conflicts');
            const statsEmployees = document.getElementById('stats-employees');

            // Export Modal Stats
            const exportConflictWarning = document.getElementById('export-conflict-warning');
            const exportConflictCount = document.getElementById('export-conflict-count');
            const exportWorkCount = document.getElementById('export-work-count');
            const exportRestCount = document.getElementById('export-rest-count');


            // --- CORE INITIALIZATION ---

            // --- LOCAL STORAGE SAVE / LOAD ---

function saveToLocalStorage() {
  try {
    const allEvents = calendar ? calendar.getEvents().map(e => ({
      title: e.title,
      start: e.startStr,
      extendedProps: e.extendedProps
    })) : [];
    localStorage.setItem('employees', JSON.stringify(employees));
    localStorage.setItem('events', JSON.stringify(allEvents));
    localStorage.setItem('employeeColors', JSON.stringify(employeeColors));
  } catch (err) {
    console.warn('saveToLocalStorage error', err);
  }
}

function loadFromLocalStorage() {
  try {
    const storedEmployees = JSON.parse(localStorage.getItem('employees'));
    const storedEvents = JSON.parse(localStorage.getItem('events'));
  
    if (storedEmployees && Object.keys(storedEmployees).length > 0) {
      employees = storedEmployees;
      if (draggableCardsContainer) draggableCardsContainer.innerHTML = '';

      // ensure employeeColors is available (persisted earlier by saveToLocalStorage)
      employeeColors = JSON.parse(localStorage.getItem('employeeColors')) || employeeColors || {};

      const baseColors = [
        '#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899', '#14b8a6',
        '#ef4444', '#22c55e', '#eab308', '#0ea5e9', '#6366f1', '#84cc16',
        '#d946ef', '#0d9488', '#fb923c', '#a855f7', '#475569', '#f97316',
        '#64748b', '#60a5fa', '#65a30d', '#f43f5e', '#059669', '#7c3aed',
        '#e11d48', '#9333ea', '#2563eb', '#9ca3af', '#15803d'
      ];

      Object.values(employees).forEach(emp => {
        // assign or reuse a color (same logic used in saveAndGenerate)
        if (!employeeColors[emp.empNo]) {
          const used = Object.values(employeeColors);
          const available = baseColors.filter(c => !used.includes(c));
          const pick = available.length ? available[0] : baseColors[Object.keys(employeeColors).length % baseColors.length];
          employeeColors[emp.empNo] = pick;
        }
        const color = employeeColors[emp.empNo];

        const workEventData = {
          title: emp.name,
          extendedProps: { type: 'work', empNo: emp.empNo, position: emp.position }
        };
        const restEventData = {
          title: emp.name,
          extendedProps: { type: 'rest', empNo: emp.empNo, position: emp.position }
        };

        const workStyle = `background:${getGradientFromBaseColor(color, 'work')}; color:#fff; border:none;`;
        const restStyle = `background:${getGradientFromBaseColor(color, 'rest')}; color:${color}; border:2px solid ${color};`;

        const cardHtml = `
          <div class="p-3 rounded-lg shadow-sm border border-gray-200" data-empno="${emp.empNo}"
               style="border-left: 6px solid ${color}; background: linear-gradient(to right, ${color}22, ${color}08);">
            <div class="flex items-center justify-between">
              <div>
                <div class="font-medium text-gray-800">${emp.name}</div>
                <div class="text-xs text-gray-500">${emp.position}</div>
              </div>
              <div class="flex space-x-2">
                <div class='fc-event-pill fc-event-work px-3 py-1 text-xs font-medium rounded-full' data-event='${JSON.stringify(workEventData)}' style="${workStyle}">🟦 Work</div>
                <div class='fc-event-pill fc-event-rest px-3 py-1 text-xs font-medium rounded-full' data-event='${JSON.stringify(restEventData)}' style="${restStyle}">🔴 Rest</div>
              </div>
            </div>
          </div>`;
        if (draggableCardsContainer) draggableCardsContainer.innerHTML += cardHtml;
      });

      // persist employeeColors if we populated new values
      try { localStorage.setItem('employeeColors', JSON.stringify(employeeColors)); } catch (e) {}

      if (draggablePlaceholder) draggablePlaceholder.classList.add('hidden');
      initializeDraggable();
    }
  
    if (storedEvents && storedEvents.length > 0 && calendar) {
      storedEvents.forEach(ev => {
        calendar.addEvent({
          title: ev.title,
          start: ev.start,
          extendedProps: ev.extendedProps
        });
      });
      runConflictDetection();
      updateStats();
    }
  } catch (err) {
    console.warn('loadFromLocalStorage error', err);
  }
}
            
function initializeCalendar() {
  if (!calendarEl) return;

  calendar = new FullCalendar.Calendar(calendarEl, {
    initialView: 'dayGridMonth',
    editable: true,
    droppable: true,
    selectable: true,
    dayMaxEvents: true,
    height: 'auto',
    headerToolbar: {
      left: 'prev,next today',
      center: 'title',
      right: 'dayGridMonth,dayGridWeek,dayGridDay'
    },

    // ---------- Event handlers (preserved exactly as before) ----------
eventReceive: function(info) {

  // ✅ Always read employee number safely
  const droppedEmpNo = info.event.extendedProps?.empNo || info.event.extendedProps?.employeeNo;
  if (!droppedEmpNo) return;

  const newEvent = info.event;
  const { type } = newEvent.extendedProps;
  const dateStr = newEvent.startStr;

  // ✅ Get employee assigned color
  const color = employeeColors[droppedEmpNo];

  // ✅ Apply gradient + pill class based on type
  if (type === "work") {
    const grad = getGradientFromBaseColor(color, "work");
    newEvent.setProp("classNames", ["fc-event-pill", "fc-event-work"]);
    newEvent.setExtendedProp("forceStyle", {
      background: grad,
      backgroundImage: "none",
      color: "#fff",
      border: "none"
    });

  } else if (type === "rest") {
    const grad = getGradientFromBaseColor(color, "rest");
    newEvent.setProp("classNames", ["fc-event-pill", "fc-event-rest"]);
    newEvent.setExtendedProp("forceStyle", {
      background: grad,
      backgroundImage: "none",
      color: color,
      border: `2px solid ${color}`
    });
  }

  // ✅ Ensure unique ID
  try {
    newEvent.setExtendedProp('id', (crypto && crypto.randomUUID) ? crypto.randomUUID() : Date.now().toString());
  } catch (e) {
    newEvent.setExtendedProp('id', Date.now().toString());
  }

  // ✅ Duplicate check
  const allEvents = calendar.getEvents();
  const isDuplicate = allEvents.find(e => 
    e.extendedProps.id !== newEvent.extendedProps.id &&
    e.startStr === dateStr &&
    e.extendedProps.empNo === droppedEmpNo &&
    e.extendedProps.type === type
  );

  if (isDuplicate) {
    showToast(`Duplicate entry blocked: ${employees[droppedEmpNo] ? employees[droppedEmpNo].name : droppedEmpNo} already has a '${type}' day on this date.`, 'error');
    newEvent.remove();
    return;
  }

  // ✅ Work requires selecting a shift
  if (type === 'work') {
    currentDroppingEvent = newEvent;
    openShiftModal();
  } else {
    // ✅ Rest is auto-accepted
    runConflictDetection();
    updateStats();
    saveToLocalStorage();
  }

},

    /**
     * Fired when an event is dragged and dropped *within* the calendar.
     */
    eventDrop: function(info) {
      runConflictDetection();
      updateStats();
      saveToLocalStorage();
    },

    /**
     * Fired when an event is removed.
     */
    eventRemove: function(info) {
      runConflictDetection();
      updateStats();
      saveToLocalStorage();
    },

    /**
     * Fired when an event is clicked.
     */
    eventClick: function(info) {
      openDeleteModal(info.event);
    },

    /**
     * Fired after an event is rendered. Used for tooltips and styling.
     */
    eventDidMount: function(info) {
      // Apply forceStyle if present (ensures drag/drop keeps gradient)
      if (info.event.extendedProps.forceStyle) {
        Object.entries(info.event.extendedProps.forceStyle).forEach(([k,v]) => {
          info.el.style.setProperty(k, v, 'important');
        });
      }

      const { type, shiftCode, isConflict, empNo } = info.event.extendedProps;
      const emp = employees[empNo];

      // Make calendar event DOM queryable by empNo (enables contextmenu / keyboard copy-paste)
      if (empNo) {
        try { info.el.setAttribute('data-empno', empNo); } catch (e) {}
      }

      // Guard: if empNo is missing, avoid creating/assigning colors and hide the event early
      if (!empNo) {
        console.warn('eventDidMount: missing empNo for event', info.event && info.event.id);
        info.el.style.display = 'none';
        return;
      }

      // --- Persistent Unique Employee Color Assignment (solid vs border) ---
      const baseColors = [
        '#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899', '#14b8a6',
        '#ef4444', '#22c55e', '#eab308', '#0ea5e9', '#6366f1', '#84cc16',
        '#d946ef', '#0d9488', '#fb923c', '#a855f7', '#475569', '#f97316',
        '#64748b', '#60a5fa', '#65a30d', '#f43f5e', '#0ea5e9', '#059669',
        '#7c3aed', '#e11d48', '#9333ea', '#2563eb', '#9ca3af', '#15803d'
      ];

      if (!employeeColors[empNo]) {
        const usedColors = Object.values(employeeColors);
        const availableColors = baseColors.filter(c => !usedColors.includes(c));
        let color;
        if (availableColors.length) {
          color = availableColors[0];
        } else {
          // deterministic fallback when all base colors are used:
          // hash the employee number/string to pick a color index
          const key = String(empNo || '');
          let hash = 0;
          for (let i = 0; i < key.length; i++) {
            hash = ((hash << 5) - hash) + key.charCodeAt(i);
            hash |= 0; // convert to 32bit int
          }
          color = baseColors[Math.abs(hash) % baseColors.length];
        }
        employeeColors[empNo] = color;
        try { localStorage.setItem('employeeColors', JSON.stringify(employeeColors)); } catch (e) {}
      }
      const color = employeeColors[empNo];


      if (type === 'work') {
        // 🔸 Gradient-based background for Work type
        try {
          const grad = getGradientFromBaseColor(color, 'work');
          info.el.style.setProperty('background', grad, 'important');
          info.el.style.setProperty('color', '#fff', 'important');
          info.el.style.setProperty('border', 'none', 'important');

          const inner = info.el.querySelector('.fc-event-main-frame') || info.el;
          inner.style.setProperty('background', grad, 'important');
          inner.style.setProperty('color', '#fff', 'important');
        } catch (e) {
          // Fallback flat color
          info.el.style.backgroundColor = color;
          info.el.style.backgroundImage = 'none';
          info.el.style.color = '#fff';
          info.el.style.border = 'none';
          const inner = info.el.querySelector('.fc-event-main-frame') || info.el;
          inner.style.backgroundColor = color;
          inner.style.backgroundImage = 'none';
          inner.style.color = '#fff';
        }
      } else if (type === 'rest') {
        // 🔸 Gradient-based background for Rest type
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
          // Fallback flat color
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

      try {
        tippy(info.el, {
          content: `
            <div class='text-sm'>
              <div><strong>Employee #:</strong> ${emp ? emp.empNo : empNo}</div>
              <div><strong>Position:</strong> ${emp ? emp.position : 'N/A'}</div>
              <div><strong>Shift:</strong> ${type === 'work' ? 'Work' : 'Rest'}</div>
            </div>
          `,
          allowHTML: true,
          theme: 'light-border',
          placement: 'top',
        });
      } catch (e) {}

      if (!emp) {
        console.error(`Event ${info.event.id || ''} has invalid employee data (empNo: ${empNo}). Hiding event.`);
        info.el.style.display = 'none';
        return;
      }

      info.el.classList.add(type === 'work' ? 'fc-event-work' : 'fc-event-rest');

      if (isConflict) {
        info.el.classList.add('fc-event-conflict');
      } else {
        info.el.classList.remove('fc-event-conflict');
      }

      const shiftInfo = shiftCode ? ` (${shiftCode})` : '';
      const title = `${info.event.title} (${emp.position})\nType: ${type.charAt(0).toUpperCase() + type.slice(1)}${shiftInfo}`;
      info.el.title = title;
    },

    /**
     * Customizes the inner content of the event.
     */
    eventContent: function(arg) {
      const { shiftCode } = arg.event.extendedProps;
      const shiftHtml = shiftCode ? `<div class="text-xs opacity-90">${shiftCode}</div>` : '';
      return {
        html: `
          <div class="fc-event-main-frame">
            <div class="fc-event-title-container">
              <div class="fc-event-title fc-sticky">${arg.event.title}</div>
            </div>
            ${shiftHtml}
          </div>`
      };
    },

    /**
     * Style weekend days
     */
    dayCellClassNames: function(arg) {
      if (arg.date.getDay() === 0 || arg.date.getDay() === 6) {
        return 'fc-day-sat-sun';
      }
      return null;
    }
  });

  calendar.render();
}
            
            /**
             * Initializes the Draggable instance for the sidebar cards.
             */
            function initializeDraggable() {
                if (!draggableCardsContainer) return;
                if (draggable) {
                    try { draggable.destroy(); } catch (e) {}
                }

try {
  draggable = new FullCalendar.Draggable(draggableCardsContainer, {
    itemSelector: '.fc-event-pill',
    eventData: function (eventEl) {
      if (eventEl._fcEventData) return eventEl._fcEventData;
      try {
        const parsed = JSON.parse(eventEl.getAttribute('data-event'));
        const t = parsed.extendedProps && parsed.extendedProps.type ? parsed.extendedProps.type : '';
        parsed.classNames = t === 'rest' ? ['fc-event-rest'] : ['fc-event-work'];
        // Let CSS handle background color per type
        eventEl._fcEventData = parsed;
        return parsed;
      } catch (err) {
        console.error('Invalid event data on draggable element', err);
        return null;
      }
    }
  });
} catch (err) {
  console.warn('initializeDraggable error', err);
}

// ✅ Ensure drag alignment matches sidebar & calendar
calendar.on('eventDragStart', function(info) {
  info.el.style.transform = 'translateY(-4px)';
});
calendar.on('eventDragStop', function(info) {
  info.el.style.transform = 'translateY(0)';
});
            }


            
            // --- EMPLOYEE MANAGEMENT ---

            /**
             * Creates and appends a new blank employee row to the table.
             */
            function addEmployeeRow() {
                if (!employeeTableBody) {
                    console.warn('employee-table-body not found in DOM. Creating fallback table.');
                    // create a minimal table fallback so adding rows still works
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
                      </div>
                    `;
                    sidebar.insertBefore(wrapper, sidebar.firstChild);
                    // rebind the variable so later code sees the newly created table
                    employeeTableBody = document.getElementById('employee-table-body');
                }
                const tbody = document.getElementById('employee-table-body');
                if (!tbody) {
                    console.error('Unable to locate or create employee table body.');
                    return;
                }

                const tr = document.createElement('tr');
                tr.className = 'align-top';
                
                const positionDropdown = positionOptions
                    .map(pos => `<option value="${pos}">${pos}</option>`)
                    .join('');

                tr.innerHTML = `
                    <td class="p-1">
                        <input type="text" class="emp-name w-full text-sm border-gray-300 rounded-md shadow-sm focus:border-blue-500 focus:ring-blue-500" placeholder="J. Dela Cruz">
                    </td>
                    <td class="p-1">
                        <input type="text" class="emp-no w-full text-sm border-gray-300 rounded-md shadow-sm focus:border-blue-500 focus:ring-blue-500" placeholder="12345">
                    </td>
                    <td class="p-1">
                        <select class="emp-pos w-full text-sm border-gray-300 rounded-md shadow-sm focus:border-blue-500 focus:ring-blue-500">
                            <option value="">Select...</option>
                            ${positionDropdown}
                        </select>
                    </td>
                    <td class="p-1 text-center">
                        <button class="remove-row-btn p-1 text-red-400 hover:text-red-600 rounded-full" title="Remove row">
                            <svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" viewBox="0 0 20 20" fill="currentColor">
                                <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clip-rule="evenodd" />
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            }

            /**
             * Removes an employee row when its 'x' button is clicked.
             */
            function removeEmployeeRow(buttonEl) {
                if (!employeeTableBody) return;
                if (employeeTableBody.rows.length > 1) {
                    const tr = buttonEl.closest('tr');
                    if (tr) tr.remove();
                } else {
                    showToast('At least one employee row is required.', 'warn');
                }
            }
            
            /**
             * Validates employee table, saves to memory, and generates draggable cards.
             */
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
                    const noInput = row.querySelector('.emp-no');
                    const posInput = row.querySelector('.emp-pos');
                    
                    const name = nameInput ? nameInput.value.trim() : '';
                    const empNo = noInput ? noInput.value.trim() : '';
                    const position = posInput ? posInput.value : '';

                    if (!name && nameInput) { nameInput.classList.add('validation-error'); isValid = false; }
                    if (!empNo && noInput) { noInput.classList.add('validation-error'); isValid = false; }
                    if (!position && posInput) { posInput.classList.add('validation-error'); isValid = false; }
                    
                    if (empNo && employees[empNo]) {
                        if (noInput) noInput.classList.add('validation-error');
                        isValid = false;
                        showToast(`Duplicate Employee No: ${empNo}. Please use unique numbers.`, 'error');
                    }
                    
                    if (isValid && name && empNo && position) {
                        employees[empNo] = { name, empNo, position };
                    }
                });

                if (!isValid) {
                    showToast('Please fix validation errors in the employee table.', 'error');
                    return;
                }
                
                if (Object.keys(employees).length === 0) {
                    showToast('No valid employee data found. Please add at least one employee.', 'warn');
                    return;
                }

// --- Generate Cards with Persistent Colors ---
Object.values(employees).forEach(emp => {
  const workEventData = {
    title: emp.name,
    extendedProps: {
      type: 'work',
      empNo: emp.empNo,
      position: emp.position
    }
  };

  const restEventData = {
    title: emp.name,
    extendedProps: {
      type: 'rest',
      empNo: emp.empNo,
      position: emp.position
    }
  };

  // --- Assign or reuse color for employee ---
  const baseColors = [
    '#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899', '#14b8a6',
    '#ef4444', '#22c55e', '#eab308', '#0ea5e9', '#6366f1', '#84cc16',
    '#d946ef', '#0d9488', '#fb923c', '#a855f7', '#475569', '#f97316',
    '#64748b', '#60a5fa', '#65a30d', '#f43f5e', '#059669', '#7c3aed',
    '#e11d48', '#9333ea', '#2563eb', '#9ca3af', '#15803d'
  ];

  if (!employeeColors[emp.empNo]) {
    const usedColors = Object.values(employeeColors);
    const availableColors = baseColors.filter(c => !usedColors.includes(c));
    const color = availableColors.length
      ? availableColors[0]
      : baseColors[Object.keys(employeeColors).length % baseColors.length];
    employeeColors[emp.empNo] = color;
  }
  const color = employeeColors[emp.empNo];

  // --- Draggable Card UI ---
const cardHtml = `
  <div class="p-3 rounded-lg shadow-sm border border-gray-200" data-empno="${emp.empNo}"
       style="border-left: 6px solid ${color}; background: linear-gradient(to right, ${color}22, ${color}08);">
    <div class="flex items-center justify-between">
      <div>
        <div class="font-medium text-gray-800">${emp.name}</div>
        <div class="text-xs text-gray-500">${emp.position}</div>
      </div>
      <div class="flex space-x-2">
        <div
          class="fc-event-pill fc-event-work px-3 py-1 text-xs font-medium rounded-full cursor-pointer select-none shadow-sm"
          data-event='${JSON.stringify(workEventData)}'
          style="background:${getGradientFromBaseColor(color, 'work')}; color:#fff; border:none;">
          🟦 Work
        </div>
        <div
          class="fc-event-pill fc-event-rest px-3 py-1 text-xs font-medium rounded-full cursor-pointer select-none shadow-sm"
          data-event='${JSON.stringify(restEventData)}'
          style="background:${getGradientFromBaseColor(color, 'rest')}; color:${color}; border:2px solid ${color};">
          🔴 Rest
        </div>
      </div>
    </div>
  </div>
`;
  if (draggableCardsContainer) draggableCardsContainer.innerHTML += cardHtml;
});
                
                if (draggablePlaceholder) draggablePlaceholder.classList.add('hidden');
                
                initializeDraggable();
                
                updateStats();
                showToast('Employees saved and cards generated. You can now drag schedules.', 'success');
                saveToLocalStorage();

document.querySelectorAll('#draggable-cards-container > div').forEach(card => {
  const nameEl = card.querySelector('.font-medium');
  const name = nameEl ? nameEl.textContent.trim() : '';
  const emp = Object.values(employees).find(e => e.name === name);
  if (emp) card.dataset.empno = emp.empNo;
});

            }
            
            // --- MODAL & UI LOGIC ---

            function openShiftModal() {
                if (!shiftModal) return;
                if (shiftPresetSelect) shiftPresetSelect.value = '';
                if (shiftCustomInput) shiftCustomInput.value = '';

                if (shiftPresetSelect && shiftPresetSelect.options.length <= 1) {
                    shiftPresets.forEach(code => {
                        const option = new Option(code, code);
                        shiftPresetSelect.add(option);
                    });
                }

                if (window.lastUsedShiftCode && shiftSelect) {
                    shiftSelect.setValue(window.lastUsedShiftCode);
                }

                shiftModal.classList.remove('hidden');

                setTimeout(() => {
                    const tomInput = document.querySelector('.ts-dropdown .ts-input input, .ts-control input');
                    if (tomInput) tomInput.focus();
                }, 200);
            }


            // === Shift Search Filter ===
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
              try {
                preset = shiftSelect ? (shiftSelect.getValue() || shiftPresetSelect.value) : (shiftPresetSelect ? shiftPresetSelect.value : '');
              } catch (e) {
                preset = shiftPresetSelect ? shiftPresetSelect.value : '';
              }
            
              const custom = shiftCustomInput ? shiftCustomInput.value.trim() : '';
            
              if (!preset && !custom && window.lastUsedShiftCode) {
                preset = window.lastUsedShiftCode;
              }
            
              const shiftCode = custom || preset;
            
              if (!shiftCode) {
                showToast('Please select a preset or enter a custom shift code.', 'warn');
                return;
              }
            
              window.lastUsedShiftCode = shiftCode;
            
              const cleanCode = String(shiftCode).split(' ')[0];
              if (currentDroppingEvent) {
                currentDroppingEvent.setExtendedProp('shiftCode', cleanCode);
                currentDroppingEvent.setProp('classNames', ['fc-event-work']);
                runConflictDetection();
                updateStats();
                saveToLocalStorage();
              }
            
              if (shiftModal) closeModal(shiftModal);
              currentDroppingEvent = null;
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
                    saveToLocalStorage();
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

                if (exportWorkCount) exportWorkCount.textContent = workCount;
                if (exportRestCount) exportRestCount.textContent = restCount;

                if (exportConflictWarning && exportConflictCount) {
                    if (conflicts.length > 0) {
                        exportConflictCount.textContent = conflicts.length;
                        exportConflictWarning.classList.remove('hidden');
                    } else {
                        exportConflictWarning.classList.add('hidden');
                    }
                }
                
                exportModal.classList.remove('hidden');
            }

            function closeModal(modalEl) {
                if (!modalEl) return;
                modalEl.classList.add('hidden');
            }

            function handleCloseModal(modalId) {
                const modalEl = document.getElementById(modalId);
                closeModal(modalEl);

                if (modalId === 'shift-modal' && currentDroppingEvent) {
                    currentDroppingEvent.remove();
                    showToast('Schedule add canceled.', 'info');
                    currentDroppingEvent = null;
                }
                
                if (modalId === 'delete-modal') {
                    currentDeletingEvent = null;
                }
            }

            function showToast(message, type = 'info') {
                const container = document.getElementById('toast-container') || document.body;
                const toast = document.createElement('div');
                
                let bgColor, textColor, borderColor;
                switch (type) {
                    case 'success':
                        bgColor = 'bg-green-50'; textColor = 'text-green-800'; borderColor = 'border-green-200';
                        break;
                    case 'error':
                        bgColor = 'bg-red-50'; textColor = 'text-red-800'; borderColor = 'border-red-200';
                        break;
                    case 'warn':
                        bgColor = 'bg-yellow-50'; textColor = 'text-yellow-800'; borderColor = 'border-yellow-200';
                        break;
                    default:
                        bgColor = 'bg-blue-50'; textColor = 'text-blue-800'; borderColor = 'border-blue-200';
                }
                
                toast.className = `toast p-4 rounded-lg shadow-lg border ${bgColor} ${textColor} ${borderColor}`;
                toast.textContent = message;
                
                container.appendChild(toast);
                
                setTimeout(() => toast.classList.add('show'), 10);
                
                setTimeout(() => {
                    toast.classList.remove('show');
                    setTimeout(() => toast.remove(), 500);
                }, 4000);
            }
            
            
            // --- CONFLICT DETECTION ---

            function getGroupedEvents() {
                const allEvents = calendar ? calendar.getEvents() : [];
                const eventsByDate = {};
                const eventsByEmp = {};

                allEvents.forEach(event => {
                    const dateStr = event.startStr;
                    const { empNo } = event.extendedProps;
                    
                    if (!eventsByDate[dateStr]) eventsByDate[dateStr] = [];
                    eventsByDate[dateStr].push(event);
                    
                    if (!eventsByEmp[empNo]) eventsByEmp[empNo] = [];
                    eventsByEmp[empNo].push(event);
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
                                events.forEach(event => {
                                    conflicts.push({
                                        empNo,
                                        date,
                                        rule: 'Work & Rest on Same Day',
                                        event
                                    });
                                });
                            }
                        }
                    }
                }

                for (const date in eventsByDate) {
                    const managerRestEvents = eventsByDate[date].filter(event => {
                        const emp = employees[event.extendedProps.empNo];
                        return event.extendedProps.type === 'rest' &&
                               emp &&
                               managerPositions.includes(emp.position);
                    });

                    if (managerRestEvents.length > 1) {
                        managerRestEvents.forEach(event => {
                            conflicts.push({
                                empNo: event.extendedProps.empNo,
                                date,
                                rule: 'Multiple Managers Resting',
                                event
                            });
                        });
                    }
                }
                
                for (const empNo in eventsByEmp) {
                    const weekendRestEvents = eventsByEmp[empNo].filter(event => {
                        const day = event.start.getDay();
                        return event.extendedProps.type === 'rest' && (day === 0 || day === 6);
                    });

                    if (weekendRestEvents.length > 2) {
                        weekendRestEvents.forEach(event => {
                            conflicts.push({
                                empNo,
                                date: event.startStr,
                                rule: 'Exceeds 2 Weekend Rest Days',
                                event
                            });
                        });
                    }
                }
                
                return conflicts;
            }
            
            function runConflictDetection() {
                const { allEvents } = getGroupedEvents();
                
                if (allEvents && allEvents.forEach) {
                  allEvents.forEach(event => event.setExtendedProp('isConflict', false));
                }
                if (conflictTableBody) conflictTableBody.innerHTML = '';
                
                const conflicts = getConflicts();
                const conflictEvents = new Set();
                const conflictTableEntries = {};
                
                conflicts.forEach(conflict => {
                    conflict.event.setExtendedProp('isConflict', true);
                    conflictEvents.add(conflict.event.extendedProps.id);
                    
                    const key = `${conflict.empNo}-${conflict.rule}`;
                    if (!conflictTableEntries[key]) {
                        conflictTableEntries[key] = {
                            empNo: conflict.empNo,
                            rule: conflict.rule,
                            dates: new Set()
                        };
                    }
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
                            <td class="p-3 text-yellow-800">${[...entry.dates].join(', ')}</td>
                        `;
                        conflictTableBody.appendChild(tr);
                    });
                } else {
                    if (conflictsPlaceholder) conflictsPlaceholder.classList.remove('hidden');
                }
                
                updateStats(conflictEvents.size);
            }
            
            
            // --- CORE APP ACTIONS ---
            
            function updateStats(conflictCount = null) {
                const allEvents = calendar ? calendar.getEvents() : [];
                const workCount = allEvents.filter(e => e.extendedProps.type === 'work').length;
                const restCount = allEvents.filter(e => e.extendedProps.type === 'rest').length;
                
                if (conflictCount === null) {
                    conflictCount = allEvents.filter(e => e.extendedProps.isConflict).length;
                }
                
                const empCount = Object.keys(employees).length;

                if (statsWork) statsWork.textContent = workCount;
                if (statsRest) statsRest.textContent = restCount;
                if (statsConflicts) statsConflicts.textContent = conflictCount;
                if (statsEmployees) statsEmployees.textContent = empCount;
            }

            function exportToExcel() {
                if (!calendar) return;
                const conflicts = getConflicts();
                const allEvents = calendar.getEvents();
                const daysOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

                const wb = XLSX.utils.book_new();

                const workData = [
                    ['Name', 'Employee No', 'Position', 'Work Date', 'Shift Code', 'Day']
                ];
                allEvents
                    .filter(e => e.extendedProps.type === 'work')
                    .sort((a,b) => a.start - b.start || a.title.localeCompare(b.title))
                    .forEach(e => {
                        const emp = employees[e.extendedProps.empNo] || { name: 'N/A', empNo: e.extendedProps.empNo, position: 'N/A' };
                        workData.push([
                            emp.name,
                            emp.empNo,
                            emp.position,
                            e.startStr,
                            e.extendedProps.shiftCode || 'N/A',
                            daysOfWeek[e.start.getDay()]
                        ]);
                    });
                const wsWork = XLSX.utils.aoa_to_sheet(workData);
                XLSX.utils.book_append_sheet(wb, wsWork, 'Work Schedule');
                
                const restData = [
                    ['Name', 'Employee No', 'Position', 'Rest Date', 'Day']
                ];
                allEvents
                    .filter(e => e.extendedProps.type === 'rest')
                    .sort((a,b) => a.start - b.start || a.title.localeCompare(b.title))
                    .forEach(e => {
                        const emp = employees[e.extendedProps.empNo] || { name: 'N/A', empNo: e.extendedProps.empNo, position: 'N/A' };
                        restData.push([
                            emp.name,
                            emp.empNo,
                            emp.position,
                            e.startStr,
                            daysOfWeek[e.start.getDay()]
                        ]);
                    });
                const wsRest = XLSX.utils.aoa_to_sheet(restData);
                XLSX.utils.book_append_sheet(wb, wsRest, 'Rest Schedule');
                
                if (conflicts.length > 0) {
                    const conflictData = [
                        ['Employee Name', 'Employee No', 'Policy Violated', 'Dates Involved']
                    ];
                    const conflictTableEntries = {};
                    conflicts.forEach(conflict => {
                        const key = `${conflict.empNo}-${conflict.rule}`;
                        if (!conflictTableEntries[key]) {
                            conflictTableEntries[key] = {
                                empNo: conflict.empNo,
                                rule: conflict.rule,
                                dates: new Set()
                            };
                        }
                        conflictTableEntries[key].dates.add(conflict.date);
                    });
                    
                    Object.values(conflictTableEntries).forEach(entry => {
                        const emp = employees[entry.empNo] || { name: entry.empNo, empNo: entry.empNo };
                        conflictData.push([
                            emp.name,
                            emp.empNo,
                            entry.rule,
                            [...entry.dates].join(', ')
                        ]);
                    });
                    const wsConflicts = XLSX.utils.aoa_to_sheet(conflictData);
                    XLSX.utils.book_append_sheet(wb, wsConflicts, 'Conflict Summary');
                }
                
                XLSX.writeFile(wb, 'Branch_Schedule_Report.xlsx');
                if (exportModal) closeModal(exportModal);
                showToast('Excel report downloaded successfully!', 'success');
            }
            
            function resetAll() {
                if (!deleteModal || !deleteModalSummary || !confirmDeleteBtn) return;
                currentDeletingEvent = null;
                deleteModalSummary.textContent = "This will clear all employees, cards, and calendar data. This action cannot be undone.";
                deleteModal.querySelector('h3').textContent = "Reset All Data?";
                confirmDeleteBtn.textContent = "Confirm Reset";
                confirmDeleteBtn.classList.remove('bg-red-600', 'hover:bg-red-700', 'focus:ring-red-500');
                confirmDeleteBtn.classList.add('bg-orange-600', 'hover:bg-orange-700', 'focus:ring-orange-500');
                
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
                    confirmDeleteBtn.classList.add('bg-red-600', 'hover:bg-red-700', 'focus:ring-red-500');
                    confirmDeleteBtn.classList.remove('bg-orange-600', 'hover:bg-orange-700', 'focus:ring-orange-500');
                    confirmDeleteBtn.onclick = handleConfirmDelete;

                    showToast('Scheduler has been reset.', 'success');
                };
            }
            

            // --- EVENT LISTENERS ---
            const addRowBtnEl = document.getElementById('add-row-btn');
            if (addRowBtnEl) addRowBtnEl.addEventListener('click', addEmployeeRow);

            // Defensive fallback: also delegate clicks on document to ensure the button works even if an earlier error prevented the above binding
            document.addEventListener('click', function (e) {
              const btn = e.target.closest && e.target.closest('#add-row-btn');
              if (btn) {
                try { addEmployeeRow(); } catch (err) { console.error('addEmployeeRow error (fallback):', err); }
              }
            });

            // Expose for debugging / REPL use
            try { window.addEmployeeRow = addEmployeeRow; } catch(e){}

            const saveGenerateBtnEl = document.getElementById('save-generate-btn');
            if (saveGenerateBtnEl) saveGenerateBtnEl.addEventListener('click', saveAndGenerate);
            if (employeeTableBody) {
              employeeTableBody.addEventListener('click', function(e) {
                  const removeBtn = e.target.closest('.remove-row-btn');
                  if (removeBtn) {
                      removeEmployeeRow(removeBtn);
                  }
              });
            }

            const exportBtnEl = document.getElementById('export-btn');
            if (exportBtnEl) exportBtnEl.addEventListener('click', openExportModal);
            const resetBtnEl = document.getElementById('reset-btn');
            if (resetBtnEl) resetBtnEl.addEventListener('click', resetAll);
            
            if (saveShiftBtn) saveShiftBtn.addEventListener('click', handleSaveShift);
            if (confirmExportBtn) confirmExportBtn.addEventListener('click', exportToExcel);
            if (confirmDeleteBtn) confirmDeleteBtn.addEventListener('click', handleConfirmDelete);
            
            document.querySelectorAll('[data-modal-close]').forEach(btn => {
                btn.addEventListener('click', () => handleCloseModal(btn.dataset.modalClose));
            });
            
            document.querySelectorAll('.modal-backdrop').forEach(backdrop => {
                backdrop.addEventListener('click', function(e) {
                    if (e.target === this) {
                        handleCloseModal(this.id);
                    }
                });
            });

            document.addEventListener('keydown', function (e) {
                if (e.key === 'Escape') {
                    if (shiftModal && !shiftModal.classList.contains('hidden')) handleCloseModal('shift-modal');
                    if (exportModal && !exportModal.classList.contains('hidden')) handleCloseModal('export-modal');
                    if (deleteModal && !deleteModal.classList.contains('hidden')) handleCloseModal('delete-modal');
                }
            });


// --- CONTEXT MENU + COPY/PASTE LOGIC ---
if (!window.__schedulerContextMenuInit) {
  window.__schedulerContextMenuInit = true;

  // Track mouse position for keyboard copy/paste detection
  let lastMouseX = 0, lastMouseY = 0;
  document.addEventListener('mousemove', function(e) {
    lastMouseX = e.clientX;
    lastMouseY = e.clientY;
  });

  document.addEventListener('contextmenu', function (e) {
    const el = document.elementFromPoint(e.clientX, e.clientY);
    if (!el) return;

    const empTarget = el.closest('[data-empno]');
    const calendarEl = el.closest('.fc-daygrid-day, .fc-timegrid-slot'); // FullCalendar day cells

    if (empTarget) {
      // Right-click on employee card → show Copy/Paste menu
      e.preventDefault();
      e.stopPropagation();
      const empNo = empTarget.dataset.empno;
      if (!empNo) return;
      showContextMenu(e.pageX, e.pageY, empNo);
    } else if (calendarEl && copiedEmployeeData) {
      // Right-click on calendar day → show Paste-to-date menu
      e.preventDefault();
      e.stopPropagation();
      const dateStr = calendarEl.getAttribute('data-date');
      if (!dateStr) return;
      showCalendarPasteMenu(e.pageX, e.pageY, dateStr);
    }
  });

  function showContextMenu(x, y, empNo) {
    document.querySelector('#context-menu')?.remove();
    const menu = document.createElement('div');
    menu.id = 'context-menu';
    menu.className = 'absolute bg-white border border-gray-300 rounded shadow-lg z-50';
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    menu.style.minWidth = '160px';
    menu.innerHTML = `
      <button id="copy-schedule" class="block w-full text-left px-4 py-2 hover:bg-gray-100">📋 Copy Schedule</button>
      <button id="paste-schedule" class="block w-full text-left px-4 py-2 hover:bg-gray-100">📥 Paste Schedule</button>
    `;
    document.body.appendChild(menu);

    document.getElementById('copy-schedule').onclick = () => { copyEmployeeSchedule(empNo); menu.remove(); };
    document.getElementById('paste-schedule').onclick = () => { pasteEmployeeSchedule(empNo); menu.remove(); };

    const remover = () => {
      menu.remove();
      document.removeEventListener('keydown', keyHandler);
      document.removeEventListener('click', remover);
    };
    const keyHandler = (ev) => { if (ev.key === 'Escape') remover(); };
    document.addEventListener('click', remover, { once: true });
    document.addEventListener('keydown', keyHandler);
  }

  function showCalendarPasteMenu(x, y, dateStr) {
    document.querySelector('#context-menu')?.remove();
    const menu = document.createElement('div');
    menu.id = 'context-menu';
    menu.className = 'absolute bg-white border border-gray-300 rounded shadow-lg z-50';
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    menu.innerHTML = `
      <button id="paste-schedule-calendar" class="block w-full text-left px-4 py-2 hover:bg-gray-100">
        📅 Paste Schedule Here (${dateStr})
      </button>
    `;
    document.body.appendChild(menu);

    document.getElementById('paste-schedule-calendar').onclick = () => {
      pasteCopiedScheduleToDate(dateStr);
      menu.remove();
    };

    const remover = () => {
      menu.remove();
      document.removeEventListener('keydown', keyHandler);
      document.removeEventListener('click', remover);
    };
    const keyHandler = (ev) => { if (ev.key === 'Escape') remover(); };
    document.addEventListener('click', remover, { once: true });
    document.addEventListener('keydown', keyHandler);
  }

  function copyEmployeeSchedule(empNo) {
    const emp = employees[empNo];
    if (!emp) {
      showToast('Employee not found.', 'error');
      copiedEmployeeData = null;
      return;
    }

    const events = calendar ? calendar.getEvents().filter(e => e.extendedProps && e.extendedProps.empNo === empNo) : [];

    copiedEmployeeData = {
      empNo: emp.empNo,
      name: emp.name,
      events: events.map(e => ({
        type: e.extendedProps.type,
        start: e.startStr,
        shiftCode: e.extendedProps.shiftCode || null
      }))
    };

    showToast(`Copied schedule for ${emp.empNo} - ${emp.name} (${copiedEmployeeData.events.length} events)`, 'info');
  }

  function pasteEmployeeSchedule(targetEmpNo) {
    if (!copiedEmployeeData) {
      showToast('Nothing copied yet.', 'error');
      return;
    }

    const target = employees[targetEmpNo];
    if (!target) {
      showToast('Target employee not found.', 'error');
      return;
    }

    if (!calendar) {
      showToast('Calendar not initialized.', 'error');
      return;
    }

    const srcEvents = copiedEmployeeData.events || [];
    if (srcEvents.length === 0) {
      showToast('No events to paste.', 'warn');
      return;
    }

    let added = 0;
    let skipped = 0;

    srcEvents.forEach(ev => {
      const exists = calendar.getEvents().some(e =>
        e.extendedProps &&
        e.extendedProps.empNo === targetEmpNo &&
        e.startStr === ev.start &&
        e.extendedProps.type === ev.type
      );

      if (exists) {
        skipped++;
        return;
      }

      const newExtended = {
        type: ev.type,
        empNo: targetEmpNo,
        position: target.position,
        shiftCode: ev.shiftCode || undefined,
        id: (crypto?.randomUUID?.() ?? `copied-${Date.now()}-${Math.random().toString(36).slice(2,8)}`)
      };

      calendar.addEvent({
        title: target.name,
        start: ev.start,
        extendedProps: newExtended
      });
      added++;
    });

    runConflictDetection();
    updateStats();
    saveToLocalStorage();

    showToast(`Pasted to ${targetEmpNo}: ${added} added, ${skipped} skipped (duplicates).`, added ? 'success' : 'warn');
  }

  // ✅ FIXED: Now global — works when right-clicking on the calendar
  function pasteCopiedScheduleToDate(targetDateStr) {
    if (!copiedEmployeeData || !calendar) {
      showToast('No copied schedule available.', 'error');
      return;
    }

    const targetDate = new Date(targetDateStr);
    if (isNaN(targetDate)) {
      showToast('Invalid date target.', 'error');
      return;
    }

    const empNo = copiedEmployeeData.empNo;
    const emp = employees[empNo];
    if (!emp) {
      showToast('Original employee data missing.', 'error');
      return;
    }

    let added = 0;
    let skipped = 0;

    const srcEvents = copiedEmployeeData.events || [];
    if (srcEvents.length === 0) {
      showToast('No events to paste.', 'warn');
      return;
    }

    const firstSrcDate = new Date(srcEvents[0].start);
    const offsetDays = Math.round((targetDate - firstSrcDate) / (1000 * 60 * 60 * 24));

    srcEvents.forEach(ev => {
      const newDate = new Date(ev.start);
      newDate.setDate(newDate.getDate() + offsetDays);
      const newDateStr = newDate.toISOString().split('T')[0];

      const exists = calendar.getEvents().some(e =>
        e.extendedProps.empNo === empNo &&
        e.startStr === newDateStr &&
        e.extendedProps.type === ev.type
      );
      if (exists) { skipped++; return; }

      const newExtended = {
        type: ev.type,
        empNo: empNo,
        position: emp.position,
        shiftCode: ev.shiftCode || undefined,
        id: (crypto?.randomUUID?.() ?? `pasted-${Date.now()}-${Math.random().toString(36).slice(2,8)}`)
      };

      calendar.addEvent({
        title: emp.name,
        start: newDateStr,
        extendedProps: newExtended
      });
      added++;
    });

    runConflictDetection();
    updateStats();
    saveToLocalStorage();

    showToast(`Pasted ${added} events starting ${targetDateStr} (${skipped} skipped)`, added ? 'success' : 'warn');
  }

  // Keyboard copy/paste shortcuts (Ctrl+C / Ctrl+V)
  document.addEventListener('keydown', (e) => {
    const tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : '';
    if (tag === 'input' || tag === 'textarea' || e.target?.isContentEditable) return;
    if (!(e.ctrlKey || e.metaKey)) return;

    const key = (e.key || '').toLowerCase();
    if (key !== 'c' && key !== 'v') return;

    const el = (typeof lastMouseX === 'number' && typeof lastMouseY === 'number')
      ? document.elementFromPoint(lastMouseX, lastMouseY)
      : document.activeElement;
    if (!el) return;

    const target = el.closest?.('[data-empno]');
    if (!target) return;

    const empNo = target.dataset.empno;
    if (!empNo) return;

    e.preventDefault();
    if (key === 'c') copyEmployeeSchedule(empNo);
    if (key === 'v') pasteEmployeeSchedule(empNo);
  });

  // --- Initialization calls ---
  initializeCalendar();
  initializeDraggable();
  addEmployeeRow();
  loadFromLocalStorage();
  updateStats();
} 
});
