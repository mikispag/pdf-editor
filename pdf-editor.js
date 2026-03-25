// ==========================================
// PDF Editor
// ==========================================
import { PDFDocument, rgb, StandardFonts, degrees } from './vendor/pdf-lib.esm.js';
import * as pdfjsLib from './vendor/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = './vendor/pdf.worker.min.mjs';

// --- Color palette ---
const COLORS = {
    black:  rgb(0, 0, 0),
    blue:   rgb(0.145, 0.388, 0.921),
    red:    rgb(0.937, 0.266, 0.266),
    green:  rgb(0.133, 0.772, 0.368),
    yellow: rgb(1, 1, 0), // highlight tool only
    white:  rgb(1, 1, 1),
};

const HIGHLIGHT_COLORS = {
    yellow:  rgb(1, 1, 0),
    fuchsia: rgb(1, 0.4, 0.8),
    cyan:    rgb(0, 0.9, 1),
    lime:    rgb(0.6, 1, 0),
    orange:  rgb(1, 0.65, 0),
};

const DEFAULT_PRESETS = [
    { key: 'black', label: 'Black', css: '#000000' },
    { key: 'blue', label: 'Blue', css: '#2563eb' },
    { key: 'red', label: 'Red', css: '#ef4444' },
    { key: 'green', label: 'Green', css: '#22c55e' },
    { key: 'white', label: 'White', css: '#ffffff' },
];

const HIGHLIGHT_PRESETS = [
    { key: 'yellow', label: 'Yellow', css: '#ffff00' },
    { key: 'fuchsia', label: 'Fuchsia', css: '#ff66cc' },
    { key: 'cyan', label: 'Cyan', css: '#00e5ff' },
    { key: 'lime', label: 'Lime', css: '#99ff00' },
    { key: 'orange', label: 'Orange', css: '#ffa600' },
];

// --- State ---
let existingPdfBytes = null;
let currentTool = 'select';
let annotations = [];
let scale = 1.5;
const ZOOM_MIN = 0.5;
const ZOOM_MAX = 3.0;
const ZOOM_STEP = 0.25;
const DEFAULT_SCALE = 1.5;
let pages = [];
let activePageIndex = 0;
let selectedIndex = -1;
let editingIndex = -1;
let currentColor = COLORS.black;
let isFillMode = true;

// Drawing state
let isDrawing = false;
let isDragging = false;
let isResizing = false;
let startX, startY;
let currentPath = [];
let dragOffsetX, dragOffsetY;
let drawRAF = null;
let commitTimeout = null;
let resizeOrigin = null; // Captures annotation state at resize start
let eraserSnapshotPushed = false;

const RESIZE_HANDLE_SIZE = 10;

let toggleShortcutPanel = null;
let pageIndicator = null;
let zoomInProgress = false;
let undoRedoInProgress = false;

// --- Coordinate conversion helpers ---
// Annotations are stored in PDF-space (points).
// Canvas/mouse coordinates are in pixel-space (points * scale).
function canvasToPdf(val) { return val / scale; }
function pdfToCanvas(val) { return val * scale; }

// New state
let currentThickness = 2;
let currentOpacity = 1.0;
let currentFontSize = 16;
let currentFontFamily = 'Helvetica';
const FONT_CANVAS_MAP = {
    'Helvetica': 'Helvetica, Arial, sans-serif',
    'TimesRoman': "'Times New Roman', Times, serif",
    'Courier': "'Courier New', Courier, monospace",
};
let signatureDataUrl = null;
let signatureImg = null;
const imageCache = new Map();
const tintedSigCache = new Map();
let pendingStampText = null;
let cachedPrimaryColor = '#2563eb';
let activeNotePopup = null;
let findMatches = [];
let findCurrentIndex = -1;
let findQuery = '';
let cachedTextContent = null; // Cached pdf.js text content per page
let searchGeneration = 0;     // Guards against overlapping searches
let formOverlays = []; // { element, fieldName, fieldType, pageIndex }
let cropRect = null; // { pageIndex, x, y, w, h } in PDF-space, null when not cropping

let pageCleanups = [];
let pageRotations = [];
let pageOrder = [];
let undoStack = [];
let redoStack = [];
let isShiftHeld = false;
const MAX_UNDO = 50;

// DOM references (lazy init)
let pagesContainer, textInput, emptyState, deleteBtn, fillToggle, fileInput, thumbnailList, splitModal;

// Cached query results
let toolBtns = null;
let colorCircles = null;

// Color picker DOM refs
let colorPicker, colorHexInput;

// Context bar DOM refs
let ctxThickness, ctxOpacity, ctxFontsize, ctxFontfamily, ctxFill, ctxColor;
let thicknessSlider, thicknessLabel, opacitySlider, opacityLabel, fontSizeInput, fontFamilySelect;
let undoBtnRef, redoBtnRef;

function cacheToolbar() {
    toolBtns = document.querySelectorAll('.sidebar-tool-btn');
    colorCircles = document.querySelectorAll('.color-circle');
    colorPicker = document.getElementById('pdf-color-picker');
    colorHexInput = document.getElementById('pdf-color-hex');

    // Context bar containers
    ctxThickness = document.getElementById('ctx-thickness');
    ctxOpacity = document.getElementById('ctx-opacity');
    ctxFontsize = document.getElementById('ctx-fontsize');
    ctxFontfamily = document.getElementById('ctx-fontfamily');
    ctxFill = document.getElementById('ctx-fill');
    ctxColor = document.getElementById('ctx-color');

    // Context bar controls
    thicknessSlider = document.getElementById('pdf-thickness');
    thicknessLabel = document.getElementById('pdf-thickness-val');
    opacitySlider = document.getElementById('pdf-opacity');
    opacityLabel = document.getElementById('pdf-opacity-val');
    fontSizeInput = document.getElementById('pdf-font-size');
    fontFamilySelect = document.getElementById('pdf-font-family');

    // Undo/redo buttons
    undoBtnRef = document.getElementById('pdf-btn-undo');
    redoBtnRef = document.getElementById('pdf-btn-redo');
}

// ==========================================
// Theme Management
// ==========================================
function initTheme() {
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)');

    function applyTheme(theme, animate = false) {
        if (animate) {
            document.documentElement.setAttribute('data-theme-transitioning', '');
            setTimeout(() => document.documentElement.removeAttribute('data-theme-transitioning'), 300);
        }
        document.documentElement.setAttribute('data-theme', theme);
        const sunIcon = document.getElementById('theme-icon-sun');
        const moonIcon = document.getElementById('theme-icon-moon');
        if (sunIcon) sunIcon.classList.toggle('hidden', theme === 'dark');
        if (moonIcon) moonIcon.classList.toggle('hidden', theme !== 'dark');
        // Update cached primary color for canvas rendering
        requestAnimationFrame(() => {
            cachedPrimaryColor = getComputedStyle(document.documentElement).getPropertyValue('--primary').trim() || '#2563eb';
        });
    }

    function getEffectiveTheme() {
        const stored = localStorage.getItem('theme');
        if (stored) return stored;
        return prefersDark.matches ? 'dark' : 'light';
    }

    applyTheme(getEffectiveTheme(), false);

    prefersDark.addEventListener('change', () => {
        if (!localStorage.getItem('theme')) {
            applyTheme(prefersDark.matches ? 'dark' : 'light', true);
        }
    });

    const toggleBtn = document.getElementById('pdf-btn-theme');
    if (toggleBtn) {
        toggleBtn.addEventListener('click', () => {
            const current = document.documentElement.getAttribute('data-theme');
            const next = current === 'dark' ? 'light' : 'dark';
            localStorage.setItem('theme', next);
            applyTheme(next, true);
        });
    }
}

// ==========================================
// Focus Trap Utility
// ==========================================
function trapFocus(container) {
    const focusable = container.querySelectorAll(
        'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
    );
    const first = focusable[0];
    const last = focusable[focusable.length - 1];

    function handler(e) {
        if (e.key !== 'Tab') return;
        if (e.shiftKey) {
            if (document.activeElement === first) {
                e.preventDefault();
                last.focus();
            }
        } else {
            if (document.activeElement === last) {
                e.preventDefault();
                first.focus();
            }
        }
    }

    container.addEventListener('keydown', handler);
    if (first) first.focus();
    return () => container.removeEventListener('keydown', handler);
}

// ==========================================
// Tooltip System
// ==========================================
function initTooltips() {
    const tooltip = document.getElementById('tooltip');
    if (!tooltip) return;

    const isMac = /mac/i.test(navigator.userAgentData?.platform || navigator.platform);
    let hoverTimer = null;
    let currentTarget = null;

    function show(target) {
        const text = target.getAttribute('data-tooltip');
        if (!text) return;
        let shortcut = target.getAttribute('data-tooltip-shortcut') || '';
        if (shortcut && isMac) {
            shortcut = shortcut.replace(/Ctrl\+/g, 'Cmd+');
        }

        tooltip.textContent = text;
        if (shortcut) {
            const kbd = document.createElement('kbd');
            kbd.textContent = shortcut;
            tooltip.appendChild(kbd);
        }

        // Position below target, centered
        const rect = target.getBoundingClientRect();
        tooltip.style.left = '0';
        tooltip.style.top = '0';
        tooltip.classList.add('visible');

        const tw = tooltip.offsetWidth;
        const th = tooltip.offsetHeight;
        let left = rect.left + rect.width / 2 - tw / 2;
        let top = rect.bottom + 6;

        // Flip above if near viewport bottom
        if (top + th > window.innerHeight - 8) {
            top = rect.top - th - 6;
        }

        // Clamp to viewport edges
        left = Math.max(4, Math.min(left, window.innerWidth - tw - 4));

        tooltip.style.left = left + 'px';
        tooltip.style.top = top + 'px';
        currentTarget = target;
    }

    function hide() {
        clearTimeout(hoverTimer);
        hoverTimer = null;
        tooltip.classList.remove('visible');
        currentTarget = null;
    }

    document.addEventListener('pointerover', (e) => {
        const target = e.target.closest('[data-tooltip]');
        if (!target) return;
        if (target === currentTarget) return;
        hide();
        hoverTimer = setTimeout(() => show(target), 400);
    });

    document.addEventListener('pointerout', (e) => {
        const target = e.target.closest('[data-tooltip]');
        if (target) hide();
    });

    document.addEventListener('pointerdown', () => hide());
}

// ==========================================
// Keyboard Shortcut Panel
// ==========================================
function initShortcutPanel() {
    const overlay = document.getElementById('shortcut-panel-overlay');
    if (!overlay) return;

    const isMac = /mac/i.test(navigator.userAgentData?.platform || navigator.platform);
    const mod = isMac ? 'Cmd' : 'Ctrl';

    const toolsList = document.getElementById('shortcut-list-tools');
    const actionsList = document.getElementById('shortcut-list-actions');
    const navList = document.getElementById('shortcut-list-nav');

    const tools = [
        ['Select', 'V'], ['Rectangle', 'R'], ['Ellipse', 'O'], ['Line', 'L'],
        ['Arrow', 'A'], ['Pen', 'P'], ['Highlight', 'H'], ['Eraser', 'E'], ['Redact', 'D'],
        ['Text', 'T'], ['Note', 'N'], ['Signature', 'S'], ['Stamp', '-'],
    ];
    const actions = [
        ['Undo', mod + '+Z'], ['Redo', mod + '+Shift+Z'],
        ['Delete', 'Del'], ['Download', mod + '+S'],
        ['Find', mod + '+F'],
    ];
    const nav = [
        ['Shortcuts', '?'],
        ['Zoom in', '+'], ['Zoom out', '-'],
        ['Reset zoom', mod + '+0'],
    ];

    function renderList(container, items) {
        items.forEach(([label, key]) => {
            const row = document.createElement('div');
            row.className = 'shortcut-row';
            const span = document.createElement('span');
            span.textContent = label;
            const kbd = document.createElement('kbd');
            kbd.textContent = key;
            row.appendChild(span);
            row.appendChild(kbd);
            container.appendChild(row);
        });
    }

    renderList(toolsList, tools);
    renderList(actionsList, actions);
    renderList(navList, nav);

    let releaseTrap = null;

    function toggle() {
        const isHidden = overlay.classList.contains('hidden');
        if (isHidden) {
            overlay.classList.remove('hidden');
            overlay.classList.add('active');
            releaseTrap = trapFocus(overlay.querySelector('.shortcut-panel'));
        } else {
            close();
        }
    }

    function close() {
        overlay.classList.add('hidden');
        overlay.classList.remove('active');
        if (releaseTrap) { releaseTrap(); releaseTrap = null; }
    }

    overlay.addEventListener('click', (e) => {
        if (e.target === overlay) close();
    });

    overlay.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' || e.key === '?') {
            e.preventDefault();
            close();
        }
    });

    toggleShortcutPanel = toggle;
}

// ==========================================
// Page Indicator
// ==========================================
function initPageIndicator() {
    const indicator = document.getElementById('page-indicator');
    const currentSpan = document.getElementById('page-indicator-current');
    const totalSpan = document.getElementById('page-indicator-total');
    if (!indicator || !currentSpan || !totalSpan) return;

    let observer = null;

    const ratioMap = new Map();

    function setupObserver() {
        if (observer) observer.disconnect();
        ratioMap.clear();
        const scrollEl = document.getElementById('pdf-pages-scroll');
        if (!scrollEl || pages.length === 0) return;

        observer = new IntersectionObserver((entries) => {
            for (const entry of entries) {
                const idx = pages.findIndex(p => p.wrapper === entry.target);
                if (idx !== -1) ratioMap.set(idx, entry.intersectionRatio);
            }
            let bestIdx = -1, bestRatio = 0;
            for (const [idx, ratio] of ratioMap) {
                if (ratio > bestRatio) { bestRatio = ratio; bestIdx = idx; }
            }
            if (bestIdx !== -1) {
                currentSpan.textContent = bestIdx + 1;
            }
        }, {
            root: scrollEl,
            threshold: [0, 0.25, 0.5, 0.75, 1],
        });

        for (const p of pages) {
            observer.observe(p.wrapper);
        }
    }

    function updateTotal() {
        totalSpan.textContent = pages.length;
    }

    // Click on current page number to jump
    currentSpan.addEventListener('click', () => {
        const input = document.createElement('input');
        input.type = 'number';
        input.className = 'page-indicator-input';
        input.min = 1;
        input.max = pages.length;
        input.value = currentSpan.textContent;
        currentSpan.classList.add('hidden');
        currentSpan.parentNode.insertBefore(input, currentSpan.nextSibling);
        input.focus();
        input.select();

        function commit() {
            const val = parseInt(input.value);
            if (val >= 1 && val <= pages.length) {
                pages[val - 1].wrapper.scrollIntoView({ behavior: 'smooth', block: 'center' });
                currentSpan.textContent = val;
            }
            cleanup();
        }

        function cleanup() {
            currentSpan.classList.remove('hidden');
            if (input.parentNode) input.parentNode.removeChild(input);
        }

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); commit(); }
            if (e.key === 'Escape') { e.preventDefault(); cleanup(); }
        });
        input.addEventListener('blur', () => commit());
    });

    pageIndicator = { setupObserver, updateTotal };
}

// ==========================================
// Initialization
// ==========================================
function init() {
    initTheme();
    initTooltips();
    initShortcutPanel();
    initPageIndicator();
    initFindBar();

    pagesContainer = document.getElementById('pdf-pages-container');
    textInput      = document.getElementById('pdf-text-input');
    emptyState     = document.getElementById('pdf-empty-state');
    deleteBtn      = document.getElementById('pdf-btn-delete');
    fillToggle     = document.getElementById('pdf-fill-toggle');
    fileInput      = document.getElementById('pdf-file-input');
    thumbnailList = document.getElementById('pdf-thumbnail-list');
    splitModal = document.getElementById('split-modal');

    const collapseBtn = document.getElementById('pdf-collapse-sidebar');
    if (collapseBtn) {
        collapseBtn.addEventListener('click', () => {
            const sidebar = document.querySelector('.pdf-right-sidebar');
            const expandBtn = document.getElementById('pdf-expand-sidebar');
            if (sidebar) {
                sidebar.classList.add('collapsed');
                if (expandBtn) expandBtn.classList.remove('hidden');
            }
        });
    }

    cacheToolbar();

    if (fileInput) {
        fileInput.addEventListener('change', handleFileUpload);
    }

    // Tool buttons (sidebar)
    const toolMap = {
        'pdf-btn-select': () => setTool('select'),
        'pdf-btn-delete': () => deleteSelected(),
        'pdf-btn-rect': () => setTool('rect'),
        'pdf-btn-ellipse': () => setTool('ellipse'),
        'pdf-btn-line': () => setTool('line'),
        'pdf-btn-arrow': () => setTool('arrow'),
        'pdf-btn-pen': () => setTool('pen'),
        'pdf-btn-highlight': () => setTool('highlight'),
        'pdf-btn-eraser': () => setTool('eraser'),
        'pdf-btn-redact': () => setTool('redact'),
        'pdf-btn-crop': () => setTool('crop'),
        'pdf-btn-text': () => setTool('text'),
        'pdf-btn-note': () => setTool('note'),
        'pdf-btn-signature': () => setTool('signature'),
        'pdf-btn-stamp': () => setTool('stamp'),
        'pdf-btn-undo': () => undo(),
        'pdf-btn-redo': () => redo(),
        'pdf-btn-download': () => downloadPDF(),
        'pdf-btn-merge': () => mergePDF(),
        'pdf-btn-split': () => showSplitModal(),
        'pdf-btn-split-cancel': () => hideSplitModal(),
        'pdf-btn-split-confirm': () => splitPDF(),
        'pdf-btn-zoom-in': () => setZoom(scale + ZOOM_STEP),
        'pdf-btn-zoom-out': () => setZoom(scale - ZOOM_STEP),
    };
    for (const [id, handler] of Object.entries(toolMap)) {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', handler);
    }

    // Crop toolbar buttons
    document.getElementById('crop-apply')?.addEventListener('click', applyCrop);
    document.getElementById('crop-cancel')?.addEventListener('click', cancelCrop);

    // Color presets
    document.querySelectorAll('.color-circle[data-color]').forEach(circle => {
        circle.addEventListener('click', () => setColor(circle.dataset.color));
    });

    // Color picker
    if (colorPicker) {
        colorPicker.addEventListener('input', () => {
            setColorFromHex(colorPicker.value);
        });
    }

    // Hex input — update live as user types
    if (colorHexInput) {
        colorHexInput.addEventListener('input', () => {
            let hex = colorHexInput.value.trim();
            if (!hex.startsWith('#')) hex = '#' + hex;
            if (/^#[0-9a-f]{6}$/i.test(hex)) {
                setColorFromHex(hex);
            }
        });
    }


    // Fill toggle
    if (fillToggle) {
        fillToggle.addEventListener('change', () => toggleFill());
    }

    // Double-click signature button to draw a new signature
    const sigBtn = document.getElementById('pdf-btn-signature');
    if (sigBtn) {
        sigBtn.addEventListener('dblclick', () => {
            signatureDataUrl = null;
            signatureImg = null;
            openSignatureModal();
        });
    }

    // Drop zone click to upload
    if (emptyState) {
        emptyState.addEventListener('click', () => fileInput.click());
    }

    // Expand sidebar
    const expandBtn = document.getElementById('pdf-expand-sidebar');
    if (expandBtn) {
        expandBtn.addEventListener('click', () => {
            const sidebar = document.querySelector('.pdf-right-sidebar');
            if (sidebar) sidebar.classList.remove('collapsed');
            expandBtn.classList.add('hidden');
        });
    }

    // Drag-and-drop PDF upload on the scroll area
    const scrollArea = document.getElementById('pdf-pages-scroll');
    if (scrollArea) {
        scrollArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            const dz = document.querySelector('.pdf-drop-zone');
            if (dz) dz.classList.add('drag-active');
        });
        scrollArea.addEventListener('dragleave', (e) => {
            if (!scrollArea.contains(e.relatedTarget)) {
                const dz = document.querySelector('.pdf-drop-zone');
                if (dz) dz.classList.remove('drag-active');
            }
        });
        scrollArea.addEventListener('drop', (e) => {
            e.preventDefault();
            const dz = document.querySelector('.pdf-drop-zone');
            if (dz) dz.classList.remove('drag-active');
            const file = e.dataTransfer.files[0];
            if (file && file.type === 'application/pdf') {
                const dt = new DataTransfer();
                dt.items.add(file);
                fileInput.files = dt.files;
                fileInput.dispatchEvent(new Event('change'));
            }
        });

        // Pinch-to-zoom (Ctrl+wheel)
        scrollArea.addEventListener('wheel', (e) => {
            if (!e.ctrlKey) return;
            e.preventDefault();
            const delta = e.deltaY > 0 ? -ZOOM_STEP : ZOOM_STEP;
            setZoom(scale + delta);
        }, { passive: false });
    }

    if (textInput) {
        textInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') commitText();
        });
        textInput.addEventListener('blur', () => {
            commitTimeout = setTimeout(commitText, 150);
        });
    }

    window.addEventListener('keydown', (e) => {
        if (document.activeElement === textInput) return;
        const findInput = document.getElementById('find-input');
        if (document.activeElement === findInput && !(e.ctrlKey || e.metaKey)) return;

        // Ctrl+F / Cmd+F to open find bar
        if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
            e.preventDefault();
            openFindBar();
            return;
        }

        // Ctrl+S / Cmd+S to download
        if ((e.ctrlKey || e.metaKey) && e.key === 's') {
            e.preventDefault();
            downloadPDF();
            return;
        }

        // Undo/Redo
        if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
            e.preventDefault(); undo(); return;
        }
        if ((e.ctrlKey || e.metaKey) && (e.key === 'z' || e.key === 'Z') && e.shiftKey) {
            e.preventDefault(); redo(); return;
        }

        // Tool shortcuts, zoom, delete, and shortcut panel (only when not in an input)
        if (!document.activeElement.matches('input, textarea, [contenteditable]')) {
            if ((e.key === 'Delete' || e.key === 'Backspace') && selectedIndex !== -1) {
                deleteSelected();
                return;
            }
            const toolShortcuts = {
                'v': 'select', 'r': 'rect', 'o': 'ellipse', 'l': 'line',
                'a': 'arrow', 'p': 'pen', 'h': 'highlight', 'e': 'eraser', 'd': 'redact',
                't': 'text', 'n': 'note', 's': 'signature',
            };
            const tool = toolShortcuts[e.key.toLowerCase()];
            if (tool) { e.preventDefault(); setTool(tool); return; }

            // Zoom shortcuts
            if (e.key === '+' || e.key === '=') { e.preventDefault(); setZoom(scale + ZOOM_STEP); return; }
            if (e.key === '-') { e.preventDefault(); setZoom(scale - ZOOM_STEP); return; }
            if ((e.ctrlKey || e.metaKey) && e.key === '0') { e.preventDefault(); setZoom(DEFAULT_SCALE); return; }

            if (e.key === '?') {
                e.preventDefault();
                if (toggleShortcutPanel) toggleShortcutPanel();
                return;
            }
        }
    });

    window.addEventListener('keydown', (e) => { if (e.key === 'Shift') isShiftHeld = true; });
    window.addEventListener('keyup', (e) => { if (e.key === 'Shift') isShiftHeld = false; });
    window.addEventListener('blur', () => { isShiftHeld = false; });
    window.addEventListener('mouseup', () => { isDrawing = false; isDragging = false; isResizing = false; resizeOrigin = null; });

    // Context bar event listeners (refs cached in cacheToolbar)
    if (thicknessSlider) {
        thicknessSlider.addEventListener('pointerdown', () => {
            if (selectedIndex !== -1) pushUndoSnapshot();
        });
        thicknessSlider.addEventListener('input', () => {
            currentThickness = parseInt(thicknessSlider.value);
            thicknessLabel.textContent = currentThickness + 'px';
            // Update stroke preview SVG
            const previewLine = document.querySelector('#stroke-preview line');
            if (previewLine) previewLine.setAttribute('stroke-width', Math.min(currentThickness, 16));
            if (selectedIndex !== -1 && annotations[selectedIndex].type !== 'text') {
                // Store thickness in PDF-space
                annotations[selectedIndex].thickness = canvasToPdf(currentThickness);
                redrawAnnotations(annotations[selectedIndex].pageIndex);
            }
        });
    }

    if (opacitySlider) {
        opacitySlider.addEventListener('pointerdown', () => {
            if (selectedIndex !== -1) pushUndoSnapshot();
        });
        opacitySlider.addEventListener('input', () => {
            currentOpacity = parseInt(opacitySlider.value) / 100;
            opacityLabel.textContent = opacitySlider.value + '%';
            if (selectedIndex !== -1) {
                annotations[selectedIndex].opacity = currentOpacity;
                redrawAnnotations(annotations[selectedIndex].pageIndex);
            }
        });
    }

    if (fontSizeInput) {
        fontSizeInput.addEventListener('change', () => {
            const val = parseInt(fontSizeInput.value);
            if (!val || val < 1) return;
            currentFontSize = val;
            if (selectedIndex !== -1 && annotations[selectedIndex].type === 'text') {
                pushUndoSnapshot();
                const ann = annotations[selectedIndex];
                ann.size = currentFontSize;
                const ctx = pages[ann.pageIndex].ctx;
                // Measure at canvas-pixel size, convert width to PDF-space
                ctx.font = `${pdfToCanvas(ann.size)}px ${FONT_CANVAS_MAP[ann.fontFamily || 'Helvetica']}`;
                ann.w = canvasToPdf(ctx.measureText(ann.text).width);
                ann.h = ann.size;
                redrawAnnotations(ann.pageIndex);
            }
        });
    }

    if (fontFamilySelect) {
        fontFamilySelect.addEventListener('change', () => {
            currentFontFamily = fontFamilySelect.value;
            if (textInput && !textInput.classList.contains('hidden')) {
                textInput.style.fontFamily = FONT_CANVAS_MAP[currentFontFamily];
            }
            if (selectedIndex !== -1 && annotations[selectedIndex].type === 'text') {
                pushUndoSnapshot();
                const ann = annotations[selectedIndex];
                ann.fontFamily = currentFontFamily;
                const ctx = pages[ann.pageIndex].ctx;
                // Measure at canvas-pixel size, convert width to PDF-space
                ctx.font = `${pdfToCanvas(ann.size)}px ${FONT_CANVAS_MAP[ann.fontFamily]}`;
                ann.w = canvasToPdf(ctx.measureText(ann.text).width);
                redrawAnnotations(ann.pageIndex);
            }
        });
    }

    const splitRangeInput = document.getElementById('split-range-input');
    if (splitRangeInput) {
        splitRangeInput.addEventListener('input', (e) => {
            const { ranges, error } = parseSplitRanges(e.target.value, pages.length);
            const preview = document.getElementById('split-preview');
            const errorEl = document.getElementById('split-error');
            if (error) {
                if (errorEl) errorEl.textContent = error;
                if (preview) preview.textContent = '';
            } else {
                if (errorEl) errorEl.textContent = '';
                if (preview) preview.textContent = ranges.map((r, i) =>
                    'File ' + (i + 1) + ': Page' + (r.start === r.end ? '' : 's') + ' ' +
                    r.start + (r.start !== r.end ? '-' + r.end : '')
                ).join(' | ');
            }
        });
    }
}

// ==========================================
// PDF Loading
// ==========================================
async function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    if (file.type !== 'application/pdf') {
        showStatus('Please upload a valid PDF file.', 'error');
        return;
    }

    try {
        showLoading('Loading PDF...');
        imageCache.clear();
        tintedSigCache.clear();
        cachedTextContent = null;
        const arrayBuffer = await file.arrayBuffer();
        existingPdfBytes = arrayBuffer;

        const loadingTask = pdfjsLib.getDocument(existingPdfBytes.slice(0));
        const pdf = await loadingTask.promise;

        if (pdf.numPages === 0) throw new Error("Empty PDF");

        if (drawRAF) { cancelAnimationFrame(drawRAF); drawRAF = null; }
        if (commitTimeout) { clearTimeout(commitTimeout); commitTimeout = null; }
        pageCleanups.forEach(fn => fn());
        pageCleanups = [];
        pagesContainer.innerHTML = '';
        pages = [];
        annotations = [];
        emptyState.classList.add('hidden');

        // On mobile, auto-fit first page fully within viewport
        if (window.innerWidth <= 600) {
            const firstPage = await pdf.getPage(1);
            const unscaledVp = firstPage.getViewport({ scale: 1 });
            const scrollEl = document.getElementById('pdf-pages-scroll');
            // Read padding from CSS so this stays in sync with the stylesheet
            const cs = getComputedStyle(scrollEl);
            const padX = parseFloat(cs.paddingLeft) + parseFloat(cs.paddingRight);
            const padY = parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom);
            const availW = scrollEl.clientWidth - padX;
            const availH = scrollEl.clientHeight - padY;
            if (availW > 0 && availH > 0) {
                const fitScale = Math.min(availW / unscaledVp.width, availH / unscaledVp.height);
                scale = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, fitScale));
            }
        }

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const viewport = page.getViewport({ scale });

            // Page wrapper
            const wrapper = document.createElement('div');
            wrapper.className = 'page-wrapper';
            wrapper.style.width = `${viewport.width}px`;
            wrapper.style.height = `${viewport.height}px`;

            // PDF render canvas
            const pdfCanvas = document.createElement('canvas');
            pdfCanvas.width = viewport.width;
            pdfCanvas.height = viewport.height;
            await page.render({
                canvasContext: pdfCanvas.getContext('2d'),
                viewport,
            }).promise;

            // Annotation draw layer
            const drawCanvas = document.createElement('canvas');
            drawCanvas.className = 'draw-layer';
            drawCanvas.width = viewport.width;
            drawCanvas.height = viewport.height;

            wrapper.appendChild(pdfCanvas);
            wrapper.appendChild(drawCanvas);
            pagesContainer.appendChild(wrapper);

            pages.push({
                wrapper,
                canvas: drawCanvas,
                ctx: drawCanvas.getContext('2d'),
            });

            pageCleanups.push(setupCanvasEvents(drawCanvas, i - 1));
        }

        pageRotations = new Array(pdf.numPages).fill(0);
        pageOrder = Array.from({ length: pdf.numPages }, (_, i) => i);
        undoStack = [];
        redoStack = [];
        updateUndoRedoButtons();

        selectedIndex = -1;
        editingIndex = -1;
        hideTextInput();
        updateDeleteButton();
        updateZoomUI();
        hideLoading();
        showStatus('PDF loaded', 'success');
        renderThumbnails();
        if (pageIndicator) {
            pageIndicator.updateTotal();
            pageIndicator.setupObserver();
        }
        clearFormOverlays();
        await detectFormFields();
        fileInput.value = '';
    } catch (err) {
        hideLoading();
        console.error(err);
        showStatus('Error loading PDF: ' + err.message, 'error');
    }
}

// ==========================================
// Canvas Event Binding
// ==========================================
function setupCanvasEvents(canvas, pageIndex) {
    const controller = new AbortController();
    const signal = controller.signal;

    canvas.addEventListener('mousedown', (e) => handleMouseDown(e, pageIndex), { signal });
    canvas.addEventListener('mousemove', (e) => handleMouseMove(e, pageIndex), { signal });
    canvas.addEventListener('mouseup', (e) => handleMouseUp(e, pageIndex), { signal });

    // Forward single-touch to mouse events
    const forwardTouch = (type, e) => {
        if (e.touches.length > 1) return;
        e.preventDefault();
        const touch = e.touches[0] || e.changedTouches[0];
        canvas.dispatchEvent(new MouseEvent(type, {
            clientX: touch.clientX,
            clientY: touch.clientY,
            bubbles: true,
        }));
    };

    canvas.addEventListener('touchstart', (e) => forwardTouch('mousedown', e), { signal });
    canvas.addEventListener('touchmove', (e) => forwardTouch('mousemove', e), { signal });
    canvas.addEventListener('touchend', (e) => forwardTouch('mouseup', e), { signal });

    return () => controller.abort();
}

// ==========================================
// Utility Functions
// ==========================================
function updateLineBounds(ann) {
    ann.x = Math.min(ann.x1, ann.x2);
    ann.y = Math.min(ann.y1, ann.y2);
    ann.w = Math.abs(ann.x2 - ann.x1) || 6;
    ann.h = Math.abs(ann.y2 - ann.y1) || 6;
}

function pointToSegmentDist(px, py, x1, y1, x2, y2) {
    const dx = x2 - x1, dy = y2 - y1;
    const lenSq = dx * dx + dy * dy;
    if (lenSq === 0) return Math.hypot(px - x1, py - y1);
    let t = ((px - x1) * dx + (py - y1) * dy) / lenSq;
    t = Math.max(0, Math.min(1, t));
    return Math.hypot(px - (x1 + t * dx), py - (y1 + t * dy));
}

function hitTestAnnotation(ann, mx, my) {
    // All coordinates (ann and mx/my) are in PDF-space
    if (ann.type === 'line' || ann.type === 'arrow') {
        const threshold = Math.max((ann.thickness ?? 2) + canvasToPdf(4), canvasToPdf(8));
        return pointToSegmentDist(mx, my, ann.x1, ann.y1, ann.x2, ann.y2) < threshold;
    }
    return mx >= ann.x && mx <= ann.x + ann.w && my >= ann.y && my <= ann.y + ann.h;
}

function drawArrowhead(ctx, x1, y1, x2, y2, size, color) {
    const angle = Math.atan2(y2 - y1, x2 - x1);
    const headLen = Math.max(size * 3, 10);
    ctx.save();
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(x2 - headLen * Math.cos(angle - Math.PI / 6), y2 - headLen * Math.sin(angle - Math.PI / 6));
    ctx.lineTo(x2 - headLen * Math.cos(angle + Math.PI / 6), y2 - headLen * Math.sin(angle + Math.PI / 6));
    ctx.closePath();
    ctx.fill();
    ctx.restore();
}

function snapAngle(x1, y1, x2, y2) {
    const dx = x2 - x1, dy = y2 - y1;
    const dist = Math.hypot(dx, dy);
    const angle = Math.round(Math.atan2(dy, dx) / (Math.PI / 12)) * (Math.PI / 12);
    return { x: x1 + dist * Math.cos(angle), y: y1 + dist * Math.sin(angle) };
}

// ==========================================
// Undo / Redo
// ==========================================
function pushUndoSnapshot() {
    undoStack.push({
        annotations: JSON.parse(JSON.stringify(annotations)),
        pageRotations: [...pageRotations],
        pageOrder: [...pageOrder],
        pdfBytes: existingPdfBytes, // Store reference to PDF bytes for crop undo
    });
    if (undoStack.length > MAX_UNDO) undoStack.shift();
    redoStack = [];
    updateUndoRedoButtons();
}

async function restoreSnapshot(fromStack, toStack) {
    if (fromStack.length === 0) return;
    toStack.push({
        annotations: JSON.parse(JSON.stringify(annotations)),
        pageRotations: [...pageRotations],
        pageOrder: [...pageOrder],
        pdfBytes: existingPdfBytes,
    });
    const snapshot = fromStack.pop();
    const needsRerender = snapshot.pdfBytes !== existingPdfBytes ||
        snapshot.pageOrder.length !== pageOrder.length ||
        snapshot.pageOrder.some((v, i) => v !== pageOrder[i]) ||
        snapshot.pageRotations.some((v, i) => v !== pageRotations[i]);
    annotations = snapshot.annotations;
    for (const ann of annotations) {
        if (ann.color) ann.color = rgb(ann.color.red, ann.color.green, ann.color.blue);
        if (ann.type === 'text' && !ann.fontFamily) ann.fontFamily = 'Helvetica';
    }
    pageRotations = snapshot.pageRotations;
    pageOrder = snapshot.pageOrder;
    if (snapshot.pdfBytes) existingPdfBytes = snapshot.pdfBytes;
    selectedIndex = -1;
    editingIndex = -1;
    if (needsRerender && existingPdfBytes) await reloadPagesFromBytes();
    updateDeleteButton();
    updateUndoRedoButtons();
    redrawAnnotations();
    clearFormOverlays();
    cachedTextContent = null;
    await detectFormFields();
    renderThumbnails();
    if (pageIndicator) { pageIndicator.updateTotal(); pageIndicator.setupObserver(); }
}

async function undo() {
    if (undoRedoInProgress) return;
    undoRedoInProgress = true;
    try { await restoreSnapshot(undoStack, redoStack); }
    finally { undoRedoInProgress = false; }
}

async function redo() {
    if (undoRedoInProgress) return;
    undoRedoInProgress = true;
    try { await restoreSnapshot(redoStack, undoStack); }
    finally { undoRedoInProgress = false; }
}

function updateUndoRedoButtons() {
    if (undoBtnRef) undoBtnRef.disabled = undoStack.length === 0;
    if (redoBtnRef) redoBtnRef.disabled = redoStack.length === 0;
}

async function reloadPagesFromBytes() {
    if (drawRAF) { cancelAnimationFrame(drawRAF); drawRAF = null; }
    if (commitTimeout) { clearTimeout(commitTimeout); commitTimeout = null; }
    pageCleanups.forEach(fn => fn());
    pageCleanups = [];
    const pdf = await pdfjsLib.getDocument(existingPdfBytes.slice(0)).promise;
    pagesContainer.innerHTML = '';
    pages = [];
    for (let i = 0; i < pageOrder.length; i++) {
        const origIdx = pageOrder[i];
        const pdfPage = await pdf.getPage(origIdx + 1);
        const viewport = pdfPage.getViewport({ scale, rotation: pageRotations[i] });

        const wrapper = document.createElement('div');
        wrapper.className = 'page-wrapper';
        wrapper.style.width = `${viewport.width}px`;
        wrapper.style.height = `${viewport.height}px`;

        const pdfCanvas = document.createElement('canvas');
        pdfCanvas.width = viewport.width;
        pdfCanvas.height = viewport.height;
        await pdfPage.render({ canvasContext: pdfCanvas.getContext('2d'), viewport }).promise;

        const drawCanvas = document.createElement('canvas');
        drawCanvas.className = 'draw-layer';
        drawCanvas.width = viewport.width;
        drawCanvas.height = viewport.height;

        wrapper.appendChild(pdfCanvas);
        wrapper.appendChild(drawCanvas);
        pagesContainer.appendChild(wrapper);

        const pageObj = { wrapper, canvas: drawCanvas, ctx: drawCanvas.getContext('2d') };
        pages.push(pageObj);
        pageCleanups.push(setupCanvasEvents(drawCanvas, i));
    }
}

function updateZoomUI() {
    const label = document.getElementById('pdf-zoom-label');
    if (label) label.textContent = Math.round(scale * 100) + '%';
    const zoomIn = document.getElementById('pdf-btn-zoom-in');
    const zoomOut = document.getElementById('pdf-btn-zoom-out');
    if (zoomIn) zoomIn.disabled = scale >= ZOOM_MAX;
    if (zoomOut) zoomOut.disabled = scale <= ZOOM_MIN;
}

// ==========================================
// Zoom
// ==========================================
async function setZoom(newScale) {
    if (zoomInProgress) return;
    zoomInProgress = true;
    hideNotePopup();
    try {
        newScale = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, newScale));
        if (newScale === scale) return;
        if (!existingPdfBytes) return;
        if (textInput && !textInput.classList.contains('hidden')) commitText();
        isDrawing = false; isDragging = false; isResizing = false;

        const savedFormValues = saveFormValues();
        const scrollEl = document.getElementById('pdf-pages-scroll');
        const centerX = (scrollEl.scrollLeft + scrollEl.clientWidth / 2) / scale;
        const centerY = (scrollEl.scrollTop + scrollEl.clientHeight / 2) / scale;

        scale = newScale;
        await reloadPagesFromBytes();

        scrollEl.scrollLeft = centerX * scale - scrollEl.clientWidth / 2;
        scrollEl.scrollTop = centerY * scale - scrollEl.clientHeight / 2;

        updateZoomUI();

        selectedIndex = -1;
        updateDeleteButton();
        redrawAnnotations();
        renderThumbnails();
        clearFormOverlays();
        await detectFormFields();
        restoreFormValues(savedFormValues);
        if (pageIndicator) {
            pageIndicator.setupObserver();
        }
    } finally {
        zoomInProgress = false;
    }
}

// ==========================================
// Mouse Handlers
// ==========================================
function getCanvasCoords(e, pageIndex) {
    const rect = pages[pageIndex].canvas.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
}

function handleMouseDown(e, pageIndex) {
    activePageIndex = pageIndex;
    const { x: mouseX, y: mouseY } = getCanvasCoords(e, pageIndex);

    if (textInput && !textInput.classList.contains('hidden')) commitText();
    if (activeNotePopup) hideNotePopup();

    // Signature tool: place signature at click position
    if (currentTool === 'signature') {
        if (!signatureImg) {
            openSignatureModal();
            return;
        }
        const pdfX = canvasToPdf(mouseX);
        const pdfY = canvasToPdf(mouseY);
        const sigW = canvasToPdf(signatureImg.width * 0.5);
        const sigH = canvasToPdf(signatureImg.height * 0.5);

        pushUndoSnapshot();
        annotations.push({
            type: 'signature',
            pageIndex: pageIndex,
            x: pdfX,
            y: pdfY,
            w: sigW,
            h: sigH,
            dataUrl: signatureDataUrl,
            color: currentColor,
            opacity: currentOpacity,
        });
        selectedIndex = annotations.length - 1;
        redrawAnnotations(pageIndex);
        setTool('select');
        return;
    }

    // Stamp tool: place stamp at click position
    if (currentTool === 'stamp' && pendingStampText) {
        const pdfX = canvasToPdf(mouseX);
        const pdfY = canvasToPdf(mouseY);
        const ctx = pages[pageIndex].ctx;
        const stampFontSize = 24;
        ctx.font = `bold ${pdfToCanvas(stampFontSize)}px Helvetica, Arial, sans-serif`;
        const tw = canvasToPdf(ctx.measureText(pendingStampText).width);
        const pad = stampFontSize * 0.4;
        const stampW = tw + pad * 2;
        const stampH = stampFontSize + pad * 2;

        pushUndoSnapshot();
        annotations.push({
            type: 'stamp',
            pageIndex: pageIndex,
            x: pdfX,
            y: pdfY,
            w: stampW,
            h: stampH,
            text: pendingStampText,
            fontFamily: 'Helvetica',
            color: currentColor,
            opacity: currentOpacity,
        });
        selectedIndex = annotations.length - 1;
        redrawAnnotations(pageIndex);
        setTool('select');
        return;
    }

    // Note tool: place a sticky note at click position
    if (currentTool === 'note') {
        const pdfX = canvasToPdf(mouseX);
        const pdfY = canvasToPdf(mouseY);
        pushUndoSnapshot();
        annotations.push({
            type: 'note',
            pageIndex: pageIndex,
            x: pdfX,
            y: pdfY,
            w: 20 / scale,
            h: 20 / scale,
            text: '',
            color: currentColor,
            opacity: 1,
        });
        const newIndex = annotations.length - 1;
        selectedIndex = newIndex;
        redrawAnnotations(pageIndex);
        setTool('select');
        showNotePopup(newIndex);
        return;
    }

    // Text tool: open input at click position
    if (currentTool === 'text') {
        showTextInput(mouseX, mouseY);
        return;
    }

    // Select tool: check resize handle, then hit-test
    if (currentTool === 'select') {
        // Convert mouse coords to PDF-space for annotation comparisons
        const pdfMX = canvasToPdf(mouseX);
        const pdfMY = canvasToPdf(mouseY);

        if (selectedIndex !== -1 && annotations[selectedIndex].pageIndex === pageIndex) {
            const ann = annotations[selectedIndex];
            // Resize handle is rendered at canvas-space position, so hit-test in canvas-space
            const handleX = pdfToCanvas(ann.x + ann.w);
            const handleY = pdfToCanvas(ann.y + ann.h);

            if (mouseX >= handleX - 5 && mouseX <= handleX + RESIZE_HANDLE_SIZE + 5 &&
                mouseY >= handleY - 5 && mouseY <= handleY + RESIZE_HANDLE_SIZE + 5) {
                pushUndoSnapshot();
                isResizing = true;
                startX = mouseX;
                startY = mouseY;
                resizeOrigin = {
                    w: ann.w, h: ann.h,
                    points: ann.type === 'scribble' ? ann.points.map(p => ({ ...p })) : null,
                };
                return;
            }
        }

        // Hit-test annotations in PDF-space (top-most first)
        const hitIndex = annotations.slice().reverse().findIndex(ann =>
            ann.pageIndex === pageIndex && hitTestAnnotation(ann, pdfMX, pdfMY)
        );

        if (hitIndex !== -1) {
            const newSelectedIndex = annotations.length - 1 - hitIndex;
            const ann = annotations[newSelectedIndex];

            // Click on already-selected text → open for editing (pass canvas-space coords for CSS positioning)
            if (ann.type === 'text' && newSelectedIndex === selectedIndex) {
                showTextInput(pdfToCanvas(ann.x), pdfToCanvas(ann.y), ann.text, selectedIndex);
                return;
            }

            // Click on already-selected note → open popup
            if (ann.type === 'note' && newSelectedIndex === selectedIndex) {
                showNotePopup(newSelectedIndex);
                return;
            }

            selectedIndex = newSelectedIndex;

            if (ann.type === 'rect' || ann.type === 'ellipse') {
                isFillMode = ann.filled;
                if (fillToggle) fillToggle.checked = isFillMode;
            }

            pushUndoSnapshot();
            isDragging = true;
            // Store drag offset in PDF-space
            dragOffsetX = pdfMX - ann.x;
            dragOffsetY = pdfMY - ann.y;

            // Sync context bar to selected annotation
            // Thickness is stored in PDF-space; convert to display pixels for slider
            if (thicknessSlider && ann.type !== 'text') {
                const displayT = Math.round(pdfToCanvas(ann.thickness ?? (2 / scale)));
                thicknessSlider.value = displayT;
                thicknessLabel.textContent = displayT + 'px';
                currentThickness = displayT;
            }
            if (opacitySlider) {
                const o = Math.round((ann.opacity ?? 1) * 100);
                opacitySlider.value = o;
                opacityLabel.textContent = o + '%';
                currentOpacity = ann.opacity ?? 1;
            }
            if (fontSizeInput && ann.type === 'text') {
                fontSizeInput.value = ann.size;
                currentFontSize = ann.size;
            }
            if (fontFamilySelect && ann.type === 'text') {
                currentFontFamily = ann.fontFamily || 'Helvetica';
                fontFamilySelect.value = currentFontFamily;
            }
            // Sync color inputs to selected annotation
            currentColor = ann.color;
            syncColorInputs(ann.color);
            updateContextBar();
        } else {
            selectedIndex = -1;
        }

        updateDeleteButton();
        redrawAnnotations(pageIndex);
        return;
    }

    // Crop tool: draw crop rectangle
    if (currentTool === 'crop') {
        isDrawing = true;
        startX = mouseX;
        startY = mouseY;
        cropRect = null;
        document.getElementById('crop-toolbar')?.classList.add('hidden');
        return;
    }

    // Eraser tool: delete annotation on click
    if (currentTool === 'eraser') {
        eraserSnapshotPushed = false;
        const eraserPdfX = canvasToPdf(mouseX);
        const eraserPdfY = canvasToPdf(mouseY);
        const hitIdx = annotations.slice().reverse().findIndex(ann =>
            ann.pageIndex === pageIndex && hitTestAnnotation(ann, eraserPdfX, eraserPdfY)
        );
        if (hitIdx !== -1) {
            pushUndoSnapshot();
            eraserSnapshotPushed = true;
            annotations.splice(annotations.length - 1 - hitIdx, 1);
            selectedIndex = -1;
            updateDeleteButton();
            redrawAnnotations(pageIndex);
        }
        isDrawing = true;
        return;
    }

    // Start drawing
    isDrawing = true;
    startX = mouseX;
    startY = mouseY;

    if (currentTool === 'pen' || currentTool === 'highlight') {
        currentPath = [{ x: mouseX, y: mouseY }];
    }
}

function handleMouseMove(e, pageIndex) {
    if (pageIndex !== activePageIndex) return;

    const { x: currentX, y: currentY } = getCanvasCoords(e, pageIndex);
    const ctx = pages[pageIndex].ctx;
    const canvas = pages[pageIndex].canvas;

    // Resize selected annotation (work in PDF-space)
    if (isResizing && selectedIndex !== -1) {
        const ann = annotations[selectedIndex];
        const pdfCurX = canvasToPdf(currentX);
        const pdfCurY = canvasToPdf(currentY);
        const newW = Math.max(canvasToPdf(5), pdfCurX - ann.x);
        const newH = Math.max(canvasToPdf(5), pdfCurY - ann.y);

        if (ann.type === 'scribble' && resizeOrigin && resizeOrigin.points) {
            const scaleX = newW / resizeOrigin.w;
            const scaleY = newH / resizeOrigin.h;
            ann.points = resizeOrigin.points.map(p => ({
                x: ann.x + (p.x - ann.x) * scaleX,
                y: ann.y + (p.y - ann.y) * scaleY,
            }));
            // Recalculate bounds from actual scaled points
            let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
            for (const p of ann.points) {
                if (p.x < minX) minX = p.x;
                if (p.x > maxX) maxX = p.x;
                if (p.y < minY) minY = p.y;
                if (p.y > maxY) maxY = p.y;
            }
            ann.x = minX; ann.y = minY;
            ann.w = maxX - minX; ann.h = maxY - minY;
        } else if (ann.type === 'text' && newH > canvasToPdf(8)) {
            ann.size = newH;
            // Measure text at canvas-pixel size, convert width back to PDF-space
            ctx.font = `${pdfToCanvas(ann.size)}px ${FONT_CANVAS_MAP[ann.fontFamily || 'Helvetica']}`;
            ann.w = canvasToPdf(ctx.measureText(ann.text).width);
            ann.h = ann.size;
        } else if (ann.type === 'line' || ann.type === 'arrow') {
            ann.x2 = pdfCurX;
            ann.y2 = pdfCurY;
            updateLineBounds(ann);
        }

        if (ann.type !== 'line' && ann.type !== 'arrow' && ann.type !== 'scribble') {
            ann.w = newW;
            if (ann.type !== 'text') ann.h = newH;
        }

        scheduleRedraw(pageIndex);
        return;
    }

    // Drag selected annotation (work in PDF-space)
    if (currentTool === 'select' && isDragging && selectedIndex !== -1) {
        const ann = annotations[selectedIndex];
        const pdfCurX = canvasToPdf(currentX);
        const pdfCurY = canvasToPdf(currentY);
        const dx = (pdfCurX - dragOffsetX) - ann.x;
        const dy = (pdfCurY - dragOffsetY) - ann.y;

        ann.x += dx;
        ann.y += dy;

        if (ann.type === 'scribble') {
            for (const p of ann.points) {
                p.x += dx;
                p.y += dy;
            }
        }

        if (ann.type === 'line' || ann.type === 'arrow') {
            ann.x1 += dx; ann.y1 += dy;
            ann.x2 += dx; ann.y2 += dy;
            updateLineBounds(ann);
        }

        scheduleRedraw(pageIndex);
        return;
    }

    if (!isDrawing) return;

    // Eraser: drag to delete multiple
    if (currentTool === 'eraser' && isDrawing) {
        const eraserPdfCX = canvasToPdf(currentX);
        const eraserPdfCY = canvasToPdf(currentY);
        const hitIdx = annotations.slice().reverse().findIndex(ann =>
            ann.pageIndex === pageIndex && hitTestAnnotation(ann, eraserPdfCX, eraserPdfCY)
        );
        if (hitIdx !== -1) {
            if (!eraserSnapshotPushed) {
                pushUndoSnapshot();
                eraserSnapshotPushed = true;
            }
            annotations.splice(annotations.length - 1 - hitIdx, 1);
            selectedIndex = -1;
            redrawAnnotations(pageIndex);
        }
        return;
    }

    // Live preview while drawing
    if (currentTool === 'rect') {
        scheduleRedraw(pageIndex, () => {
            ctx.beginPath();
            ctx.rect(startX, startY, currentX - startX, currentY - startY);
            ctx.globalAlpha = currentOpacity;
            if (isFillMode) {
                ctx.fillStyle = getCssColor(currentColor);
                ctx.fill();
            } else {
                ctx.strokeStyle = getCssColor(currentColor);
                ctx.lineWidth = currentThickness;
                ctx.stroke();
            }
            ctx.globalAlpha = 1;
        });
    } else if (currentTool === 'ellipse') {
        scheduleRedraw(pageIndex, () => {
            const cx = (startX + currentX) / 2;
            const cy = (startY + currentY) / 2;
            const rx = Math.abs(currentX - startX) / 2;
            const ry = Math.abs(currentY - startY) / 2;
            if (rx < 1 || ry < 1) return;
            ctx.beginPath();
            ctx.ellipse(cx, cy, rx, ry, 0, 0, 2 * Math.PI);
            ctx.globalAlpha = currentOpacity;
            if (isFillMode) { ctx.fillStyle = getCssColor(currentColor); ctx.fill(); }
            else { ctx.strokeStyle = getCssColor(currentColor); ctx.lineWidth = currentThickness; ctx.stroke(); }
            ctx.globalAlpha = 1;
        });
    } else if (currentTool === 'line' || currentTool === 'arrow') {
        scheduleRedraw(pageIndex, () => {
            let ex = currentX, ey = currentY;
            if (isShiftHeld) {
                const snapped = snapAngle(startX, startY, currentX, currentY);
                ex = snapped.x; ey = snapped.y;
            }
            ctx.beginPath();
            ctx.moveTo(startX, startY);
            ctx.lineTo(ex, ey);
            ctx.strokeStyle = getCssColor(currentColor);
            ctx.lineWidth = currentThickness;
            ctx.globalAlpha = currentOpacity;
            ctx.stroke();
            if (currentTool === 'arrow') {
                drawArrowhead(ctx, startX, startY, ex, ey, currentThickness, getCssColor(currentColor));
            }
            ctx.globalAlpha = 1;
        });
    } else if (currentTool === 'pen' || currentTool === 'highlight') {
        currentPath.push({ x: currentX, y: currentY });
        const lastPt = currentPath[currentPath.length - 2];

        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';

        ctx.lineWidth = currentThickness;
        ctx.strokeStyle = getCssColor(currentColor);
        ctx.globalAlpha = currentOpacity;

        ctx.beginPath();
        ctx.moveTo(lastPt.x, lastPt.y);
        ctx.lineTo(currentX, currentY);
        ctx.stroke();
        ctx.globalAlpha = 1;
    } else if (currentTool === 'redact') {
        scheduleRedraw(pageIndex, () => {
            const rx = Math.min(startX, currentX);
            const ry = Math.min(startY, currentY);
            const rw = Math.abs(currentX - startX);
            const rh = Math.abs(currentY - startY);
            ctx.fillStyle = 'rgba(239, 68, 68, 0.3)';
            ctx.fillRect(rx, ry, rw, rh);
            ctx.strokeStyle = 'rgba(239, 68, 68, 0.8)';
            ctx.lineWidth = 1;
            ctx.setLineDash([4, 4]);
            ctx.strokeRect(rx, ry, rw, rh);
            ctx.setLineDash([]);
        });
    } else if (currentTool === 'crop') {
        scheduleRedraw(pageIndex, () => {
            const cx = Math.min(startX, currentX);
            const cy = Math.min(startY, currentY);
            const cw = Math.abs(currentX - startX);
            const ch = Math.abs(currentY - startY);
            // Dark overlay outside crop region
            ctx.fillStyle = 'rgba(0, 0, 0, 0.5)';
            ctx.fillRect(0, 0, canvas.width, cy);
            ctx.fillRect(0, cy + ch, canvas.width, canvas.height - cy - ch);
            ctx.fillRect(0, cy, cx, ch);
            ctx.fillRect(cx + cw, cy, canvas.width - cx - cw, ch);
            // Crop border
            ctx.strokeStyle = 'white';
            ctx.lineWidth = 2;
            ctx.setLineDash([6, 3]);
            ctx.strokeRect(cx, cy, cw, ch);
            ctx.setLineDash([]);
        });
    }
}

function handleMouseUp(e, pageIndex) {
    if (pageIndex !== activePageIndex) return;

    isDragging = false;
    isResizing = false;
    resizeOrigin = null;

    if (currentTool === 'select' || currentTool === 'eraser' || !isDrawing) {
        isDrawing = false;
        return;
    }
    isDrawing = false;

    const { x: endX, y: endY } = getCanvasCoords(e, pageIndex);

    // Crop tool: commit crop rectangle and show toolbar
    if (currentTool === 'crop') {
        const w = endX - startX;
        const h = endY - startY;
        if (Math.abs(w) > 5 && Math.abs(h) > 5) {
            cropRect = {
                pageIndex,
                x: canvasToPdf(w < 0 ? startX + w : startX),
                y: canvasToPdf(h < 0 ? startY + h : startY),
                w: canvasToPdf(Math.abs(w)),
                h: canvasToPdf(Math.abs(h)),
            };
            document.getElementById('crop-toolbar')?.classList.remove('hidden');
            redrawAnnotations(pageIndex);
        }
        return;
    }

    if (currentTool === 'rect') {
        const w = endX - startX;
        const h = endY - startY;

        if (Math.abs(w) > 2 && Math.abs(h) > 2) {
            pushUndoSnapshot();
            annotations.push({
                type: 'rect', pageIndex,
                x: canvasToPdf(w < 0 ? startX + w : startX),
                y: canvasToPdf(h < 0 ? startY + h : startY),
                w: canvasToPdf(Math.abs(w)), h: canvasToPdf(Math.abs(h)),
                color: currentColor, filled: isFillMode,
                opacity: currentOpacity, thickness: canvasToPdf(currentThickness),
            });
            selectedIndex = annotations.length - 1;
            setTool('select');
        }
    } else if (currentTool === 'ellipse') {
        const w = endX - startX, h = endY - startY;
        if (Math.abs(w) > 2 && Math.abs(h) > 2) {
            pushUndoSnapshot();
            annotations.push({
                type: 'ellipse', pageIndex,
                x: canvasToPdf(w < 0 ? startX + w : startX),
                y: canvasToPdf(h < 0 ? startY + h : startY),
                w: canvasToPdf(Math.abs(w)), h: canvasToPdf(Math.abs(h)),
                color: currentColor, filled: isFillMode,
                opacity: currentOpacity, thickness: canvasToPdf(currentThickness),
            });
            selectedIndex = annotations.length - 1;
            setTool('select');
        }
    } else if (currentTool === 'line' || currentTool === 'arrow') {
        let ex = endX, ey = endY;
        if (isShiftHeld) {
            const snapped = snapAngle(startX, startY, endX, endY);
            ex = snapped.x; ey = snapped.y;
        }
        if (Math.hypot(ex - startX, ey - startY) >= 3) {
            pushUndoSnapshot();
            const ann = {
                type: currentTool, pageIndex,
                x1: canvasToPdf(startX), y1: canvasToPdf(startY),
                x2: canvasToPdf(ex), y2: canvasToPdf(ey),
                x: 0, y: 0, w: 0, h: 0,
                color: currentColor, opacity: currentOpacity,
                thickness: canvasToPdf(currentThickness),
            };
            updateLineBounds(ann);
            annotations.push(ann);
            selectedIndex = annotations.length - 1;
            setTool('select');
        }
    } else if (currentTool === 'pen' || currentTool === 'highlight') {
        if (currentPath.length > 2) {
            // Convert path points to PDF-space
            const pdfPath = currentPath.map(p => ({ x: canvasToPdf(p.x), y: canvasToPdf(p.y) }));
            let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
            for (const p of pdfPath) {
                if (p.x < minX) minX = p.x;
                if (p.x > maxX) maxX = p.x;
                if (p.y < minY) minY = p.y;
                if (p.y > maxY) maxY = p.y;
            }

            const isHighlight = currentTool === 'highlight';
            pushUndoSnapshot();
            annotations.push({
                type: 'scribble', pageIndex,
                points: pdfPath,
                x: minX, y: minY,
                w: maxX - minX, h: maxY - minY,
                color: currentColor,
                opacity: isHighlight ? 0.5 : currentOpacity,
                thickness: canvasToPdf(isHighlight ? 12 : currentThickness),
            });
            selectedIndex = annotations.length - 1;
            setTool('select');
        }
        currentPath = [];
    } else if (currentTool === 'redact') {
        const w = endX - startX;
        const h = endY - startY;
        if (Math.abs(w) > 2 && Math.abs(h) > 2) {
            pushUndoSnapshot();
            annotations.push({
                type: 'redact', pageIndex,
                x: canvasToPdf(w < 0 ? startX + w : startX),
                y: canvasToPdf(h < 0 ? startY + h : startY),
                w: canvasToPdf(Math.abs(w)), h: canvasToPdf(Math.abs(h)),
            });
            selectedIndex = annotations.length - 1;
            setTool('select');
        }
    }

    redrawAnnotations(pageIndex);
}

// ==========================================
// Rendering (scoped to page + rAF batching)
// ==========================================
function scheduleRedraw(pageIndex, afterDraw) {
    if (drawRAF) cancelAnimationFrame(drawRAF);
    drawRAF = requestAnimationFrame(() => {
        redrawAnnotations(pageIndex);
        if (afterDraw) afterDraw();
        drawRAF = null;
    });
}

function redrawAnnotations(targetPage) {
    // Scope to a single page when possible, otherwise redraw all
    const pagesToRedraw = (targetPage !== undefined)
        ? [pages[targetPage]]
        : pages;

    for (const p of pagesToRedraw) {
        p.ctx.clearRect(0, 0, p.canvas.width, p.canvas.height);
    }

    const pageSet = (targetPage !== undefined)
        ? new Set([targetPage])
        : null;

    const primaryColor = cachedPrimaryColor;

    for (let index = 0; index < annotations.length; index++) {
        const ann = annotations[index];

        // Skip if scoped and not on this page
        if (pageSet && !pageSet.has(ann.pageIndex)) continue;
        // Skip the annotation being edited inline
        if (index === editingIndex) continue;

        const ctx = pages[ann.pageIndex].ctx;
        const cssColor = ann.color ? getCssColor(ann.color) : '#000';

        ctx.save();

        // Convert PDF-space coordinates to canvas-pixel space for rendering
        const rx = pdfToCanvas(ann.x), ry = pdfToCanvas(ann.y);
        const rw = pdfToCanvas(ann.w), rh = pdfToCanvas(ann.h);
        const rThick = pdfToCanvas(ann.thickness ?? (2 / scale));

        if (ann.type === 'rect') {
            ctx.globalAlpha = ann.opacity ?? 1;
            if (ann.filled) {
                ctx.fillStyle = cssColor;
                ctx.fillRect(rx, ry, rw, rh);
            } else {
                ctx.strokeStyle = cssColor;
                ctx.lineWidth = rThick;
                ctx.strokeRect(rx, ry, rw, rh);
            }
        } else if (ann.type === 'ellipse') {
            ctx.globalAlpha = ann.opacity ?? 1;
            ctx.beginPath();
            ctx.ellipse(rx + rw / 2, ry + rh / 2, rw / 2, rh / 2, 0, 0, 2 * Math.PI);
            if (ann.filled) {
                ctx.fillStyle = cssColor;
                ctx.fill();
            } else {
                ctx.strokeStyle = cssColor;
                ctx.lineWidth = rThick;
                ctx.stroke();
            }
        } else if (ann.type === 'line' || ann.type === 'arrow') {
            const rx1 = pdfToCanvas(ann.x1), ry1 = pdfToCanvas(ann.y1);
            const rx2 = pdfToCanvas(ann.x2), ry2 = pdfToCanvas(ann.y2);
            ctx.globalAlpha = ann.opacity ?? 1;
            ctx.beginPath();
            ctx.moveTo(rx1, ry1);
            ctx.lineTo(rx2, ry2);
            ctx.strokeStyle = cssColor;
            ctx.lineWidth = rThick;
            ctx.stroke();
            if (ann.type === 'arrow') {
                drawArrowhead(ctx, rx1, ry1, rx2, ry2, rThick, cssColor);
            }
        } else if (ann.type === 'scribble') {
            if (ann.points.length > 0) {
                ctx.beginPath();
                ctx.moveTo(pdfToCanvas(ann.points[0].x), pdfToCanvas(ann.points[0].y));
                for (let i = 1; i < ann.points.length; i++) {
                    ctx.lineTo(pdfToCanvas(ann.points[i].x), pdfToCanvas(ann.points[i].y));
                }
                ctx.lineCap = 'round';
                ctx.lineJoin = 'round';
                ctx.lineWidth = rThick;
                ctx.strokeStyle = cssColor;
                ctx.globalAlpha = ann.opacity ?? 1;
                ctx.stroke();
            }
        } else if (ann.type === 'text') {
            ctx.font = `${pdfToCanvas(ann.size)}px ${FONT_CANVAS_MAP[ann.fontFamily || 'Helvetica']}`;
            ctx.fillStyle = cssColor;
            ctx.textBaseline = 'top';
            ctx.globalAlpha = ann.opacity ?? 1;
            ctx.fillText(ann.text, rx, ry);
        } else if (ann.type === 'signature') {
            const img = getCachedImage(ann.dataUrl);
            ctx.globalAlpha = ann.opacity ?? 1;
            if (img.complete) {
                if (ann.color) {
                    // Tint signature using cached offscreen canvas
                    const cacheKey = ann.dataUrl + '|' + cssColor + '|' + Math.round(rw) + 'x' + Math.round(rh);
                    let offscreen = tintedSigCache.get(cacheKey);
                    if (!offscreen) {
                        offscreen = document.createElement('canvas');
                        offscreen.width = rw;
                        offscreen.height = rh;
                        const offCtx = offscreen.getContext('2d');
                        offCtx.drawImage(img, 0, 0, rw, rh);
                        offCtx.globalCompositeOperation = 'source-in';
                        offCtx.fillStyle = cssColor;
                        offCtx.fillRect(0, 0, rw, rh);
                        tintedSigCache.set(cacheKey, offscreen);
                    }
                    ctx.drawImage(offscreen, rx, ry);
                } else {
                    ctx.drawImage(img, rx, ry, rw, rh);
                }
            } else {
                img.onload = () => scheduleRedraw(ann.pageIndex);
            }
        } else if (ann.type === 'stamp') {
            const fontSize = pdfToCanvas(ann.h * 0.6);
            ctx.save();
            ctx.translate(rx + rw / 2, ry + rh / 2);
            ctx.rotate(-15 * Math.PI / 180);
            ctx.font = `bold ${fontSize}px Helvetica, Arial, sans-serif`;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.strokeStyle = cssColor;
            ctx.lineWidth = 3;
            ctx.globalAlpha = ann.opacity ?? 0.7;
            // Draw border
            const tw = ctx.measureText(ann.text).width;
            const pad = fontSize * 0.4;
            ctx.strokeRect(-tw/2 - pad, -fontSize/2 - pad, tw + pad*2, fontSize + pad*2);
            // Draw text
            ctx.fillStyle = cssColor;
            ctx.fillText(ann.text, 0, 0);
            ctx.restore();
        } else if (ann.type === 'note') {
            // Draw note icon
            const iconSize = pdfToCanvas(ann.w);
            ctx.fillStyle = cssColor;
            ctx.globalAlpha = ann.opacity ?? 1;
            ctx.fillRect(rx, ry, iconSize, iconSize);
            // Small fold effect
            ctx.fillStyle = 'rgba(0,0,0,0.15)';
            ctx.beginPath();
            ctx.moveTo(rx + iconSize - 6, ry);
            ctx.lineTo(rx + iconSize, ry + 6);
            ctx.lineTo(rx + iconSize, ry);
            ctx.fill();
            // White lines if has text
            if (ann.text) {
                ctx.strokeStyle = 'rgba(255,255,255,0.6)';
                ctx.lineWidth = 1;
                for (let i = 1; i <= 3; i++) {
                    const ly = ry + iconSize * (0.3 + i * 0.15);
                    ctx.beginPath();
                    ctx.moveTo(rx + 4, ly);
                    ctx.lineTo(rx + iconSize - 4, ly);
                    ctx.stroke();
                }
            }
        } else if (ann.type === 'redact') {
            // Semi-transparent red overlay while editing
            ctx.fillStyle = 'rgba(239, 68, 68, 0.3)';
            ctx.fillRect(rx, ry, rw, rh);
            ctx.strokeStyle = 'rgba(239, 68, 68, 0.8)';
            ctx.lineWidth = 1;
            ctx.setLineDash([4, 4]);
            ctx.strokeRect(rx, ry, rw, rh);
            ctx.setLineDash([]);
        }

        ctx.restore();

        // Selection outline + resize handle (in canvas-pixel space)
        if (index === selectedIndex) {
            ctx.strokeStyle = primaryColor;
            ctx.lineWidth = 1;
            ctx.strokeRect(rx - 2, ry - 2, rw + 4, rh + 4);
            // Notes are fixed-size icons — no resize handle
            if (ann.type !== 'note') {
                ctx.fillStyle = primaryColor;
                ctx.fillRect(rx + rw, ry + rh, RESIZE_HANDLE_SIZE, RESIZE_HANDLE_SIZE);
            }
        }
    }

    // Draw find highlights on top of annotations
    if (findMatches.length > 0) {
        for (let mi = 0; mi < findMatches.length; mi++) {
            const match = findMatches[mi];
            if (pageSet && !pageSet.has(match.pageIndex)) continue;

            const mCtx = pages[match.pageIndex].ctx;
            const mx = pdfToCanvas(match.x);
            const my = pdfToCanvas(match.y);
            const mw = pdfToCanvas(match.w);
            const mh = pdfToCanvas(match.h);

            mCtx.fillStyle = mi === findCurrentIndex
                ? 'rgba(249, 115, 22, 0.4)'
                : 'rgba(250, 204, 21, 0.3)';
            mCtx.fillRect(mx, my, mw, mh);

            if (mi === findCurrentIndex) {
                mCtx.strokeStyle = 'rgba(249, 115, 22, 0.8)';
                mCtx.lineWidth = 2;
                mCtx.strokeRect(mx, my, mw, mh);
            }
        }
    }

    // Draw crop overlay if active
    if (cropRect) {
        const cropPageIdx = cropRect.pageIndex;
        if (!pageSet || pageSet.has(cropPageIdx)) {
            const cCtx = pages[cropPageIdx].ctx;
            const cCanvas = pages[cropPageIdx].canvas;
            const cx = pdfToCanvas(cropRect.x);
            const cy = pdfToCanvas(cropRect.y);
            const cw = pdfToCanvas(cropRect.w);
            const ch = pdfToCanvas(cropRect.h);

            cCtx.fillStyle = 'rgba(0, 0, 0, 0.5)';
            cCtx.fillRect(0, 0, cCanvas.width, cy);
            cCtx.fillRect(0, cy + ch, cCanvas.width, cCanvas.height - cy - ch);
            cCtx.fillRect(0, cy, cx, ch);
            cCtx.fillRect(cx + cw, cy, cCanvas.width - cx - cw, ch);

            cCtx.strokeStyle = 'white';
            cCtx.lineWidth = 2;
            cCtx.setLineDash([6, 3]);
            cCtx.strokeRect(cx, cy, cw, ch);
            cCtx.setLineDash([]);
        }
    }

    // Update affected thumbnails
    if (targetPage !== undefined) {
        updateThumbnail(targetPage);
    } else {
        for (let i = 0; i < pages.length; i++) updateThumbnail(i);
    }
}

// ==========================================
// Public API
// ==========================================
function setTool(tool) {
    if (textInput && !textInput.classList.contains('hidden')) {
        commitText();
        hideTextInput();
    }

    // Cancel crop if switching away from crop tool
    if (currentTool === 'crop' && tool !== 'crop' && cropRect) {
        cropRect = null;
        document.getElementById('crop-toolbar')?.classList.add('hidden');
        redrawAnnotations();
    }

    const wasHighlight = currentTool === 'highlight';
    currentTool = tool;

    if (tool !== 'select') {
        selectedIndex = -1;
        updateDeleteButton();
        redrawAnnotations();
    }

    // Update active button (use cached NodeList)
    if (toolBtns) {
        toolBtns.forEach(b => b.classList.remove('active'));
    }
    const btn = document.getElementById(`pdf-btn-${tool}`);
    if (btn) btn.classList.add('active');

    if (tool === 'text' && textInput) {
        textInput.style.color = getCssColor(currentColor);
        textInput.style.fontFamily = FONT_CANVAS_MAP[currentFontFamily];
    }

    if (tool === 'signature') {
        const saved = localStorage.getItem('pdfEditorSignature');
        if (saved && !signatureDataUrl) {
            signatureDataUrl = saved;
            signatureImg = new Image();
            signatureImg.src = saved;
        }
        if (!signatureDataUrl) {
            openSignatureModal();
            return;
        }
    }

    if (tool === 'stamp') {
        pendingStampText = null;
        openStampModal();
    }

    // Update cursors — each tool gets a distinct SVG cursor via base64
    const svgCursor = (svg, x, y) =>
        `url("data:image/svg+xml;base64,${btoa(svg)}") ${x} ${y}, auto`;
    const cursorColor = getComputedStyle(document.documentElement).getPropertyValue('--text').trim() || '#1e293b';
    const head = `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="${cursorColor}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">`;
    const cursors = {
        select: 'default',
        rect:      svgCursor(`${head}<rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="3" x2="9" y2="9" opacity=".4"/></svg>`, 3, 3),
        ellipse:   svgCursor(`${head}<circle cx="12" cy="12" r="10"/><line x1="2" y1="2" x2="7" y2="7" opacity=".4"/></svg>`, 2, 2),
        line:      svgCursor(`${head}<line x1="4" y1="20" x2="20" y2="4"/><circle cx="4" cy="20" r="2" fill="${cursorColor}"/></svg>`, 4, 20),
        arrow:     svgCursor(`${head}<line x1="4" y1="20" x2="20" y2="4"/><polyline points="10 4 20 4 20 14"/></svg>`, 4, 20),
        pen:       svgCursor(`${head}<path d="M17 3a2.83 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/></svg>`, 2, 22),
        highlight: svgCursor(`<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="#facc15" fill-opacity=".6" stroke="${cursorColor}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="6" width="16" height="12" rx="2"/></svg>`, 4, 18),
        eraser:    svgCursor(`${head}<path d="m7 21-4.3-4.3c-1-1-1-2.5 0-3.4l9.6-9.6c1-1 2.5-1 3.4 0l5.6 5.6c1 1 1 2.5 0 3.4L13 21"/><path d="M22 21H7"/></svg>`, 4, 20),
        redact:    svgCursor(`${head}<rect x="3" y="3" width="18" height="18" rx="2" fill="rgba(239,68,68,0.3)"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="8" x2="21" y2="8"/><line x1="3" y1="16" x2="21" y2="16"/></svg>`, 3, 3),
        text:      svgCursor(`${head}<path d="M4 7V4h16v3"/><path d="M9 20h6"/><path d="M12 4v16"/></svg>`, 12, 20),
        note:      svgCursor(`${head}<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>`, 3, 21),
        stamp:     svgCursor(`${head}<path d="M5 22h14"/><path d="M12 17V6"/><rect x="6" y="2" width="12" height="4" rx="1"/></svg>`, 12, 20),
        crop:      svgCursor(`${head}<path d="M6 2v4H2"/><path d="M18 22v-4h4"/><rect x="6" y="6" width="12" height="12"/></svg>`, 6, 6),
    };
    const cursor = cursors[tool] || 'crosshair';
    for (const p of pages) {
        p.canvas.style.cursor = cursor;
    }

    // Set highlight defaults
    if (tool === 'highlight') {
        currentOpacity = 0.5;
        currentThickness = 12;
        if (opacitySlider) { opacitySlider.value = 50; }
        if (opacityLabel) { opacityLabel.textContent = '50%'; }
        if (thicknessSlider) { thicknessSlider.value = 12; }
        if (thicknessLabel) { thicknessLabel.textContent = '12px'; }
    }

    // Swap color presets for highlight tool
    if (tool === 'highlight' && !wasHighlight) {
        updateColorPresets(HIGHLIGHT_PRESETS);
        setColor('yellow');
    }
    if (wasHighlight && tool !== 'highlight') {
        updateColorPresets(DEFAULT_PRESETS);
        // Reset color without applyColor to avoid overwriting a just-selected annotation
        currentColor = COLORS.black;
        syncColorInputs(currentColor);
    }

    updateContextBar();
}

function setContextVisible(el, visible) {
    if (!el) return;
    if (visible) {
        el.classList.remove('hidden', 'ctx-exit');
    } else if (!el.classList.contains('hidden') && !el.classList.contains('ctx-exit')) {
        el.classList.add('ctx-exit');
        setTimeout(() => {
            if (el.classList.contains('ctx-exit')) el.classList.add('hidden');
            el.classList.remove('ctx-exit');
        }, 150);
    }
}

function updateContextBar() {
    const t = currentTool;
    const sel = selectedIndex !== -1 ? annotations[selectedIndex] : null;
    const showThickness = ['pen','highlight','rect','ellipse','line','arrow'].includes(t) ||
                          (t === 'select' && sel && sel.type !== 'text' && sel.type !== 'signature' && sel.type !== 'stamp' && sel.type !== 'redact' && sel.type !== 'note');
    const showOpacity = (t !== 'select' && t !== 'eraser' && t !== 'redact' && t !== 'crop') ||
                        (t === 'select' && sel && sel.type !== 'redact');
    const showFontSize = t === 'text' || (t === 'select' && sel && sel.type === 'text');
    const showFill = ['rect','ellipse'].includes(t) ||
                     (t === 'select' && sel && ['rect','ellipse'].includes(sel.type));
    const showColors = t !== 'eraser' && t !== 'redact' && t !== 'crop' &&
                       !(t === 'select' && sel && sel.type === 'redact');

    setContextVisible(ctxThickness, showThickness);
    setContextVisible(ctxOpacity, showOpacity);
    setContextVisible(ctxFontsize, showFontSize);
    setContextVisible(ctxFontfamily, showFontSize);
    setContextVisible(ctxFill, showFill);
    setContextVisible(ctxColor, showColors);
}

// --- Color management ---
function updateColorPresets(presets) {
    if (!colorCircles) return;
    const circles = Array.from(colorCircles);
    presets.forEach((preset, i) => {
        if (!circles[i]) return;
        circles[i].dataset.color = preset.key;
        circles[i].dataset.tooltip = preset.label;
        // Swap swatch class
        circles[i].className = circles[i].className.replace(/color-swatch-\S+/g, '');
        circles[i].classList.add('color-swatch-' + preset.key);
    });
}

function setColor(colorName) {
    const color = COLORS[colorName] || HIGHLIGHT_COLORS[colorName];
    if (color) applyColor(color, colorName);
}

function setColorFromHex(hex) {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    applyColor(rgb(r / 255, g / 255, b / 255), null);
}

function applyColor(color, presetName) {
    currentColor = color;

    if (textInput) textInput.style.color = getCssColor(currentColor);

    // Update stroke preview color
    const previewLine = document.querySelector('#stroke-preview line');
    if (previewLine) previewLine.setAttribute('stroke', getCssColor(currentColor));

    // Update preset circles
    if (colorCircles) {
        colorCircles.forEach(c => {
            c.classList.remove('selected');
            c.setAttribute('aria-checked', 'false');
        });
    }
    if (presetName) {
        const circle = colorCircles
            ? Array.from(colorCircles).find(c => c.dataset.color === presetName)
            : null;
        if (circle) {
            circle.classList.add('selected');
            circle.setAttribute('aria-checked', 'true');
        }
    }

    // Sync all color inputs
    syncColorInputs(currentColor);

    if (selectedIndex !== -1) {
        pushUndoSnapshot();
        annotations[selectedIndex].color = currentColor;
        redrawAnnotations(annotations[selectedIndex].pageIndex);
    }
}

function syncColorInputs(color) {
    const r = Math.round(color.red * 255);
    const g = Math.round(color.green * 255);
    const b = Math.round(color.blue * 255);
    const hex = '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('');

    if (colorPicker) colorPicker.value = hex;
    if (colorHexInput) colorHexInput.value = hex;
}

function toggleFill() {
    if (fillToggle) isFillMode = fillToggle.checked;

    if (selectedIndex !== -1 && ['rect', 'ellipse'].includes(annotations[selectedIndex].type)) {
        pushUndoSnapshot();
        annotations[selectedIndex].filled = isFillMode;
        redrawAnnotations(annotations[selectedIndex].pageIndex);
    }
}

function deleteSelected() {
    if (selectedIndex === -1) return;
    pushUndoSnapshot();
    const pageIdx = annotations[selectedIndex].pageIndex;
    annotations.splice(selectedIndex, 1);
    selectedIndex = -1;
    updateDeleteButton();
    redrawAnnotations(pageIdx);
    showStatus('Annotation deleted', 'info', { undoable: true });
}

function updateDeleteButton() {
    if (deleteBtn) deleteBtn.disabled = selectedIndex === -1;
}

// ==========================================
// Text Input
// ==========================================
function showTextInput(x, y, existingText = '', index = -1) {
    if (!textInput) return;
    if (commitTimeout) { clearTimeout(commitTimeout); commitTimeout = null; }

    // x, y are in canvas-pixel space (for CSS left/top positioning)
    pages[activePageIndex].wrapper.appendChild(textInput);
    textInput.style.left = `${x}px`;
    textInput.style.top = `${y}px`;
    textInput.value = existingText;
    textInput.classList.remove('hidden');
    textInput.style.width = 'auto';
    // Font size: annotation.size is in PDF-space, multiply by scale for CSS display
    const pdfFontSize = (index !== -1 && annotations[index].size) ? annotations[index].size : currentFontSize;
    textInput.style.fontSize = `${pdfToCanvas(pdfFontSize)}px`;
    const ff = (index !== -1 && annotations[index].fontFamily) || currentFontFamily;
    textInput.style.fontFamily = FONT_CANVAS_MAP[ff];
    editingIndex = index;
    redrawAnnotations(activePageIndex);
    // Delay focus so the mousedown event finishes first and doesn't steal focus back
    setTimeout(() => textInput.focus(), 0);
}

function hideTextInput() {
    if (textInput) {
        textInput.classList.add('hidden');
        textInput.value = '';
    }
    editingIndex = -1;
}

function commitText() {
    if (commitTimeout) { clearTimeout(commitTimeout); commitTimeout = null; }
    if (!textInput || textInput.classList.contains('hidden')) return;

    const text = textInput.value.trim();
    // textInput CSS left/top are in canvas-pixel space; convert to PDF-space
    const x = canvasToPdf(parseFloat(textInput.style.left));
    const y = canvasToPdf(parseFloat(textInput.style.top));

    // Hide input FIRST to prevent re-entrant calls from setTool → commitText
    textInput.classList.add('hidden');
    textInput.value = '';
    const prevEditingIndex = editingIndex;
    editingIndex = -1;

    if (text) {
        const ctx = pages[activePageIndex].ctx;
        const fontSize = (prevEditingIndex >= 0) ? annotations[prevEditingIndex].size : currentFontSize;
        const fontFamily = (prevEditingIndex >= 0) ? (annotations[prevEditingIndex].fontFamily || 'Helvetica') : currentFontFamily;
        // Measure text at canvas-pixel size, then convert width to PDF-space
        ctx.font = `${pdfToCanvas(fontSize)}px ${FONT_CANVAS_MAP[fontFamily]}`;
        const w = canvasToPdf(ctx.measureText(text).width);

        pushUndoSnapshot();
        if (prevEditingIndex >= 0) {
            annotations[prevEditingIndex].text = text;
            annotations[prevEditingIndex].x = x;
            annotations[prevEditingIndex].y = y;
            annotations[prevEditingIndex].w = w;
        } else {
            annotations.push({
                type: 'text',
                pageIndex: activePageIndex,
                x, y, w,
                h: fontSize,
                text,
                size: fontSize,
                fontFamily: fontFamily,
                color: currentColor,
                opacity: currentOpacity,
            });
            selectedIndex = annotations.length - 1;
        }
    } else if (prevEditingIndex >= 0) {
        annotations.splice(prevEditingIndex, 1);
        selectedIndex = -1;
    }

    redrawAnnotations(activePageIndex);
    setTool('select');
}

// ==========================================
// Helpers
// ==========================================
function getCssColor(pdfColor) {
    return `rgb(${Math.round(pdfColor.red * 255)},${Math.round(pdfColor.green * 255)},${Math.round(pdfColor.blue * 255)})`;
}

function getCachedImage(dataUrl) {
    if (imageCache.has(dataUrl)) return imageCache.get(dataUrl);
    const img = new Image();
    img.src = dataUrl;
    imageCache.set(dataUrl, img);
    return img;
}

function openSignatureModal() {
    const overlay = document.getElementById('signature-modal-overlay');
    const canvas = document.getElementById('signature-canvas');
    const ctx = canvas.getContext('2d');

    canvas.width = 500;
    canvas.height = 200;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.lineWidth = 2.5;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = '#1e293b';

    overlay.classList.remove('hidden');
    overlay.classList.add('active');
    const removeTrap = trapFocus(overlay);

    let drawing = false;
    let lastX, lastY;

    function getPos(e) {
        const rect = canvas.getBoundingClientRect();
        return {
            x: (e.clientX - rect.left) * (canvas.width / rect.width),
            y: (e.clientY - rect.top) * (canvas.height / rect.height),
        };
    }

    function onDown(e) {
        drawing = true;
        const pos = getPos(e);
        lastX = pos.x;
        lastY = pos.y;
    }

    function onMove(e) {
        if (!drawing) return;
        const pos = getPos(e);
        ctx.beginPath();
        ctx.moveTo(lastX, lastY);
        ctx.lineTo(pos.x, pos.y);
        ctx.stroke();
        lastX = pos.x;
        lastY = pos.y;
    }

    function onUp() { drawing = false; }

    canvas.addEventListener('mousedown', onDown);
    canvas.addEventListener('mousemove', onMove);
    canvas.addEventListener('mouseup', onUp);
    canvas.addEventListener('mouseleave', onUp);

    function escHandler(e) {
        if (e.key === 'Escape') {
            cleanup();
            setTool('select');
        }
    }

    function cleanup() {
        canvas.removeEventListener('mousedown', onDown);
        canvas.removeEventListener('mousemove', onMove);
        canvas.removeEventListener('mouseup', onUp);
        canvas.removeEventListener('mouseleave', onUp);
        overlay.removeEventListener('keydown', escHandler);
        overlay.classList.add('hidden');
        overlay.classList.remove('active');
        removeTrap();
    }

    document.getElementById('signature-clear').onclick = () => {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
    };

    document.getElementById('signature-cancel').onclick = () => {
        cleanup();
        setTool('select');
    };

    document.getElementById('signature-apply').onclick = () => {
        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const hasContent = imageData.data.some((val, i) => i % 4 === 3 && val > 0);
        if (!hasContent) return;

        const dataUrl = canvas.toDataURL('image/png');
        signatureDataUrl = dataUrl;
        signatureImg = new Image();
        signatureImg.src = dataUrl;

        if (document.getElementById('signature-save-check').checked) {
            localStorage.setItem('pdfEditorSignature', dataUrl);
        }

        cleanup();
        currentTool = 'select';
        setTool('signature');
    };

    overlay.addEventListener('keydown', escHandler);
}

// ==========================================
// Stamp Modal
// ==========================================
function openStampModal() {
    const overlay = document.getElementById('stamp-modal-overlay');
    if (!overlay) return;

    overlay.classList.remove('hidden');
    overlay.classList.add('active');
    const removeTrap = trapFocus(overlay.querySelector('.stamp-modal'));

    function selectStamp(text) {
        pendingStampText = text;
        cleanup();
    }

    function escHandler(e) {
        if (e.key === 'Escape') { cleanup(); setTool('select'); }
    }

    function backdropHandler(e) {
        if (e.target === overlay) { cleanup(); setTool('select'); }
    }

    function cleanup() {
        overlay.removeEventListener('keydown', escHandler);
        overlay.removeEventListener('click', backdropHandler);
        overlay.classList.add('hidden');
        overlay.classList.remove('active');
        if (removeTrap) removeTrap();
    }

    // Preset buttons
    overlay.querySelectorAll('.stamp-preset-btn').forEach(btn => {
        btn.onclick = () => selectStamp(btn.dataset.stamp);
    });

    // Custom text
    const customInput = document.getElementById('stamp-custom-input');
    const customApply = document.getElementById('stamp-custom-apply');
    if (customInput) customInput.value = '';

    if (customApply) {
        customApply.onclick = () => {
            const text = customInput ? customInput.value.trim() : '';
            if (text) selectStamp(text.toUpperCase());
        };
    }

    if (customInput) {
        customInput.onkeydown = (e) => {
            if (e.key === 'Enter') {
                const text = customInput.value.trim();
                if (text) selectStamp(text.toUpperCase());
            }
        };
    }

    // Cancel
    const cancelBtn = document.getElementById('stamp-cancel');
    if (cancelBtn) {
        cancelBtn.onclick = () => { cleanup(); setTool('select'); };
    }

    overlay.addEventListener('keydown', escHandler);
    overlay.addEventListener('click', backdropHandler);
}

// ==========================================
// Loading Overlay
// ==========================================
function showLoading(message) {
    const overlay = document.getElementById('loading-overlay');
    if (!overlay) return;
    const label = overlay.querySelector('.loading-label');
    if (label) label.textContent = message || '';
    overlay.classList.remove('hidden');
}

function hideLoading() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) overlay.classList.add('hidden');
}

// ==========================================
// Confirmation Modal
// ==========================================
function confirmDialog({ title, message, confirmText = 'Confirm', destructive = false } = {}) {
    return new Promise((resolve) => {
        const overlay = document.getElementById('confirm-modal-overlay');
        if (!overlay) { resolve(false); return; }

        const titleEl = document.getElementById('confirm-modal-title');
        const messageEl = document.getElementById('confirm-modal-message');
        const confirmBtn = document.getElementById('confirm-modal-confirm');
        const cancelBtn = document.getElementById('confirm-modal-cancel');

        if (titleEl) titleEl.textContent = title || '';
        if (messageEl) messageEl.textContent = message || '';
        if (confirmBtn) {
            confirmBtn.textContent = confirmText;
            confirmBtn.className = destructive ? 'btn btn-destructive' : 'btn';
        }

        const previousFocus = document.activeElement;
        overlay.classList.remove('hidden');
        const removeTrap = trapFocus(overlay.querySelector('.confirm-modal'));

        function close(result) {
            overlay.classList.add('hidden');
            if (removeTrap) removeTrap();
            confirmBtn.removeEventListener('click', onConfirm);
            cancelBtn.removeEventListener('click', onCancel);
            overlay.removeEventListener('click', onBackdrop);
            document.removeEventListener('keydown', onKey);
            if (previousFocus) previousFocus.focus();
            resolve(result);
        }

        function onConfirm() { close(true); }
        function onCancel() { close(false); }
        function onBackdrop(e) { if (e.target === overlay) close(false); }
        function onKey(e) {
            if (e.key === 'Escape') close(false);
            if (e.key === 'Enter' && document.activeElement === confirmBtn) close(true);
        }

        confirmBtn.addEventListener('click', onConfirm);
        cancelBtn.addEventListener('click', onCancel);
        overlay.addEventListener('click', onBackdrop);
        document.addEventListener('keydown', onKey);
    });
}

// ==========================================
// Toast System
// ==========================================
const MAX_TOASTS = 3;

function removeToast(toast) {
    if (!toast || !toast.parentNode) return;
    toast.classList.add('removing');
    toast.addEventListener('animationend', () => {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
    }, { once: true });
}

function showStatus(msg, type = 'info', { undoable = false } = {}) {
    const container = document.getElementById('toast-container');
    if (!container) return;

    // Max 3 toasts — remove oldest when exceeded
    while (container.children.length >= MAX_TOASTS) {
        removeToast(container.firstElementChild);
    }

    const toast = document.createElement('div');
    const typeClass = { error: 'toast-error', success: 'toast-success', info: 'toast-info' };
    toast.className = `toast ${typeClass[type] || 'toast-info'}`;

    const msgSpan = document.createElement('span');
    msgSpan.textContent = msg;
    toast.appendChild(msgSpan);

    if (undoable) {
        const undoBtn = document.createElement('button');
        undoBtn.className = 'toast-undo';
        undoBtn.textContent = 'Undo';
        const capturedLen = undoStack.length;
        undoBtn.addEventListener('click', () => {
            if (undoStack.length === capturedLen) {
                undo();
                removeToast(toast);
            }
        });
        toast.appendChild(undoBtn);
    }

    container.appendChild(toast);

    // Announce to screen reader
    const srEl = document.getElementById('sr-announcements');
    if (srEl) srEl.textContent = msg;

    // Auto-dismiss with hover pause
    let remaining = 4000;
    let startTime = Date.now();
    let timerId = setTimeout(() => removeToast(toast), remaining);

    toast.addEventListener('mouseenter', () => {
        clearTimeout(timerId);
        remaining = Math.max(0, remaining - (Date.now() - startTime));
    });

    toast.addEventListener('mouseleave', () => {
        remaining = 2000;
        startTime = Date.now();
        timerId = setTimeout(() => removeToast(toast), remaining);
    });
}

// ==========================================
// Page Thumbnails
// ==========================================
function updateThumbnail(pageIndex) {
    if (!thumbnailList) return;
    const card = thumbnailList.children[pageIndex];
    if (!card) return;
    const thumbCanvas = card.querySelector('.thumbnail-canvas');
    if (!thumbCanvas) return;
    const p = pages[pageIndex];
    const pdfCanvas = p.wrapper.querySelector('canvas:not(.draw-layer)');
    const tCtx = thumbCanvas.getContext('2d');
    tCtx.clearRect(0, 0, thumbCanvas.width, thumbCanvas.height);
    tCtx.drawImage(pdfCanvas, 0, 0, thumbCanvas.width, thumbCanvas.height);
    tCtx.drawImage(p.canvas, 0, 0, thumbCanvas.width, thumbCanvas.height);
}

function renderThumbnails() {
    if (!thumbnailList) return;
    thumbnailList.innerHTML = '';
    pages.forEach((p, i) => {
        const card = document.createElement('div');
        card.className = 'thumbnail-card' + (i === activePageIndex ? ' active' : '');
        card.draggable = true;
        card.dataset.pageIndex = i;

        // Mini canvas thumbnail
        const thumbCanvas = document.createElement('canvas');
        thumbCanvas.className = 'thumbnail-canvas';
        const pdfCanvas = p.wrapper.querySelector('canvas:not(.draw-layer)');
        const drawCanvas = p.canvas;
        const aspect = pdfCanvas.height / pdfCanvas.width;
        thumbCanvas.width = 130;
        thumbCanvas.height = Math.round(130 * aspect);
        const tCtx = thumbCanvas.getContext('2d');
        tCtx.drawImage(pdfCanvas, 0, 0, thumbCanvas.width, thumbCanvas.height);
        tCtx.drawImage(drawCanvas, 0, 0, thumbCanvas.width, thumbCanvas.height);

        // Page number
        const numLabel = document.createElement('div');
        numLabel.className = 'thumbnail-label';
        numLabel.textContent = 'Page ' + (i + 1);

        // Action buttons
        const actions = document.createElement('div');
        actions.className = 'thumbnail-actions';

        const rotLeftBtn = document.createElement('button');
        rotLeftBtn.setAttribute('data-tooltip', 'Rotate left');
        rotLeftBtn.textContent = '↺';
        rotLeftBtn.addEventListener('click', () => rotatePage(i, -90));

        const rotRightBtn = document.createElement('button');
        rotRightBtn.setAttribute('data-tooltip', 'Rotate right');
        rotRightBtn.textContent = '↻';
        rotRightBtn.addEventListener('click', () => rotatePage(i, 90));

        const delBtn = document.createElement('button');
        delBtn.className = 'del-page-btn';
        delBtn.setAttribute('data-tooltip', 'Delete page');
        delBtn.textContent = '✕';
        delBtn.disabled = pages.length <= 1;
        delBtn.addEventListener('click', () => deletePage(i));

        actions.appendChild(rotLeftBtn);
        actions.appendChild(rotRightBtn);
        actions.appendChild(delBtn);

        card.appendChild(thumbCanvas);
        card.appendChild(numLabel);
        card.appendChild(actions);

        // Click to scroll to page
        card.addEventListener('click', (e) => {
            if (e.target.closest('.thumbnail-actions')) return;
            p.wrapper.scrollIntoView({ behavior: 'smooth', block: 'center' });
            activePageIndex = i;
            renderThumbnails();
        });

        // Drag-and-drop for reorder
        card.addEventListener('dragstart', (e) => {
            e.dataTransfer.setData('text/plain', i.toString());
            card.classList.add('dragging');
        });
        card.addEventListener('dragend', () => { card.classList.remove('dragging'); });
        card.addEventListener('dragover', (e) => {
            e.preventDefault();
            card.classList.add('drag-over');
        });
        card.addEventListener('dragleave', () => { card.classList.remove('drag-over'); });
        card.addEventListener('drop', (e) => {
            e.preventDefault();
            card.classList.remove('drag-over');
            const fromIdx = parseInt(e.dataTransfer.getData('text/plain'));
            if (fromIdx !== i) reorderPages(fromIdx, i);
        });

        thumbnailList.appendChild(card);
    });
}

// ==========================================
// Page Operations
// ==========================================
async function rotatePage(displayIndex, amount) {
    hideNotePopup();
    cachedTextContent = null;
    pushUndoSnapshot();
    pageRotations[displayIndex] = ((pageRotations[displayIndex] || 0) + amount + 360) % 360;

    // Re-render the page with new rotation
    const origIndex = pageOrder[displayIndex];
    const pdf = await pdfjsLib.getDocument(existingPdfBytes.slice(0)).promise;
    const pdfPage = await pdf.getPage(origIndex + 1);
    const viewport = pdfPage.getViewport({ scale, rotation: pageRotations[displayIndex] });

    const p = pages[displayIndex];
    p.wrapper.style.width = viewport.width + 'px';
    p.wrapper.style.height = viewport.height + 'px';

    const pdfCanvas = p.wrapper.querySelector('canvas:not(.draw-layer)');
    pdfCanvas.width = viewport.width;
    pdfCanvas.height = viewport.height;
    await pdfPage.render({ canvasContext: pdfCanvas.getContext('2d'), viewport }).promise;

    p.canvas.width = viewport.width;
    p.canvas.height = viewport.height;
    redrawAnnotations(displayIndex);
    renderThumbnails();
    clearFormOverlays();
    await detectFormFields();
}

async function reorderPages(fromIdx, toIdx) {
    hideNotePopup();
    pushUndoSnapshot();

    // Reorder pages array
    const [movedPage] = pages.splice(fromIdx, 1);
    pages.splice(toIdx, 0, movedPage);

    // Reorder pageOrder and pageRotations
    const [movedOrder] = pageOrder.splice(fromIdx, 1);
    pageOrder.splice(toIdx, 0, movedOrder);
    const [movedRot] = pageRotations.splice(fromIdx, 1);
    pageRotations.splice(toIdx, 0, movedRot);
    cachedTextContent = null;

    // Build old→new index map, then apply in one pass
    const n = pages.length;
    const oldIndices = Array.from({ length: n }, (_, i) => i);
    const [moved] = oldIndices.splice(fromIdx, 1);
    oldIndices.splice(toIdx, 0, moved);
    // oldIndices[newPos] = oldPos → invert to remap[oldPos] = newPos
    const remap = new Array(n);
    for (let i = 0; i < n; i++) remap[oldIndices[i]] = i;
    for (const ann of annotations) {
        ann.pageIndex = remap[ann.pageIndex];
    }

    // Re-sort DOM
    pagesContainer.innerHTML = '';
    pages.forEach(p => pagesContainer.appendChild(p.wrapper));

    // Re-bind canvas events with correct pageIndex
    pageCleanups.forEach(fn => fn());
    pageCleanups = [];
    pages.forEach((p, i) => {
        pageCleanups.push(setupCanvasEvents(p.canvas, i));
    });

    selectedIndex = -1;
    updateDeleteButton();
    redrawAnnotations();
    renderThumbnails();
    if (pageIndicator) { pageIndicator.setupObserver(); }
    clearFormOverlays();
    await detectFormFields();
}

async function deletePage(displayIndex) {
    hideNotePopup();
    if (pages.length <= 1) return;

    const annCount = annotations.filter(a => a.pageIndex === displayIndex).length;
    if (annCount > 0) {
        const ok = await confirmDialog({
            title: 'Delete page?',
            message: 'Page ' + (displayIndex + 1) + ' has ' + annCount + ' annotation(s) that will be permanently removed.',
            confirmText: 'Delete',
            destructive: true,
        });
        if (!ok) return;
    }

    pushUndoSnapshot();

    if (pageCleanups[displayIndex]) pageCleanups[displayIndex]();
    pageCleanups.splice(displayIndex, 1);

    // Remove page
    const removed = pages.splice(displayIndex, 1)[0];
    removed.wrapper.remove();
    pageOrder.splice(displayIndex, 1);
    pageRotations.splice(displayIndex, 1);
    cachedTextContent = null;

    // Remove annotations on deleted page, re-index the rest
    annotations = annotations.filter(a => a.pageIndex !== displayIndex);
    for (const ann of annotations) {
        if (ann.pageIndex > displayIndex) ann.pageIndex--;
    }

    // Reset selection unconditionally — indices have shifted
    selectedIndex = -1;
    if (editingIndex !== -1) {
        hideTextInput();
        editingIndex = -1;
    }

    activePageIndex = Math.min(activePageIndex, pages.length - 1);
    updateDeleteButton();
    redrawAnnotations();
    renderThumbnails();
    if (pageIndicator) { pageIndicator.updateTotal(); pageIndicator.setupObserver(); }
    clearFormOverlays();
    detectFormFields();
    showStatus('Page deleted', 'info', { undoable: true });
}

// ==========================================
// Merge / Split
// ==========================================
async function mergePDF() {
    const mergeInput = document.getElementById('pdf-merge-input');
    if (!mergeInput || !existingPdfBytes) {
        showStatus('Upload a PDF first', 'error');
        return;
    }

    // Remove any stale listener from a prior cancelled file dialog
    mergeInput.onchange = null;
    mergeInput.onchange = async (e) => {
        mergeInput.onchange = null; // Self-remove
        const file = e.target.files[0];
        if (!file) return;
        mergeInput.value = '';

        try {
            showLoading('Merging documents...');
            const newBytes = await file.arrayBuffer();
            const newPdf = await pdfjsLib.getDocument(newBytes.slice(0)).promise;
            const prevCount = pages.length;

            for (let i = 1; i <= newPdf.numPages; i++) {
                const pdfPage = await newPdf.getPage(i);
                const viewport = pdfPage.getViewport({ scale });

                const wrapper = document.createElement('div');
                wrapper.className = 'page-wrapper';
                wrapper.style.width = viewport.width + 'px';
                wrapper.style.height = viewport.height + 'px';

                const pdfCanvas = document.createElement('canvas');
                pdfCanvas.width = viewport.width;
                pdfCanvas.height = viewport.height;
                await pdfPage.render({ canvasContext: pdfCanvas.getContext('2d'), viewport }).promise;

                const drawCanvas = document.createElement('canvas');
                drawCanvas.className = 'draw-layer';
                drawCanvas.width = viewport.width;
                drawCanvas.height = viewport.height;

                wrapper.appendChild(pdfCanvas);
                wrapper.appendChild(drawCanvas);
                pagesContainer.appendChild(wrapper);

                const pageObj = {
                    wrapper,
                    canvas: drawCanvas,
                    ctx: drawCanvas.getContext('2d'),
                };
                pages.push(pageObj);
                pageCleanups.push(setupCanvasEvents(drawCanvas, prevCount + i - 1));
                pageRotations.push(0);
                pageOrder.push(prevCount + i - 1);
            }

            // Merge pdf-lib documents
            const mainDoc = await PDFDocument.load(existingPdfBytes);
            const mergeDoc = await PDFDocument.load(newBytes);
            const copiedPages = await mainDoc.copyPages(mergeDoc, mergeDoc.getPageIndices());
            copiedPages.forEach(p => mainDoc.addPage(p));
            existingPdfBytes = await mainDoc.save();

            // Merge changes existingPdfBytes — clear undo stack since
            // snapshots don't include the PDF bytes and would be inconsistent
            undoStack = [];
            redoStack = [];
            cachedTextContent = null;
            updateUndoRedoButtons();

            renderThumbnails();
            if (pageIndicator) { pageIndicator.updateTotal(); pageIndicator.setupObserver(); }
            hideLoading();
            showStatus('Merged ' + newPdf.numPages + ' pages', 'success');
        } catch (err) {
            hideLoading();
            console.error(err);
            showStatus('Error merging: ' + err.message, 'error');
        }
    };
    mergeInput.click();
}

let removeSplitTrap = null;
let splitModalKeyHandler = null;

function showSplitModal() {
    if (!splitModal || !existingPdfBytes) {
        showStatus('Upload a PDF first', 'error');
        return;
    }
    splitModal.classList.add('active');
    removeSplitTrap = trapFocus(splitModal);
    const input = document.getElementById('split-range-input');
    const preview = document.getElementById('split-preview');
    const errorEl = document.getElementById('split-error');
    if (input) input.value = '';
    if (preview) preview.textContent = 'Total pages: ' + pages.length;
    if (errorEl) errorEl.textContent = '';

    splitModalKeyHandler = (e) => { if (e.key === 'Escape') hideSplitModal(); };
    document.addEventListener('keydown', splitModalKeyHandler);
}

function hideSplitModal() {
    if (splitModal) splitModal.classList.remove('active');
    if (removeSplitTrap) { removeSplitTrap(); removeSplitTrap = null; }
    if (splitModalKeyHandler) { document.removeEventListener('keydown', splitModalKeyHandler); splitModalKeyHandler = null; }
}

function parseSplitRanges(input, totalPages) {
    const ranges = [];
    const parts = input.split(',').map(s => s.trim()).filter(Boolean);
    for (const part of parts) {
        const match = part.match(/^(\d+)(?:-(\d+))?$/);
        if (!match) return { error: 'Invalid range: "' + part + '"' };
        const start = parseInt(match[1]);
        const end = match[2] ? parseInt(match[2]) : start;
        if (start < 1 || end < 1 || start > totalPages || end > totalPages) {
            return { error: 'Range out of bounds: "' + part + '" (1-' + totalPages + ')' };
        }
        if (start > end) return { error: 'Invalid range: start > end in "' + part + '"' };
        ranges.push({ start, end });
    }
    return { ranges };
}

async function splitPDF() {
    const input = document.getElementById('split-range-input');
    const errorEl = document.getElementById('split-error');
    if (!input || !input.value.trim()) {
        if (errorEl) errorEl.textContent = 'Enter page ranges';
        return;
    }

    const { ranges, error } = parseSplitRanges(input.value, pages.length);
    if (error) {
        if (errorEl) errorEl.textContent = error;
        return;
    }

    showLoading('Splitting pages...');
    try {
        // Fill form fields before splitting
        let splitWorkingBytes = existingPdfBytes;
        if (formOverlays.length > 0) {
            try {
                const formDoc = await PDFDocument.load(existingPdfBytes, { ignoreEncryption: true });
                const form = formDoc.getForm();
                for (const overlay of formOverlays) {
                    try {
                        const field = form.getField(overlay.fieldName);
                        if (overlay.fieldType === 'PDFTextField') field.setText(overlay.element.value);
                        else if (overlay.fieldType === 'PDFCheckBox') {
                            if (overlay.element.checked) field.check(); else field.uncheck();
                        } else if (overlay.fieldType === 'PDFDropdown') field.select(overlay.element.value);
                    } catch (e) { console.warn('Could not fill field:', overlay.fieldName, e); }
                }
                form.flatten();
                splitWorkingBytes = await formDoc.save();
            } catch (e) { console.warn('Form filling error during split:', e); }
        }

        for (let ri = 0; ri < ranges.length; ri++) {
            const { start, end } = ranges[ri];
            const srcDoc = await PDFDocument.load(splitWorkingBytes, { ignoreEncryption: true });
            const newDoc = await PDFDocument.create();
            const indices = [];
            for (let p = start - 1; p < end; p++) indices.push(pageOrder[p]);
            const copied = await newDoc.copyPages(srcDoc, indices);
            copied.forEach(p => newDoc.addPage(p));

            // Collect annotations for these pages and embed only needed fonts
            const rangeAnns = annotations.filter(a => a.pageIndex >= start - 1 && a.pageIndex < end);
            const embeddedFonts = await embedUsedFonts(newDoc, rangeAnns);

            // Embed signature images for this range
            const splitSigImages = new Map();
            for (const ann of rangeAnns) {
                if (ann.type === 'signature' && !splitSigImages.has(ann.dataUrl)) {
                    const pngBytes = await fetch(ann.dataUrl).then(r => r.arrayBuffer());
                    const pngImage = await newDoc.embedPng(pngBytes);
                    splitSigImages.set(ann.dataUrl, pngImage);
                }
            }

            // Bake annotations for these pages
            const newPages = newDoc.getPages();
            for (let pi = 0; pi < indices.length; pi++) {
                const displayIdx = start - 1 + pi;
                const pageAnns = annotations.filter(a => a.pageIndex === displayIdx);
                const docPage = newPages[pi];
                const { width: rawW, height: rawH } = docPage.getSize();
                const rot = pageRotations[displayIdx] || 0;
                const isRot = (rot === 90 || rot === 270);
                const pdfH = isRot ? rawW : rawH;

                if (rot) {
                    docPage.setRotation(degrees(rot));
                }

                for (const ann of pageAnns) {
                    const font = embeddedFonts[ann.fontFamily || 'Helvetica'];
                    exportAnnotation(docPage, ann, pdfH, font, splitSigImages);
                }
            }

            const bytes = await newDoc.save();
            const blob = new Blob([bytes], { type: 'application/pdf' });
            const url = URL.createObjectURL(blob);
            const rangeName = start === end ? '' + start : start + '-' + end;

            setTimeout(() => {
                const link = document.createElement('a');
                link.href = url;
                link.download = 'document_pages_' + rangeName + '.pdf';
                link.click();
                setTimeout(() => URL.revokeObjectURL(url), 1000);
            }, ri * 500);
        }
        hideSplitModal();
        hideLoading();
        showStatus('Split into ' + ranges.length + ' file(s)', 'success');
    } catch (err) {
        hideLoading();
        console.error(err);
        if (errorEl) errorEl.textContent = 'Error: ' + err.message;
    }
}

// ==========================================
// Annotation Export (shared by download and split)
// ==========================================
async function embedUsedFonts(pdfDoc, anns) {
    const FONT_PDF_MAP = {
        'Helvetica': StandardFonts.Helvetica,
        'TimesRoman': StandardFonts.TimesRoman,
        'Courier': StandardFonts.Courier,
    };
    const usedFontNames = new Set(
        anns.filter(a => a.type === 'text' || a.type === 'stamp' || (a.type === 'note' && a.text)).map(a => a.fontFamily || 'Helvetica')
    );
    const fonts = {};
    for (const name of usedFontNames) {
        fonts[name] = await pdfDoc.embedFont(FONT_PDF_MAP[name]);
    }
    return fonts;
}

function exportAnnotation(page, ann, pageHeight, font, signatureImages) {
    // Annotations are stored in PDF-space, so coordinates map 1:1 with PDF points.
    const h = pageHeight;

    if (ann.type === 'rect') {
        const opts = {
            x: ann.x,
            y: h - (ann.y + ann.h),
            width: ann.w,
            height: ann.h,
            opacity: ann.opacity ?? 1,
        };
        if (ann.filled) opts.color = ann.color;
        else {
            opts.borderColor = ann.color;
            opts.borderWidth = ann.thickness ?? 2;
        }
        page.drawRectangle(opts);
    } else if (ann.type === 'ellipse') {
        const opts = {
            x: ann.x + ann.w / 2,
            y: h - (ann.y + ann.h / 2),
            xScale: ann.w / 2,
            yScale: ann.h / 2,
            opacity: ann.opacity ?? 1,
        };
        if (ann.filled) opts.color = ann.color;
        else {
            opts.borderColor = ann.color;
            opts.borderWidth = ann.thickness ?? 2;
        }
        page.drawEllipse(opts);
    } else if (ann.type === 'scribble' && ann.points.length > 1) {
        for (let i = 1; i < ann.points.length; i++) {
            page.drawLine({
                start: { x: ann.points[i - 1].x, y: h - ann.points[i - 1].y },
                end: { x: ann.points[i].x, y: h - ann.points[i].y },
                thickness: ann.thickness ?? 2,
                color: ann.color,
                opacity: ann.opacity ?? 1,
            });
        }
    } else if (ann.type === 'line' || ann.type === 'arrow') {
        const sx = ann.x1, sy = h - ann.y1;
        const ex = ann.x2, ey = h - ann.y2;
        page.drawLine({
            start: { x: sx, y: sy },
            end: { x: ex, y: ey },
            thickness: ann.thickness ?? 2,
            color: ann.color,
            opacity: ann.opacity ?? 1,
        });
        if (ann.type === 'arrow') {
            const headLen = Math.max((ann.thickness ?? 2) * 3, 10);
            const angle = Math.atan2(ey - sy, ex - sx);
            const lx = ex - headLen * Math.cos(angle - Math.PI / 6);
            const ly = ey - headLen * Math.sin(angle - Math.PI / 6);
            const rx = ex - headLen * Math.cos(angle + Math.PI / 6);
            const ry = ey - headLen * Math.sin(angle + Math.PI / 6);
            // Draw arrowhead as two lines (avoids drawSvgPath Y-axis issues)
            const headOpts = { thickness: ann.thickness ?? 2, color: ann.color, opacity: ann.opacity ?? 1 };
            page.drawLine({ start: { x: ex, y: ey }, end: { x: lx, y: ly }, ...headOpts });
            page.drawLine({ start: { x: ex, y: ey }, end: { x: rx, y: ry }, ...headOpts });
            page.drawLine({ start: { x: lx, y: ly }, end: { x: rx, y: ry }, ...headOpts });
        }
    } else if (ann.type === 'text') {
        page.drawText(ann.text, {
            x: ann.x,
            y: h - ann.y - ann.size,
            size: ann.size,
            font: font,
            color: ann.color,
            opacity: ann.opacity ?? 1,
        });
    } else if (ann.type === 'signature' && signatureImages) {
        // Note: color tinting is canvas-only; PDF export renders signature as original ink color.
        // pdf-lib doesn't support image tinting or blend modes for this purpose.
        const img = signatureImages.get(ann.dataUrl);
        if (img) {
            page.drawImage(img, {
                x: ann.x,
                y: h - ann.y - ann.h,
                width: ann.w,
                height: ann.h,
                opacity: ann.opacity ?? 1,
            });
        }
    } else if (ann.type === 'redact') {
        page.drawRectangle({
            x: ann.x,
            y: h - ann.y - ann.h,
            width: ann.w,
            height: ann.h,
            color: rgb(0, 0, 0),
            opacity: 1,
        });
    } else if (ann.type === 'stamp') {
        const stampFont = font || null;
        const stampSize = ann.h * 0.6;
        const stampX = ann.x + ann.w / 2;
        const stampY = h - ann.y - ann.h / 2;
        // Draw border rectangle with rotation
        const pad = stampSize * 0.4;
        // Approximate text width using font if available
        let textWidth = ann.w - pad * 2;
        if (stampFont) {
            textWidth = stampFont.widthOfTextAtSize(ann.text, stampSize);
        }
        page.drawRectangle({
            x: stampX - textWidth / 2 - pad,
            y: stampY - stampSize / 2 - pad,
            width: textWidth + pad * 2,
            height: stampSize + pad * 2,
            borderColor: ann.color,
            borderWidth: 3,
            opacity: ann.opacity ?? 0.7,
            rotate: degrees(-15),
        });
        if (stampFont) {
            page.drawText(ann.text, {
                x: stampX - textWidth / 2,
                y: stampY - stampSize / 2,
                size: stampSize,
                font: stampFont,
                color: ann.color,
                opacity: ann.opacity ?? 0.7,
                rotate: degrees(-15),
            });
        }
    } else if (ann.type === 'note') {
        page.drawRectangle({
            x: ann.x,
            y: h - ann.y - ann.h,
            width: ann.w,
            height: ann.h,
            color: ann.color,
            opacity: ann.opacity ?? 1,
        });
        if (ann.text && font) {
            const noteFontSize = 8;
            page.drawText(ann.text, {
                x: ann.x + ann.w + 4,
                y: h - ann.y - noteFontSize,
                size: noteFontSize,
                font: font,
                color: ann.color,
                opacity: ann.opacity ?? 1,
                maxWidth: 150,
            });
        }
    }
}

// ==========================================
// Find Text
// ==========================================
function openFindBar() {
    const bar = document.getElementById('find-bar');
    const input = document.getElementById('find-input');
    if (!bar || !input) return;
    bar.classList.remove('hidden');
    input.focus();
    input.select();
}

function closeFindBar() {
    if (findDebounceTimer) { clearTimeout(findDebounceTimer); findDebounceTimer = null; }
    const bar = document.getElementById('find-bar');
    if (bar) bar.classList.add('hidden');
    findMatches = [];
    findCurrentIndex = -1;
    findQuery = '';
    redrawAnnotations();
}

let findDebounceTimer = null;

async function performSearch(query) {
    findMatches = [];
    findCurrentIndex = -1;
    findQuery = query.toLowerCase();
    const gen = ++searchGeneration;

    if (!findQuery || !existingPdfBytes) {
        updateFindUI();
        redrawAnnotations();
        return;
    }

    // Cache text content so we don't re-parse the PDF on every keystroke
    if (!cachedTextContent) {
        const pdf = await pdfjsLib.getDocument(existingPdfBytes.slice(0)).promise;
        if (gen !== searchGeneration) return; // Superseded by newer search
        cachedTextContent = [];
        for (let i = 0; i < pages.length; i++) {
            const origIdx = pageOrder[i];
            const pdfPage = await pdf.getPage(origIdx + 1);
            const textContent = await pdfPage.getTextContent();
            const viewport = pdfPage.getViewport({ scale: 1 });
            if (gen !== searchGeneration) return;
            cachedTextContent.push({ items: textContent.items, viewportHeight: viewport.height });
        }
    }

    for (let i = 0; i < cachedTextContent.length; i++) {
        const { items, viewportHeight } = cachedTextContent[i];
        for (const item of items) {
            const text = item.str.toLowerCase();
            let startIdx = 0;
            while ((startIdx = text.indexOf(findQuery, startIdx)) !== -1) {
                const tx = item.transform[4];
                const ty = item.transform[5];
                const fontSize = item.transform[0];

                const charWidth = item.width / item.str.length;
                const matchX = tx + startIdx * charWidth;
                const matchW = findQuery.length * charWidth;
                const matchH = item.height || fontSize;
                const matchY = viewportHeight - ty - matchH;

                findMatches.push({
                    pageIndex: i,
                    x: matchX,
                    y: matchY,
                    w: matchW,
                    h: matchH,
                });
                startIdx += findQuery.length;
            }
        }
    }

    if (gen !== searchGeneration) return;
    if (findMatches.length > 0) findCurrentIndex = 0;
    updateFindUI();
    redrawAnnotations();
    if (findMatches.length > 0) scrollToMatch(0);
}

function updateFindUI() {
    const countEl = document.getElementById('find-count');
    const prevBtn = document.getElementById('find-prev');
    const nextBtn = document.getElementById('find-next');

    if (findMatches.length === 0) {
        if (countEl) countEl.textContent = findQuery ? 'No matches' : '';
        if (prevBtn) prevBtn.disabled = true;
        if (nextBtn) nextBtn.disabled = true;
    } else {
        if (countEl) countEl.textContent = (findCurrentIndex + 1) + ' of ' + findMatches.length;
        if (prevBtn) prevBtn.disabled = findMatches.length <= 1;
        if (nextBtn) nextBtn.disabled = findMatches.length <= 1;
    }
}

function scrollToMatch(index) {
    const match = findMatches[index];
    if (!match) return;
    findCurrentIndex = index;
    updateFindUI();

    const wrapper = pages[match.pageIndex].wrapper;
    wrapper.scrollIntoView({ behavior: 'smooth', block: 'center' });

    redrawAnnotations();
}

function initFindBar() {
    const input = document.getElementById('find-input');
    const prevBtn = document.getElementById('find-prev');
    const nextBtn = document.getElementById('find-next');
    const closeBtn = document.getElementById('find-close');

    if (input) {
        input.addEventListener('input', () => {
            clearTimeout(findDebounceTimer);
            findDebounceTimer = setTimeout(() => {
                performSearch(input.value.trim());
            }, 300);
        });

        input.addEventListener('keydown', (e) => {
            e.stopPropagation();
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                if (findMatches.length > 0) {
                    findCurrentIndex = (findCurrentIndex + 1) % findMatches.length;
                    scrollToMatch(findCurrentIndex);
                }
            } else if (e.key === 'Enter' && e.shiftKey) {
                e.preventDefault();
                if (findMatches.length > 0) {
                    findCurrentIndex = (findCurrentIndex - 1 + findMatches.length) % findMatches.length;
                    scrollToMatch(findCurrentIndex);
                }
            } else if (e.key === 'Escape') {
                e.preventDefault();
                closeFindBar();
            }
        });
    }

    if (prevBtn) {
        prevBtn.addEventListener('click', () => {
            if (findMatches.length === 0) return;
            findCurrentIndex = (findCurrentIndex - 1 + findMatches.length) % findMatches.length;
            scrollToMatch(findCurrentIndex);
        });
    }

    if (nextBtn) {
        nextBtn.addEventListener('click', () => {
            if (findMatches.length === 0) return;
            findCurrentIndex = (findCurrentIndex + 1) % findMatches.length;
            scrollToMatch(findCurrentIndex);
        });
    }

    if (closeBtn) {
        closeBtn.addEventListener('click', () => closeFindBar());
    }
}

// ==========================================
// Sticky Note Popup
// ==========================================
function showNotePopup(annIndex) {
    hideNotePopup();
    const ann = annotations[annIndex];
    if (!ann || ann.type !== 'note') return;

    let undoPushed = false;

    const popup = document.createElement('div');
    popup.className = 'sticky-note-popup';

    const header = document.createElement('div');
    header.className = 'sticky-note-header';
    header.textContent = 'Note';
    const closeBtn = document.createElement('button');
    closeBtn.className = 'sticky-note-close';
    closeBtn.textContent = '\u00d7';
    closeBtn.addEventListener('click', () => hideNotePopup());
    header.appendChild(closeBtn);

    const body = document.createElement('div');
    body.className = 'sticky-note-body';
    const textarea = document.createElement('textarea');
    textarea.value = ann.text || '';
    textarea.placeholder = 'Add a comment...';
    textarea.addEventListener('input', () => {
        if (!undoPushed) { pushUndoSnapshot(); undoPushed = true; }
        ann.text = textarea.value;
    });
    textarea.addEventListener('keydown', (e) => {
        // Prevent global shortcuts while typing in the note
        e.stopPropagation();
        if (e.key === 'Escape') hideNotePopup();
    });
    body.appendChild(textarea);

    popup.appendChild(header);
    popup.appendChild(body);

    // Position near the icon
    const wrapper = pages[ann.pageIndex].wrapper;
    const iconX = pdfToCanvas(ann.x);
    const iconY = pdfToCanvas(ann.y);
    popup.style.left = (iconX + pdfToCanvas(ann.w) + 5) + 'px';
    popup.style.top = iconY + 'px';

    wrapper.appendChild(popup);
    activeNotePopup = { element: popup, annIndex };
    textarea.focus();
}

function hideNotePopup() {
    if (activeNotePopup) {
        activeNotePopup.element.remove();
        activeNotePopup = null;
        redrawAnnotations();
    }
}

// ==========================================
// Crop Pages
// ==========================================
async function applyCrop() {
    if (!cropRect) return;

    showLoading('Cropping page...');
    try {
        pushUndoSnapshot();
        const { pageIndex, x, y, w, h } = cropRect;

        const pdfDoc = await PDFDocument.load(existingPdfBytes, { ignoreEncryption: true });
        const origIdx = pageOrder[pageIndex];
        const page = pdfDoc.getPages()[origIdx];

        const { height: pageH } = page.getSize();

        // CropBox uses PDF coordinate system (origin bottom-left)
        // Our y is in PDF-space with top-left origin, so convert
        const cropBoxY = pageH - y - h;

        page.setCropBox(x, cropBoxY, w, h);

        existingPdfBytes = await pdfDoc.save();

        // Shift annotations on the cropped page to match new origin
        for (const ann of annotations) {
            if (ann.pageIndex === pageIndex) {
                ann.x -= x;
                ann.y -= y;
                if (ann.points) {
                    for (const p of ann.points) { p.x -= x; p.y -= y; }
                }
                if (ann.x1 !== undefined) {
                    ann.x1 -= x; ann.y1 -= y;
                    ann.x2 -= x; ann.y2 -= y;
                }
            }
        }

        cropRect = null;
        cachedTextContent = null;
        document.getElementById('crop-toolbar')?.classList.add('hidden');

        // Clear form overlays (page geometry changed)
        clearFormOverlays();

        await reloadPagesFromBytes();
        redrawAnnotations();
        renderThumbnails();
        await detectFormFields();

        if (pageIndicator) {
            pageIndicator.setupObserver();
        }

        showStatus('Page cropped', 'success');
    } catch (err) {
        console.error(err);
        showStatus('Error cropping: ' + err.message, 'error');
    } finally {
        hideLoading();
    }

    setTool('select');
}

function cancelCrop() {
    cropRect = null;
    document.getElementById('crop-toolbar')?.classList.add('hidden');
    redrawAnnotations();
    setTool('select');
}

// ==========================================
// Form Field Detection & Filling
// ==========================================
function clearFormOverlays() {
    for (const overlay of formOverlays) {
        overlay.element.remove();
    }
    formOverlays = [];
}

function saveFormValues() {
    const values = {};
    for (const overlay of formOverlays) {
        if (overlay.fieldType === 'PDFCheckBox') {
            values[overlay.fieldName] = overlay.element.checked;
        } else {
            values[overlay.fieldName] = overlay.element.value;
        }
    }
    return values;
}

function restoreFormValues(values) {
    for (const overlay of formOverlays) {
        if (overlay.fieldName in values) {
            if (overlay.fieldType === 'PDFCheckBox') {
                overlay.element.checked = values[overlay.fieldName];
            } else {
                overlay.element.value = values[overlay.fieldName];
            }
        }
    }
}

async function detectFormFields() {
    if (!existingPdfBytes) return;

    try {
        const pdfDoc = await PDFDocument.load(existingPdfBytes, { ignoreEncryption: true });
        const form = pdfDoc.getForm();
        const fields = form.getFields();

        if (fields.length === 0) return;

        clearFormOverlays();

        for (const field of fields) {
            const fieldName = field.getName();
            const fieldType = field.constructor.name;
            let widgets;
            try {
                widgets = field.acroField.getWidgets();
            } catch (e) {
                continue;
            }

            for (const widget of widgets) {
                let rect;
                try {
                    rect = widget.getRectangle();
                } catch (e) {
                    continue;
                }

                // Determine which page this widget belongs to
                let pageIdx = 0;
                try {
                    const pageRef = widget.P();
                    const allPages = pdfDoc.getPages();
                    for (let i = 0; i < allPages.length; i++) {
                        if (allPages[i].ref === pageRef) {
                            pageIdx = i;
                            break;
                        }
                    }
                } catch (e) {
                    // Default to page 0
                }

                // Map original page index through pageOrder to find display index
                const displayIdx = pageOrder.indexOf(pageIdx);
                if (displayIdx === -1) continue;

                createFormOverlay(field, fieldName, fieldType, rect, displayIdx);
            }
        }
    } catch (e) {
        // Silently fail — PDF may not have form fields or may be malformed
        console.warn('Form detection skipped:', e.message);
    }
}

function createFormOverlay(field, fieldName, fieldType, rect, displayIdx) {
    const wrapper = pages[displayIdx].wrapper;
    const page = pages[displayIdx];
    // PDF page height in PDF-space: canvas height divided by scale
    const canvasHeight = page.canvas.height / scale;

    let input;
    if (fieldType === 'PDFDropdown') {
        input = document.createElement('select');
    } else {
        input = document.createElement('input');
    }

    // Convert PDF coordinates (bottom-left origin) to canvas coordinates (top-left origin)
    const left = pdfToCanvas(rect.x);
    const top = pdfToCanvas(canvasHeight - rect.y - rect.height);
    const width = pdfToCanvas(rect.width);
    const height = pdfToCanvas(rect.height);

    input.style.left = left + 'px';
    input.style.top = top + 'px';
    input.style.width = width + 'px';
    input.style.height = height + 'px';
    input.className = 'form-field-overlay';
    input.dataset.fieldName = fieldName;

    if (fieldType === 'PDFCheckBox') {
        input.type = 'checkbox';
        input.className = 'form-field-overlay form-checkbox-overlay';
        try { input.checked = field.isChecked(); } catch (e) { console.debug('Form field init:', fieldName, e); }
    } else if (fieldType === 'PDFTextField') {
        input.type = 'text';
        try { input.value = field.getText() || ''; } catch (e) { console.debug('Form field init:', fieldName, e); }
        try { input.placeholder = fieldName; } catch (e) { console.debug('Form field init:', fieldName, e); }
    } else if (fieldType === 'PDFDropdown') {
        try {
            const options = field.getOptions();
            options.forEach(opt => {
                const o = document.createElement('option');
                o.value = opt;
                o.textContent = opt;
                input.appendChild(o);
            });
            try { input.value = field.getSelected()?.[0] || ''; } catch (e) { console.debug('Form field init:', fieldName, e); }
        } catch (e) { console.debug('Form field init:', fieldName, e); }
    } else {
        // Unsupported field type, skip
        return;
    }

    // Prevent click events from bubbling to the draw layer
    input.addEventListener('mousedown', (e) => e.stopPropagation());
    input.addEventListener('mouseup', (e) => e.stopPropagation());
    input.addEventListener('click', (e) => e.stopPropagation());
    input.addEventListener('keydown', (e) => e.stopPropagation());

    wrapper.appendChild(input);
    formOverlays.push({ element: input, fieldName, fieldType, pageIndex: displayIdx });
}

// ==========================================
// PDF Download
// ==========================================
async function downloadPDF() {
    if (!existingPdfBytes) {
        showStatus("Please upload a PDF first", "error");
        return;
    }

    try {
        showLoading('Saving PDF...');
        if (textInput && !textInput.classList.contains('hidden')) commitText();

        // Fill form fields from overlays before building the output
        let workingBytes = existingPdfBytes;
        if (formOverlays.length > 0) {
            try {
                const formDoc = await PDFDocument.load(existingPdfBytes, { ignoreEncryption: true });
                const form = formDoc.getForm();
                for (const overlay of formOverlays) {
                    try {
                        const field = form.getField(overlay.fieldName);
                        if (overlay.fieldType === 'PDFTextField') field.setText(overlay.element.value);
                        else if (overlay.fieldType === 'PDFCheckBox') {
                            if (overlay.element.checked) field.check(); else field.uncheck();
                        } else if (overlay.fieldType === 'PDFDropdown') field.select(overlay.element.value);
                    } catch (e) { console.warn('Could not fill field:', overlay.fieldName, e); }
                }
                form.flatten();
                workingBytes = await formDoc.save();
            } catch (e) { console.warn('Form filling error:', e); }
        }

        const srcDoc = await PDFDocument.load(workingBytes, { ignoreEncryption: true });
        const outDoc = await PDFDocument.create();
        const embeddedFonts = await embedUsedFonts(outDoc, annotations);

        // Embed signature images
        const signatureImages = new Map();
        for (const ann of annotations) {
            if (ann.type === 'signature' && !signatureImages.has(ann.dataUrl)) {
                const pngBytes = await fetch(ann.dataUrl).then(r => r.arrayBuffer());
                const pngImage = await outDoc.embedPng(pngBytes);
                signatureImages.set(ann.dataUrl, pngImage);
            }
        }

        for (let i = 0; i < pageOrder.length; i++) {
            const origIdx = pageOrder[i];
            const [copiedPage] = await outDoc.copyPages(srcDoc, [origIdx]);
            outDoc.addPage(copiedPage);
            const outPage = outDoc.getPages()[i];

            if (pageRotations[i]) {
                outPage.setRotation(degrees(pageRotations[i]));
            }

            const { width: rawW, height: rawH } = outPage.getSize();
            const rot = pageRotations[i] || 0;
            const isRotated = (rot === 90 || rot === 270);
            const pdfH = isRotated ? rawW : rawH;
            const pageAnns = annotations.filter(a => a.pageIndex === i);

            for (const ann of pageAnns) {
                const font = embeddedFonts[ann.fontFamily || 'Helvetica'];
                exportAnnotation(outPage, ann, pdfH, font, signatureImages);
            }
        }

        const pdfBytes = await outDoc.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'annotated_document.pdf';
        link.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        hideLoading();
        showStatus("PDF downloaded", "success");
    } catch (err) {
        hideLoading();
        console.error(err);
        showStatus("Error saving PDF: " + err.message, "error");
    }
}

init();
