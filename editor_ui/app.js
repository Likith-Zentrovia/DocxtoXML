/**
 * RittDoc Editor - DOCX to XML
 * Client-side application for viewing and editing RittDoc-compliant DocBook XML
 */

// ============================================================================
// State
// ============================================================================

const state = {
    xml: '',
    html: '',
    title: 'Untitled',
    currentView: 'preview',
    isDirty: false
};

// ============================================================================
// DOM Elements
// ============================================================================

const elements = {
    documentTitle: document.getElementById('documentTitle'),
    previewPanel: document.getElementById('previewPanel'),
    previewContent: document.getElementById('previewContent'),
    xmlPanel: document.getElementById('xmlPanel'),
    xmlEditor: document.getElementById('xmlEditor'),
    resizer: document.getElementById('resizer'),
    statusMessage: document.getElementById('statusMessage'),
    statusInfo: document.getElementById('statusInfo'),
    toastContainer: document.getElementById('toastContainer'),
    btnSave: document.getElementById('btnSave'),
    btnRefresh: document.getElementById('btnRefresh'),
    btnDownload: document.getElementById('btnDownload'),
    downloadMenu: document.getElementById('downloadMenu'),
    btnFormat: document.getElementById('btnFormat'),
    btnValidate: document.getElementById('btnValidate'),
    viewButtons: document.querySelectorAll('.view-btn')
};

// ============================================================================
// Initialization
// ============================================================================

async function init() {
    try {
        setStatus('Loading document...');
        const response = await fetch('/api/init');

        if (!response.ok) {
            throw new Error('Failed to load document');
        }

        const data = await response.json();

        state.xml = data.xml;
        state.html = data.html;
        state.title = data.title;

        // Update UI
        elements.documentTitle.textContent = state.title;
        elements.previewContent.innerHTML = state.html;
        elements.xmlEditor.value = state.xml;

        setStatus('Ready');
        updateStatusInfo();

    } catch (error) {
        console.error('Init error:', error);
        elements.previewContent.innerHTML = `
            <div class="error">
                <h3>Failed to load document</h3>
                <p>${error.message}</p>
            </div>
        `;
        setStatus('Error loading document');
    }
}

// ============================================================================
// View Management
// ============================================================================

function setView(view) {
    state.currentView = view;
    const container = document.querySelector('.editor-container');

    // Remove all view classes
    container.classList.remove('split-view');

    // Update panels visibility
    switch (view) {
        case 'preview':
            elements.previewPanel.classList.remove('hidden');
            elements.xmlPanel.classList.add('hidden');
            break;
        case 'xml':
            elements.previewPanel.classList.add('hidden');
            elements.xmlPanel.classList.remove('hidden');
            break;
        case 'split':
            elements.previewPanel.classList.remove('hidden');
            elements.xmlPanel.classList.remove('hidden');
            container.classList.add('split-view');
            break;
    }

    // Update active button
    elements.viewButtons.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.view === view);
    });
}

// ============================================================================
// Save & Refresh
// ============================================================================

async function save() {
    try {
        setStatus('Saving...');

        const response = await fetch('/api/save', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ xml: state.xml })
        });

        const data = await response.json();

        if (response.ok && data.success) {
            state.isDirty = false;
            showToast('Document saved successfully', 'success');
            setStatus('Saved');
        } else {
            throw new Error(data.error || 'Save failed');
        }

    } catch (error) {
        console.error('Save error:', error);
        showToast(`Save failed: ${error.message}`, 'error');
        setStatus('Save failed');
    }
}

async function refresh() {
    try {
        setStatus('Refreshing preview...');

        // Update state from editor
        state.xml = elements.xmlEditor.value;

        const response = await fetch('/api/render-html', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ xml: state.xml })
        });

        const data = await response.json();

        state.html = data.html;
        elements.previewContent.innerHTML = state.html;

        setStatus('Preview updated');
        showToast('Preview refreshed', 'success');

    } catch (error) {
        console.error('Refresh error:', error);
        showToast(`Refresh failed: ${error.message}`, 'error');
        setStatus('Refresh failed');
    }
}

// ============================================================================
// XML Utilities
// ============================================================================

function formatXML(xml) {
    try {
        // Simple XML formatting
        let formatted = '';
        let indent = 0;
        const lines = xml.replace(/>\s*</g, '>\n<').split('\n');

        lines.forEach(line => {
            line = line.trim();
            if (!line) return;

            // Check if closing tag
            if (line.match(/^<\/\w/)) {
                indent = Math.max(0, indent - 1);
            }

            formatted += '  '.repeat(indent) + line + '\n';

            // Check if opening tag (not self-closing, not closing)
            if (line.match(/^<\w[^>]*[^\/]>.*$/)) {
                indent++;
            }
            // Handle self-closing tags
            if (line.match(/\/>$/)) {
                // Don't change indent
            }
        });

        return formatted.trim();
    } catch (e) {
        return xml;
    }
}

function validateXML(xml) {
    try {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'application/xml');
        const error = doc.querySelector('parsererror');

        if (error) {
            return { valid: false, error: error.textContent };
        }
        return { valid: true };
    } catch (e) {
        return { valid: false, error: e.message };
    }
}

// ============================================================================
// UI Helpers
// ============================================================================

function setStatus(message) {
    elements.statusMessage.textContent = message;
}

function updateStatusInfo() {
    const lines = state.xml.split('\n').length;
    const chars = state.xml.length;
    elements.statusInfo.textContent = `${lines} lines | ${chars} characters`;
}

function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;

    elements.toastContainer.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(100%)';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// ============================================================================
// Resizer
// ============================================================================

function initResizer() {
    let isResizing = false;

    elements.resizer.addEventListener('mousedown', (e) => {
        isResizing = true;
        elements.resizer.classList.add('dragging');
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
    });

    document.addEventListener('mousemove', (e) => {
        if (!isResizing) return;

        const container = document.querySelector('.editor-container');
        const containerRect = container.getBoundingClientRect();
        const percentage = ((e.clientX - containerRect.left) / containerRect.width) * 100;

        if (percentage > 20 && percentage < 80) {
            elements.previewPanel.style.flex = `0 0 ${percentage}%`;
            elements.xmlPanel.style.flex = `0 0 ${100 - percentage}%`;
        }
    });

    document.addEventListener('mouseup', () => {
        if (isResizing) {
            isResizing = false;
            elements.resizer.classList.remove('dragging');
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });
}

// ============================================================================
// Event Listeners
// ============================================================================

function setupEventListeners() {
    // View toggle buttons
    elements.viewButtons.forEach(btn => {
        btn.addEventListener('click', () => setView(btn.dataset.view));
    });

    // Save button
    elements.btnSave.addEventListener('click', save);

    // Refresh button
    elements.btnRefresh.addEventListener('click', refresh);

    // Format button
    elements.btnFormat.addEventListener('click', () => {
        elements.xmlEditor.value = formatXML(elements.xmlEditor.value);
        state.xml = elements.xmlEditor.value;
        showToast('XML formatted', 'success');
    });

    // Validate button
    elements.btnValidate.addEventListener('click', () => {
        const result = validateXML(elements.xmlEditor.value);
        if (result.valid) {
            showToast('XML is valid', 'success');
        } else {
            showToast(`Invalid XML: ${result.error}`, 'error');
        }
    });

    // Download dropdown
    elements.btnDownload.addEventListener('click', () => {
        elements.downloadMenu.classList.toggle('show');
    });

    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
        if (!elements.btnDownload.contains(e.target)) {
            elements.downloadMenu.classList.remove('show');
        }
    });

    // XML editor changes
    elements.xmlEditor.addEventListener('input', () => {
        state.xml = elements.xmlEditor.value;
        state.isDirty = true;
        updateStatusInfo();
    });

    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
        // Ctrl+S or Cmd+S to save
        if ((e.ctrlKey || e.metaKey) && e.key === 's') {
            e.preventDefault();
            save();
        }

        // Ctrl+R or Cmd+R to refresh
        if ((e.ctrlKey || e.metaKey) && e.key === 'r') {
            e.preventDefault();
            refresh();
        }

        // Escape to close dropdowns
        if (e.key === 'Escape') {
            elements.downloadMenu.classList.remove('show');
        }
    });

    // Warn before leaving with unsaved changes
    window.addEventListener('beforeunload', (e) => {
        if (state.isDirty) {
            e.preventDefault();
            e.returnValue = '';
        }
    });
}

// ============================================================================
// Start
// ============================================================================

document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
    initResizer();
    init();
});
