const sourceInput = document.getElementById('sourceInput');
const previewDiv = document.getElementById('previewContent');
const fileInput = document.getElementById('fileInput');

// --- History Management ---
const historyStack = [];
let historyIndex = -1;
let isHistoryAction = false;

function saveHistory() {
    if(isHistoryAction) return;
    if(historyIndex < historyStack.length - 1) {
        historyStack.splice(historyIndex + 1);
    }
    historyStack.push(previewDiv.innerHTML);
    historyIndex++;
    if(historyStack.length > 50) {
        historyStack.shift();
        historyIndex--;
    }
}

function undo() {
    if(historyIndex > 0) {
        isHistoryAction = true;
        historyIndex--;
        previewDiv.innerHTML = historyStack[historyIndex];
        attachEvents(previewDiv);
        isHistoryAction = false;
    }
}

function redo() {
    if(historyIndex < historyStack.length - 1) {
        isHistoryAction = true;
        historyIndex++;
        previewDiv.innerHTML = historyStack[historyIndex];
        attachEvents(previewDiv);
        isHistoryAction = false;
    }
}

// --- File I/O ---
function loadFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        sourceInput.value = e.target.result;
        renderCode(); 
    };
    reader.readAsText(file, 'UTF-8');
}
fileInput.addEventListener('change', (e) => { loadFile(e.target.files[0]); fileInput.value = ''; });
sourceInput.addEventListener('dragover', (e) => { e.preventDefault(); sourceInput.classList.add('drag-over'); });
sourceInput.addEventListener('dragleave', (e) => { e.preventDefault(); sourceInput.classList.remove('drag-over'); });
sourceInput.addEventListener('drop', (e) => { 
    e.preventDefault(); sourceInput.classList.remove('drag-over');
    if(e.dataTransfer.files.length) loadFile(e.dataTransfer.files[0]); 
});

// --- Core Functions ---
function renderCode() {
    if(!sourceInput.value.trim()) { previewDiv.innerHTML = ""; return; }
    previewDiv.innerHTML = sourceInput.value;
    
    // History Init
    historyStack.length = 0; 
    historyIndex = -1;
    saveHistory(); 

    attachEvents(previewDiv);
}

function calculateAll() {
    runCalculationPass();
    runCalculationPass();
    updateTooltips();
}

function runCalculationPass() {
    const tables = previewDiv.querySelectorAll('table');
    if(tables.length === 0) return;
    tables.forEach(table => {
        const grid = mapTableToGrid(table);
        const formulaElements = table.querySelectorAll('[data-dze-formula]');
        formulaElements.forEach(el => {
            let formula = el.getAttribute('data-dze-formula');
            let safeFormula = formula;
            try {
                if(safeFormula.startsWith('=')) safeFormula = safeFormula.substring(1);
                if(safeFormula.includes("SUM")) safeFormula = safeFormula.replace(/SUM\(([^)]+)\)/g, (m, a) => a.includes(':') ? getRangeSum(grid, a).toString() : m);
                safeFormula = safeFormula.replace(/\b[A-Z]+[0-9]+\b/g, (ref) => getVal(grid, ref) < 0 ? `(${getVal(grid, ref)})` : getVal(grid, ref));
                safeFormula = safeFormula.replace(/([0-9.]+)%/g, "$1*0.01");
                if(safeFormula.includes("PRODUCT")) safeFormula = safeFormula.replace(/PRODUCT\(([^)]+)\)/g, (m, a) => { let r=1; a.split(',').forEach(p=>r*=parseFloat(new Function('return '+p)())); return r; });
                if(safeFormula.includes("SUM")) safeFormula = safeFormula.replace(/SUM\(([^)]+)\)/g, (m, a) => { let r=0; a.split(',').forEach(p=>r+=parseFloat(new Function('return '+p)())); return r; });
                let resultValue = new Function('return ' + safeFormula)();
                const pTag = el.querySelector('p') || el;
                const hasSep = el.getAttribute('dze_format_separator') === ',';
                let displayVal = Math.round(resultValue);
                if(hasSep) displayVal = displayVal.toLocaleString();
                pTag.innerText = displayVal;
                const addr = el.getAttribute('data-addr');
                if(addr) grid[addr] = resultValue;
            } catch (e) { }
        });
    });
}

function updateTooltips() {
    const tables = previewDiv.querySelectorAll('table');
    tables.forEach(table => {
        const rows = table.rows; if(rows.length === 0) return;
        const headerMap = {};
        for(let c=0; c<rows[0].cells.length; c++) headerMap[parseInt(rows[0].cells[c].getAttribute('data-col'))] = rows[0].cells[c].innerText.trim();
        table.querySelectorAll('td').forEach(td => {
            if(parseInt(td.getAttribute('data-row')) === 0) return;
            let title = td.getAttribute('data-org-title') || td.getAttribute('title');
            let lines = [];
            if(headerMap[parseInt(td.getAttribute('data-col'))]) lines.push(`[${headerMap[parseInt(td.getAttribute('data-col'))]}]`);
            if(title) lines.push(`Title: ${title}`);
            if(td.getAttribute('data-dze-formula')) lines.push(`𝑓𝑥  ${td.getAttribute('data-dze-formula')}`);
            if(lines.length) { 
                td.setAttribute('data-tooltip', lines.join('\n')); 
                if(td.hasAttribute('title')) { td.setAttribute('data-org-title', td.getAttribute('title')); td.removeAttribute('title'); }
            } else td.removeAttribute('data-tooltip');
        });
    });
}

function attachEvents(container) {
    const tables = container.querySelectorAll('table');
    tables.forEach(table => {
        initCoordinates(table);
        updateTooltips();
        table.addEventListener('mousedown', function(e) {
            const td = e.target.closest('td');
            if (!td) return;
            container.querySelectorAll('.selected-cell').forEach(c => c.classList.remove('selected-cell'));
            const editing = container.querySelector('.editing-cell');
            if (editing && editing !== td) finishEdit(editing); 
            isDragging = true; startCell = td; currentTable = table;
            selectCell(td); previewDiv.focus();
        });
        table.addEventListener('mouseover', function(e) {
            if (!isDragging) return;
            const td = e.target.closest('td');
            if (!td || table !== currentTable) return;
            selectRange(currentTable, startCell, td);
        });
        table.addEventListener('dblclick', function(e) {
            const td = e.target.closest('td');
            if (!td) return;
            if(td.hasAttribute('data-dze-formula')) { alert("수식 셀은 수정 불가"); return; }
            makeEditable(td);
        });
    });

    document.addEventListener('mouseup', () => isDragging = false);

    previewDiv.addEventListener('keydown', function(e) {
        if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) { e.preventDefault(); undo(); return; }
        if (((e.ctrlKey || e.metaKey) && e.key === 'y') || ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'z')) { e.preventDefault(); redo(); return; }

        if (e.key === 'Delete') {
            const selectedCells = previewDiv.querySelectorAll('.selected-cell');
            if(selectedCells.length > 0) {
                selectedCells.forEach(td => {
                    if(!td.hasAttribute('data-dze-formula')) {
                        const pTag = td.querySelector('p');
                        if(pTag) pTag.innerText = ""; else td.innerText = "";
                    }
                });
                calculateAll();
                saveHistory(); 
            }
        }
    });
}

let isDragging = false; let startCell = null; let currentTable = null;

// Helpers
function initCoordinates(table) {
    const matrix = [];
    const rows = table.rows;
    for(let r=0; r<rows.length; r++) {
        if(!matrix[r]) matrix[r] = [];
        const cells = rows[r].cells;
        let colIdx = 0;
        for(let c=0; c<cells.length; c++) {
            while(matrix[r][colIdx]) colIdx++;
            const cell = cells[c];
            const rowspan = cell.rowSpan || 1;
            const colspan = cell.colSpan || 1;
            cell.setAttribute('data-row', r); cell.setAttribute('data-col', colIdx);
            for(let i=0; i<rowspan; i++) for(let j=0; j<colspan; j++) { if(!matrix[r+i]) matrix[r+i]=[]; matrix[r+i][colIdx+j]=true; }
            colIdx += colspan;
        }
    }
}
function mapTableToGrid(table) {
    const grid = {};
    table.querySelectorAll('td').forEach(td => {
        const r = parseInt(td.getAttribute('data-row'));
        const c = parseInt(td.getAttribute('data-col'));
        const address = `${getColLetter(c)}${r + 1}`;
        let text = td.innerText.replace(/,/g, '').trim();
        let val = (text === "") ? 0 : parseFloat(text);
        grid[address] = isNaN(val) ? 0 : val;
        td.setAttribute('data-addr', address);
    });
    return grid;
}
function getColLetter(i) { let l=''; while(i>=0){ l=String.fromCharCode((i%26)+65)+l; i=Math.floor(i/26)-1; } return l; }
function getVal(g, r) { return g[r] !== undefined ? g[r] : 0; }
function getRangeSum(g, r) {
    const [s, e] = r.split(':');
    const sc = s.match(/[A-Z]+/)[0], ec = e.match(/[A-Z]+/)[0];
    const sr = parseInt(s.match(/\d+/)[0]), er = parseInt(e.match(/\d+/)[0]);
    const sci = colToNum(sc), eci = colToNum(ec);
    let sum = 0;
    for(let i=sr; i<=er; i++) for(let j=sci; j<=eci; j++) sum += getVal(g, `${getColLetter(j)}${i}`);
    return sum;
}
function colToNum(c) { let n=0; for(let i=0; i<c.length; i++) n=n*26+c.charCodeAt(i)-64; return n-1; }
function selectRange(t, s, e) {
    t.parentElement.querySelectorAll('.selected-cell').forEach(c => c.classList.remove('selected-cell'));
    const r1=parseInt(s.getAttribute('data-row')), c1=parseInt(s.getAttribute('data-col'));
    const r2=parseInt(e.getAttribute('data-row')), c2=parseInt(e.getAttribute('data-col'));
    t.querySelectorAll('td').forEach(td => {
        const r=parseInt(td.getAttribute('data-row')), c=parseInt(td.getAttribute('data-col'));
        if(r>=Math.min(r1,r2) && r<=Math.max(r1,r2) && c>=Math.min(c1,c2) && c<=Math.max(c1,c2)) td.classList.add('selected-cell');
    });
}
function selectCell(td) { td.classList.add('selected-cell'); }
function makeEditable(td) {
    td.classList.remove('selected-cell'); td.classList.add('editing-cell'); td.setAttribute('contenteditable', 'true'); td.focus();
    document.execCommand('selectAll', false, null);
    td.onkeydown = (e) => { if(e.key === 'Enter') { e.preventDefault(); finishEdit(td); } };
    td.onblur = () => finishEdit(td);
}
function finishEdit(td) {
    td.classList.remove('editing-cell'); td.setAttribute('contenteditable', 'false'); td.onkeydown=null; td.onblur=null;
    calculateAll();
    saveHistory(); 
}
