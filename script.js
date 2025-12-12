// --- DOM Elements ---
const sourceInput = document.getElementById('sourceInput');
const previewDiv = document.getElementById('previewContent');
const fileInput = document.getElementById('fileInput');

// --- Global State ---
let isDragging = false;
let startCell = null;
let currentTable = null;

// [Resizing Variables]
let isResizing = false;
let resizeType = null;      
let resizeTarget = null;    
let resizeStartVal = 0;     
let resizeStartPos = 0;     
let resizeUnit = 'px';      
let resizeParentSize = 0;   
let targetColElement = null; 

// --- History Management ---
const historyStack = [];
let historyIndex = -1;
let isHistoryAction = false;

function saveHistory() {
    if(isHistoryAction) return;
    if(historyIndex < historyStack.length - 1) historyStack.splice(historyIndex + 1);
    historyStack.push(previewDiv.innerHTML);
    historyIndex++;
    if(historyStack.length > 50) { historyStack.shift(); historyIndex--; }
}

function undo() {
    if(historyIndex > 0) {
        isHistoryAction = true; historyIndex--;
        previewDiv.innerHTML = historyStack[historyIndex];
        attachEvents(previewDiv); isHistoryAction = false;
    }
}

function redo() {
    if(historyIndex < historyStack.length - 1) {
        isHistoryAction = true; historyIndex++;
        previewDiv.innerHTML = historyStack[historyIndex];
        attachEvents(previewDiv); isHistoryAction = false;
    }
}

// --- File I/O ---
function loadFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) { sourceInput.value = e.target.result; renderCode(); };
    reader.readAsText(file, 'UTF-8');
}
if(fileInput) fileInput.addEventListener('change', (e) => { loadFile(e.target.files[0]); fileInput.value = ''; });
if(sourceInput) {
    sourceInput.addEventListener('dragover', (e) => { e.preventDefault(); sourceInput.classList.add('drag-over'); });
    sourceInput.addEventListener('dragleave', (e) => { e.preventDefault(); sourceInput.classList.remove('drag-over'); });
    sourceInput.addEventListener('drop', (e) => { e.preventDefault(); sourceInput.classList.remove('drag-over'); if(e.dataTransfer.files.length) loadFile(e.dataTransfer.files[0]); });
}

// --- Core Functions ---
function renderCode() {
    if(!sourceInput.value.trim()) { previewDiv.innerHTML = ""; return; }
    
    // HTML 파싱 에러 방지
    const tempContainer = document.createElement('div');
    tempContainer.innerHTML = sourceInput.value;
    previewDiv.innerHTML = tempContainer.innerHTML;
    
    historyStack.length = 0; historyIndex = -1; saveHistory();
    attachEvents(previewDiv);
}

function calculateAll() { runCalculationPass(); runCalculationPass(); updateTooltips(); }

function runCalculationPass() {
    const tables = previewDiv.querySelectorAll('table');
    if(tables.length === 0) return;
    tables.forEach(table => {
        const grid = mapTableToGrid(table);
        table.querySelectorAll('[data-dze-formula]').forEach(el => {
            let formula = el.getAttribute('data-dze-formula'), safeFormula = formula;
            try {
                if(safeFormula.startsWith('=')) safeFormula = safeFormula.substring(1);
                if(safeFormula.includes("SUM")) safeFormula = safeFormula.replace(/SUM\(([^)]+)\)/g, (m, a) => a.includes(':') ? getRangeSum(grid, a).toString() : m);
                safeFormula = safeFormula.replace(/\b[A-Z]+[0-9]+\b/g, (ref) => getVal(grid, ref) < 0 ? `(${getVal(grid, ref)})` : getVal(grid, ref));
                safeFormula = safeFormula.replace(/([0-9.]+)%/g, "$1*0.01");
                if(safeFormula.includes("PRODUCT")) safeFormula = safeFormula.replace(/PRODUCT\(([^)]+)\)/g, (m, a) => { let r=1; a.split(',').forEach(p=>r*=parseFloat(new Function('return '+p)())); return r; });
                let resultValue = new Function('return ' + safeFormula)();
                const pTag = el.querySelector('p') || el;
                let displayVal = Math.round(resultValue);
                if(el.getAttribute('dze_format_separator') === ',') displayVal = displayVal.toLocaleString();
                pTag.innerText = displayVal;
                const addr = el.getAttribute('data-addr'); if(addr) grid[addr] = resultValue;
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

/* --- Global Resize Logic --- */
document.addEventListener('mousemove', (e) => {
    if (!isResizing || !resizeTarget) return;

    let newPixelVal = 0;
    if (resizeType === 'col') {
        const diff = e.clientX - resizeStartPos;
        newPixelVal = Math.max(10, resizeStartVal + diff);
    } else if (resizeType === 'row') {
        const diff = e.clientY - resizeStartPos;
        newPixelVal = Math.max(10, resizeStartVal + diff);
    }

    let finalVal = '';
    if (resizeUnit === '%') {
        const newPercent = (newPixelVal / resizeParentSize) * 100;
        finalVal = newPercent.toFixed(2) + '%';
    } else {
        finalVal = newPixelVal + 'px';
    }

    if (resizeType === 'col') {
        resizeTarget.style.width = finalVal;
        resizeTarget.setAttribute('width', finalVal); 
        if(targetColElement && targetColElement.parentNode) {
            targetColElement.style.width = finalVal;
            targetColElement.setAttribute('width', finalVal);
        }
    } else {
        resizeTarget.style.height = finalVal;
        resizeTarget.setAttribute('height', finalVal);
        if(resizeTarget.parentElement) resizeTarget.parentElement.style.height = finalVal;
    }
});

document.addEventListener('mouseup', () => {
    if (isResizing) {
        isResizing = false; resizeTarget = null; resizeType = null; targetColElement = null;
        document.body.classList.remove('is-resizing', 'is-resizing-col', 'is-resizing-row');
        saveHistory();
    }
    isDragging = false;
});

function attachEvents(container) {
    const tables = container.querySelectorAll('table');
    tables.forEach(table => {
        initCoordinates(table); updateTooltips();

        table.addEventListener('mousemove', function(e) {
            if (isResizing) return;
            const td = e.target.closest('td');
            if (!td) return;
            const rect = td.getBoundingClientRect();
            const rightDist = Math.abs(rect.right - e.clientX);
            const bottomDist = Math.abs(rect.bottom - e.clientY);
            const sensitive = 5;

            td.removeAttribute('data-resize-col'); td.removeAttribute('data-resize-row'); table.style.cursor = '';

            if (rightDist <= sensitive) {
                td.setAttribute('data-resize-col', 'true'); table.style.cursor = 'col-resize';
            } else if (bottomDist <= sensitive) {
                td.setAttribute('data-resize-row', 'true'); table.style.cursor = 'row-resize';
            }
        });

        table.addEventListener('mousedown', function(e) {
            const td = e.target.closest('td');
            if (!td) return;

            if (td.getAttribute('data-resize-col') === 'true') {
                isResizing = true; resizeType = 'col'; resizeTarget = td;
                resizeStartVal = td.getBoundingClientRect().width; 
                resizeStartPos = e.clientX;
                targetColElement = null;
                const colIdx = parseInt(td.getAttribute('data-col'));
                const colgroup = table.querySelector('colgroup');
                if(colgroup) {
                    const cols = colgroup.querySelectorAll('col');
                    if(cols[colIdx]) targetColElement = cols[colIdx];
                }
                if (td.style.width && td.style.width.includes('%')) {
                    resizeUnit = '%'; resizeParentSize = table.getBoundingClientRect().width;
                } else { resizeUnit = 'px'; }
                document.body.classList.add('is-resizing', 'is-resizing-col'); e.preventDefault(); return;
            }

            if (td.getAttribute('data-resize-row') === 'true') {
                isResizing = true; resizeType = 'row'; resizeTarget = td;
                resizeStartVal = td.getBoundingClientRect().height;
                resizeStartPos = e.clientY;
                if (td.style.height && td.style.height.includes('%')) {
                    resizeUnit = '%'; resizeParentSize = table.getBoundingClientRect().height;
                } else { resizeUnit = 'px'; }
                document.body.classList.add('is-resizing', 'is-resizing-row'); e.preventDefault(); return;
            }

            container.querySelectorAll('.selected-cell').forEach(c => c.classList.remove('selected-cell'));
            const editing = container.querySelector('.editing-cell');
            if (editing && editing !== td) finishEdit(editing);
            isDragging = true; startCell = td; currentTable = table;
            selectCell(td); 
        });

        table.addEventListener('mouseover', function(e) {
            if (isResizing || !isDragging) return;
            const td = e.target.closest('td'); if (!td || table !== currentTable) return;
            selectRange(currentTable, startCell, td);
        });

        table.addEventListener('dblclick', function(e) {
            const td = e.target.closest('td'); if (!td) return;
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
                selectedCells.forEach(td => { if(!td.hasAttribute('data-dze-formula')) { const p=td.querySelector('p'); if(p) p.innerText=""; else td.innerText=""; } });
                calculateAll(); saveHistory(); 
            }
        }
    });
}

// --- Helper Functions ---
function initCoordinates(table) {
    const matrix = []; const rows = table.rows;
    for(let r=0; r<rows.length; r++) {
        if(!matrix[r]) matrix[r] = []; const cells = rows[r].cells; let colIdx = 0;
        for(let c=0; c<cells.length; c++) {
            while(matrix[r][colIdx]) colIdx++;
            const cell = cells[c]; cell.setAttribute('data-row', r); cell.setAttribute('data-col', colIdx);
            const rs = cell.rowSpan || 1, cs = cell.colSpan || 1;
            for(let i=0; i<rs; i++) for(let j=0; j<cs; j++) { if(!matrix[r+i]) matrix[r+i]=[]; matrix[r+i][colIdx+j]=true; } colIdx += cs;
        }
    }
}
function mapTableToGrid(table) {
    const grid = {}; table.querySelectorAll('td').forEach(td => {
        const r = parseInt(td.getAttribute('data-row')), c = parseInt(td.getAttribute('data-col'));
        const addr = `${getColLetter(c)}${r + 1}`;
        let text = td.innerText.replace(/,/g, '').trim(); let val = text===""?0:parseFloat(text);
        grid[addr] = isNaN(val) ? 0 : val; td.setAttribute('data-addr', addr);
    }); return grid;
}
function getColLetter(i) { let l=''; while(i>=0){ l=String.fromCharCode((i%26)+65)+l; i=Math.floor(i/26)-1; } return l; }
function getVal(g, r) { return g[r] !== undefined ? g[r] : 0; }
function getRangeSum(g, r) {
    const [s, e] = r.split(':'); const sc = s.match(/[A-Z]+/)[0], ec = e.match(/[A-Z]+/)[0];
    const sr = parseInt(s.match(/\d+/)[0]), er = parseInt(e.match(/\d+/)[0]);
    const sci = colToNum(sc), eci = colToNum(ec); let sum = 0;
    for(let i=sr; i<=er; i++) for(let j=sci; j<=eci; j++) sum += getVal(g, `${getColLetter(j)}${i}`); return sum;
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
    calculateAll(); saveHistory(); 
}

/* --- Export Helpers --- */
function getCleanHTML() {
    const clone = previewDiv.cloneNode(true);
    const dirtyClasses = ['selected-cell', 'editing-cell', 'drag-over'];
    dirtyClasses.forEach(c => clone.querySelectorAll('.' + c).forEach(el => el.classList.remove(c)));
    clone.querySelectorAll('*[class=""]').forEach(el => el.removeAttribute('class'));
    const dirtyAttrs = ['data-row', 'data-col', 'data-addr', 'data-tooltip', 'data-org-title', 'contenteditable', 'spellcheck', 'data-resize-col', 'data-resize-row'];
    clone.querySelectorAll('*').forEach(el => {
        dirtyAttrs.forEach(a => el.removeAttribute(a));
        if(el.hasAttribute('data-org-title')) { el.setAttribute('title', el.getAttribute('data-org-title')); el.removeAttribute('data-org-title'); }
        if(el.getAttribute('style') === '') el.removeAttribute('style');
    });
    return clone.innerHTML.replace(/></g, '>\n<').trim();
}

function copyPreviewHTML() {
    const html = getCleanHTML();
    if(!html) { alert("내용이 없습니다."); return; }
    navigator.clipboard.writeText(html).then(() => alert("깨끗한 HTML이 복사되었습니다.")).catch(()=>alert("실패"));
}

function exportToCSharp() {
    const html = getCleanHTML();
    if(!html) { alert("내용이 없습니다."); return; }
    const lines = html.split('\n');
    let code = 'StringBuilder sb = new StringBuilder();\n';
    lines.forEach(l => {
        if(l.trim()==='') return;
        code += `sb.AppendLine($@"${l.replace(/"/g, '""').replace(/\r/g, '')}");\n`;
    });
    navigator.clipboard.writeText(code).then(() => alert("C# 코드 복사 완료.")).catch(()=>alert("실패"));
}

// [최종 수정] 딜레이 삭제 + 취소 시 자동 닫기 기능 추가
function printPreview() {
    if(!previewDiv.innerHTML.trim()) { alert("인쇄할 내용이 없습니다."); return; }

    // 팝업 생성
    const printWindow = window.open('', '_blank', 'width=1000,height=800');
    if (!printWindow) { alert("팝업 차단을 해제해주세요."); return; }

    const content = previewDiv.innerHTML;

    printWindow.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>인쇄 미리보기</title>
            <style>
                body { font-family: 'Pretendard', sans-serif; margin: 0; padding: 0; }
                
                /* 테이블 스타일 복구 */
                table { border-collapse: collapse; width: 100% !important; margin: 0 auto; }
                td, th { border: 1px solid #dee2e6; padding: 4px; font-size: 12px; }

                /* [중요] 배경색 강제 출력 */
                * {
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                }
                
                /* 인쇄용 페이지 설정 */
                @page { size: A4; margin: 10mm; }
                body { width: 100%; }
            </style>
        </head>
        <body>
            <div class="print-content">
                ${content}
            </div>
            <script>
                window.onload = function() {
                    window.focus();
                    // 딜레이 없이 바로 인쇄 실행
                    window.print();
                };

                // [핵심] 인쇄 대화상자가 닫히면(인쇄했든 취소했든) 팝업창도 닫아라
                window.onafterprint = function() {
                    window.close();
                };
            <\/script>
        </body>
        </html>
    `);
    
    printWindow.document.close();
}


