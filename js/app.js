/**
 * Main Application Logic
 */

const APP_STATE = {
    pdfHandler: new PDFHandler(),
    ocrEngine: new OCREngine(),
    grader: new Grader(),

    // Config
    config: {
        threshold: 211,
        sensitivity: 0.2,
        studentIdRegion: { x: 50, y: 100, w: 200, h: 500 },
        studentIdGrid: { rows: 10, cols: 7 },

        // 4 separate answer blocks
        questionsPerBlock: 25,
        answerBlocks: [
            { x: 300, y: 100, w: 200, h: 500 }, // Block 1
            { x: 550, y: 100, w: 200, h: 500 }, // Block 2
            { x: 800, y: 100, w: 200, h: 500 }, // Block 3
            { x: 1050, y: 100, w: 200, h: 500 }  // Block 4
        ]
    },

    // Runtime
    isSelecting: false, // 'id', 'block1', 'block2', 'block3', 'block4'
    selectionStart: null,
    currentPage: 1,
    totalPages: 0,
    answerKey: []
};

document.addEventListener('DOMContentLoaded', () => {
    initUI();
    initCanvasEvents();
});

function initUI() {
    // Set initial UI values from config
    document.getElementById('thresholdSlider').value = APP_STATE.config.threshold;
    document.getElementById('thresholdValue').textContent = APP_STATE.config.threshold;
    document.getElementById('sensitivitySlider').value = APP_STATE.config.sensitivity;
    document.getElementById('sensitivityValue').textContent = Math.round(APP_STATE.config.sensitivity * 100) + '%';

    // File Upload
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
    dropZone.addEventListener('drop', handleFileDrop);
    fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));

    // Sliders
    document.getElementById('thresholdSlider').addEventListener('input', (e) => {
        APP_STATE.config.threshold = parseInt(e.target.value);
        document.getElementById('thresholdValue').textContent = APP_STATE.config.threshold;
        requestRender();
    });

    document.getElementById('sensitivitySlider').addEventListener('input', (e) => {
        APP_STATE.config.sensitivity = parseFloat(e.target.value);
        document.getElementById('sensitivityValue').textContent = Math.round(APP_STATE.config.sensitivity * 100) + '%';
        // Re-analyze view only?
    });

    // Inputs
    document.getElementById('idDigits').addEventListener('change', updateGridConfig);
    document.getElementById('idOptions').addEventListener('change', updateGridConfig);
    document.getElementById('numBlocks').addEventListener('change', updateBlockVisibility);
    document.getElementById('questionsPerBlock').addEventListener('change', updateQuestionsPerBlock);

    // Buttons
    document.getElementById('setStudentIdAreaBtn').addEventListener('click', () => startSelection('id'));
    document.getElementById('resetStudentIdAreaBtn').addEventListener('click', resetStudentIdArea);

    document.getElementById('setBlock1Btn').addEventListener('click', () => startSelection('block1'));
    document.getElementById('setBlock2Btn').addEventListener('click', () => startSelection('block2'));
    document.getElementById('setBlock3Btn').addEventListener('click', () => startSelection('block3'));
    document.getElementById('setBlock4Btn').addEventListener('click', () => startSelection('block4'));
    document.getElementById('resetAllBlocksBtn').addEventListener('click', resetAllBlocks);

    document.getElementById('analyzeFirstPageBtn').addEventListener('click', analyzeFirstPage);
    document.getElementById('startGradingBtn').addEventListener('click', startGradingFlow);
    document.getElementById('exportExcelBtn').addEventListener('click', exportExcel);

    // Debug
    document.getElementById('copyConfigBtn').addEventListener('click', copyConfigToClipboard);

    // Nav
    document.getElementById('prevPage').addEventListener('click', () => changePage(-1));
    document.getElementById('nextPage').addEventListener('click', () => changePage(1));

    // View Options
    document.getElementById('showBinarized').addEventListener('change', requestRender);
    document.getElementById('showRegions').addEventListener('change', drawOverlay);

    // Window Resize
    window.addEventListener('resize', () => {
        // Debounce resize
        setTimeout(requestRender, 200);
    });
}

async function handleFileDrop(e) {
    e.preventDefault();
    document.getElementById('dropZone').classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) {
        handleFile(e.dataTransfer.files[0]);
    }
}

async function handleFile(file) {
    if (file.type !== 'application/pdf') {
        alert('PDFファイルを選択してください');
        return;
    }

    document.getElementById('fileName').textContent = file.name;
    document.getElementById('fileInfo').style.display = 'flex';
    document.getElementById('dropZone').style.display = 'none'; // Compact view

    document.getElementById('loader').style.display = 'block';

    try {
        APP_STATE.totalPages = await APP_STATE.pdfHandler.load(file);
        document.getElementById('pageCount').textContent = `${APP_STATE.totalPages} ページ`;
        document.getElementById('analyzeFirstPageBtn').disabled = false;

        APP_STATE.currentPage = 1;
        updateNav();
        await requestRender();
    } catch (err) {
        alert('PDF読み込みエラー: ' + err.message);
    } finally {
        document.getElementById('loader').style.display = 'none';
    }
}

async function changePage(delta) {
    const newPage = APP_STATE.currentPage + delta;
    if (newPage >= 1 && newPage <= APP_STATE.totalPages) {
        APP_STATE.currentPage = newPage;
        updateNav();
        await requestRender();
        // If we already have results for this page, show them
        showResultsForCurrentPage();
    }
}

function updateNav() {
    document.getElementById('pageIndicator').textContent = `Page ${APP_STATE.currentPage} / ${APP_STATE.totalPages}`;
    document.getElementById('prevPage').disabled = APP_STATE.currentPage <= 1;
    document.getElementById('nextPage').disabled = APP_STATE.currentPage >= APP_STATE.totalPages;
}

// Canvas & Rendering
async function requestRender() {
    if (!APP_STATE.totalPages) return;

    // Render underlying PDF
    const viewport = await APP_STATE.pdfHandler.renderPage(APP_STATE.currentPage);

    // Binarize preview if checked
    const showBin = document.getElementById('showBinarized').checked;
    if (showBin) {
        const imageData = APP_STATE.pdfHandler.getImageData();
        const binarized = APP_STATE.ocrEngine.binarize(imageData, APP_STATE.config.threshold);
        APP_STATE.pdfHandler.ctx.putImageData(binarized, 0, 0);
    }

    drawOverlay();
}

// Canvas Interaction (Region Selection)
function initCanvasEvents() {
    const canvas = document.getElementById('overlayCanvas');

    canvas.addEventListener('mousedown', (e) => {
        if (!APP_STATE.isSelecting) return;

        const rect = canvas.getBoundingClientRect();
        const scaleX = canvas.width / rect.width;
        const scaleY = canvas.height / rect.height;

        APP_STATE.selectionStart = {
            x: (e.clientX - rect.left) * scaleX,
            y: (e.clientY - rect.top) * scaleY
        };
    });

    canvas.addEventListener('mousemove', (e) => {
        if (!APP_STATE.isSelecting || !APP_STATE.selectionStart) return;

        const rect = canvas.getBoundingClientRect();
        const scaleX = canvas.width / rect.width;
        const scaleY = canvas.height / rect.height;
        const currentX = (e.clientX - rect.left) * scaleX;
        const currentY = (e.clientY - rect.top) * scaleY;

        drawOverlay(); // Clear and redraw existing

        // Draw drag rect
        const ctx = canvas.getContext('2d');
        ctx.strokeStyle = '#2563eb';
        ctx.lineWidth = 2;
        ctx.setLineDash([5, 5]);
        ctx.strokeRect(
            APP_STATE.selectionStart.x,
            APP_STATE.selectionStart.y,
            currentX - APP_STATE.selectionStart.x,
            currentY - APP_STATE.selectionStart.y
        );
        ctx.setLineDash([]);
    });

    canvas.addEventListener('mouseup', (e) => {
        if (!APP_STATE.isSelecting || !APP_STATE.selectionStart) return;

        const rect = canvas.getBoundingClientRect();
        const scaleX = canvas.width / rect.width;
        const scaleY = canvas.height / rect.height;
        const currentX = (e.clientX - rect.left) * scaleX;
        const currentY = (e.clientY - rect.top) * scaleY;

        // Normalize rect
        const x = Math.min(APP_STATE.selectionStart.x, currentX);
        const y = Math.min(APP_STATE.selectionStart.y, currentY);
        const w = Math.abs(currentX - APP_STATE.selectionStart.x);
        const h = Math.abs(currentY - APP_STATE.selectionStart.y);

        // Update config
        if (APP_STATE.isSelecting === 'id') {
            APP_STATE.config.studentIdRegion = { x, y, w, h };
        } else if (APP_STATE.isSelecting === 'block1') {
            APP_STATE.config.answerBlocks[0] = { x, y, w, h };
            updateBlockStatus(1, true);
        } else if (APP_STATE.isSelecting === 'block2') {
            APP_STATE.config.answerBlocks[1] = { x, y, w, h };
            updateBlockStatus(2, true);
        } else if (APP_STATE.isSelecting === 'block3') {
            APP_STATE.config.answerBlocks[2] = { x, y, w, h };
            updateBlockStatus(3, true);
        } else if (APP_STATE.isSelecting === 'block4') {
            APP_STATE.config.answerBlocks[3] = { x, y, w, h };
            updateBlockStatus(4, true);
        }

        APP_STATE.isSelecting = false;
        APP_STATE.selectionStart = null;
        document.body.style.cursor = 'default';

        // Reset button states
        document.getElementById('setStudentIdAreaBtn').classList.remove('btn-active'); // Add active style handling if needed

        drawOverlay();
    });
}

function startSelection(type) {
    APP_STATE.isSelecting = type;
    document.body.style.cursor = 'crosshair';
}

function updateGridConfig() {
    // ID Config - directly from inputs
    const idCols = parseInt(document.getElementById('idDigits').value) || 7;
    const idRows = parseInt(document.getElementById('idOptions').value) || 10;
    APP_STATE.config.studentIdGrid = { rows: idRows, cols: idCols };

    drawOverlay();
}

function updateQuestionsPerBlock() {
    const qPerBlock = parseInt(document.getElementById('questionsPerBlock').value) || 25;
    APP_STATE.config.questionsPerBlock = qPerBlock;

    // Update block labels
    for (let i = 0; i < 4; i++) {
        const start = i * qPerBlock + 1;
        const end = (i + 1) * qPerBlock;
        const blockNum = i + 1;
        const h4 = document.querySelector(`#setBlock${blockNum}Btn`).previousElementSibling;
        if (h4 && h4.tagName === 'H4') {
            h4.textContent = `ブロック${blockNum}（問${start}-${end}）`;
        }
    }
}

function updateBlockVisibility() {
    const numBlocks = parseInt(document.getElementById('numBlocks').value) || 4;

    // Show/hide block config divs
    for (let i = 1; i <= 4; i++) {
        const blockDiv = document.querySelector(`#setBlock${i}Btn`).closest('.block-config');
        if (blockDiv) {
            blockDiv.style.display = i <= numBlocks ? 'block' : 'none';
        }
    }

    drawOverlay();
}

function updateBlockStatus(blockNum, isSet) {
    const statusEl = document.getElementById(`block${blockNum}Status`);
    if (statusEl) {
        statusEl.textContent = isSet ? '設定済み' : '未設定';
        if (isSet) {
            statusEl.classList.add('set');
        } else {
            statusEl.classList.remove('set');
        }
    }
}

function resetStudentIdArea() {
    APP_STATE.config.studentIdRegion = { x: 50, y: 100, w: 200, h: 500 };
    drawOverlay();
    alert('学籍番号エリアをリセットしました。「エリアを指定」ボタンで再設定してください。');
}

function resetAllBlocks() {
    APP_STATE.config.answerBlocks = [
        { x: 300, y: 100, w: 200, h: 500 },
        { x: 550, y: 100, w: 200, h: 500 },
        { x: 800, y: 100, w: 200, h: 500 },
        { x: 1050, y: 100, w: 200, h: 500 }
    ];
    for (let i = 1; i <= 4; i++) {
        updateBlockStatus(i, false);
    }
    drawOverlay();
    alert('全ブロックをリセットしました。各ブロックを再設定してください。');
}

function drawOverlay() {
    if (!document.getElementById('showRegions').checked) {
        const canvas = document.getElementById('overlayCanvas');
        const ctx = canvas.getContext('2d');
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        return;
    }

    const canvas = document.getElementById('overlayCanvas');
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Draw ID Region
    const idR = APP_STATE.config.studentIdRegion;
    if (idR) {
        ctx.strokeStyle = '#ef4444'; // Red
        ctx.lineWidth = 3;
        ctx.strokeRect(idR.x, idR.y, idR.w, idR.h);

        // Draw Detailed Grid
        drawGrid(ctx, idR, APP_STATE.config.studentIdGrid, '#ef4444');
    }

    // Draw 4 Answer Blocks (only active ones)
    const numBlocks = parseInt(document.getElementById('numBlocks')?.value || 4);
    const colors = ['#10b981', '#3b82f6', '#f59e0b', '#8b5cf6']; // Green, Blue, Orange, Purple

    for (let idx = 0; idx < numBlocks; idx++) {
        const block = APP_STATE.config.answerBlocks[idx];
        if (block) {
            ctx.strokeStyle = colors[idx];
            ctx.lineWidth = 3;
            ctx.strokeRect(block.x, block.y, block.w, block.h);

            // Draw grid for this block
            const grid = { rows: APP_STATE.config.questionsPerBlock, cols: 10 };
            drawGrid(ctx, block, grid, colors[idx]);

            // Draw block label
            ctx.fillStyle = colors[idx];
            ctx.font = 'bold 14px sans-serif';
            ctx.fillText(`Block ${idx + 1}`, block.x + 5, block.y - 5);
        }
    }

    // Update config output for debugging
    if (typeof updateConfigOutput === 'function') {
        updateConfigOutput();
    }
}

function drawGrid(ctx, rect, grid, color) {
    if (!rect || !grid) return;

    const cellW = rect.w / grid.cols;
    const cellH = rect.h / grid.rows;

    ctx.save();
    ctx.strokeStyle = color;
    ctx.lineWidth = 1;
    ctx.globalAlpha = 0.5;

    for (let r = 0; r < grid.rows; r++) {
        for (let c = 0; c < grid.cols; c++) {
            const x = rect.x + c * cellW;
            const y = rect.y + r * cellH;

            // Draw cell frame
            ctx.strokeRect(x, y, cellW, cellH);
        }
    }

    ctx.restore();
}

// Logic: Analysis & Grading

async function analyzeFirstPage() {
    // Force set page 1
    if (APP_STATE.currentPage !== 1) {
        APP_STATE.currentPage = 1;
        updateNav();
        await requestRender();
    }

    const imageData = APP_STATE.pdfHandler.getImageData();
    const binImage = APP_STATE.ocrEngine.binarize(imageData, APP_STATE.config.threshold);

    // 1. Detect ID
    const idRes = APP_STATE.ocrEngine.detectMarks(binImage, APP_STATE.config.studentIdRegion, APP_STATE.config.studentIdGrid, APP_STATE.config.sensitivity);
    const studentId = APP_STATE.ocrEngine.readVerticalID(idRes.matrix);

    // 2. Detect Answers from active blocks only
    const numBlocks = parseInt(document.getElementById('numBlocks').value) || 4;
    const allAnswers = [];
    const allDebugData = [];

    for (let i = 0; i < numBlocks; i++) {
        const block = APP_STATE.config.answerBlocks[i];
        const grid = { rows: APP_STATE.config.questionsPerBlock, cols: 10 };

        const blockRes = APP_STATE.ocrEngine.detectMarks(binImage, block, grid, APP_STATE.config.sensitivity);
        const blockAnswers = APP_STATE.ocrEngine.readHorizontalAnswers(blockRes.matrix);

        allAnswers.push(...blockAnswers);
        allDebugData.push(blockRes.debug);
    }

    APP_STATE.answerKey = allAnswers;
    APP_STATE.grader.setAnswerKey(allAnswers);

    // Draw detection results on overlay for debugging
    drawDetectionResultsBlocks(idRes, allDebugData);

    // Show Results with actual answer preview
    displayResultsWithAnswers(1, studentId, allAnswers);

    // Enable Grading
    document.getElementById('startGradingBtn').style.display = 'inline-block';

    alert(`1枚目を解析しました。\n学籍番号: ${studentId}\n検出回答数: ${allAnswers.filter(a => a !== null).length}\n\nプレビューを確認し、問題なければ「全ページ採点開始」を押してください。`);
}

function drawDetectionResults(idRes, ansRes) {
    const canvas = document.getElementById('overlayCanvas');
    const ctx = canvas.getContext('2d');

    // Draw marked cells with color coding
    ctx.save();

    // Draw ID marks
    idRes.debug.forEach(cell => {
        if (cell.isMarked) {
            ctx.fillStyle = 'rgba(239, 68, 68, 0.4)'; // Red
            ctx.fillRect(cell.x, cell.y, cell.w, cell.h);

            // Draw ratio text
            ctx.fillStyle = '#fff';
            ctx.font = '10px monospace';
            ctx.fillText(Math.round(cell.ratio * 100) + '%', cell.x + 2, cell.y + 12);
        }
    });

    // Draw answer marks
    ansRes.debug.forEach(cell => {
        if (cell.isMarked) {
            ctx.fillStyle = 'rgba(16, 185, 129, 0.4)'; // Green
            ctx.fillRect(cell.x, cell.y, cell.w, cell.h);
        }
    });

    ctx.restore();
}

function drawDetectionResultsBlocks(idRes, allDebugData) {
    const canvas = document.getElementById('overlayCanvas');
    const ctx = canvas.getContext('2d');

    ctx.save();

    // Draw ID marks
    idRes.debug.forEach(cell => {
        if (cell.isMarked) {
            ctx.fillStyle = 'rgba(239, 68, 68, 0.4)'; // Red
            ctx.fillRect(cell.x, cell.y, cell.w, cell.h);

            ctx.fillStyle = '#fff';
            ctx.font = '10px monospace';
            ctx.fillText(Math.round(cell.ratio * 100) + '%', cell.x + 2, cell.y + 12);
        }
    });

    // Draw answer marks from all blocks
    const colors = [
        'rgba(16, 185, 129, 0.4)',  // Green
        'rgba(59, 130, 246, 0.4)',  // Blue
        'rgba(245, 158, 11, 0.4)',  // Orange
        'rgba(139, 92, 246, 0.4)'   // Purple
    ];

    allDebugData.forEach((blockDebug, blockIdx) => {
        blockDebug.forEach(cell => {
            if (cell.isMarked) {
                ctx.fillStyle = colors[blockIdx];
                ctx.fillRect(cell.x, cell.y, cell.w, cell.h);
            }
        });
    });

    ctx.restore();
}


function flattenAnswers(matrix, rowsPerBlock, optsPerBlock, numBlocks) {
    // matrix is [TotalRows][TotalCols]
    // If Grid was configured as single big grid:
    // Rows = rowsPerBlock (e.g. 25)
    // Cols = numBlocks * optsPerBlock (e.g. 40)

    const answers = [];

    // Iterate by Block then by Row
    for (let b = 0; b < numBlocks; b++) {
        for (let r = 0; r < rowsPerBlock; r++) {
            // Get row from matrix
            if (!matrix[r]) continue;

            // Slice the columns for this block
            const colStart = b * optsPerBlock;
            const colEnd = colStart + optsPerBlock;

            // Check finding marks in this slice
            let markedVal = null;
            let markCount = 0;

            for (let c = colStart; c < colEnd; c++) {
                if (matrix[r][c] && matrix[r][c].isMarked) {
                    markedVal = c - colStart; // Value is relative to block start (0-9)
                    markCount++;
                }
            }

            if (markCount > 1) answers.push("MULTIPLE");
            else answers.push(markedVal);
        }
    }
    return answers;
}

async function startGradingFlow() {
    document.getElementById('startGradingBtn').disabled = true;
    document.getElementById('loader').style.display = 'block';

    APP_STATE.grader.reset();

    // Loop through all pages starting from 2
    for (let p = 2; p <= APP_STATE.totalPages; p++) {
        // Render to canvas (hidden is fine, but we use the main one)
        // Note: Render forces the canvas to update, which is slow but necessary for pixel data
        await APP_STATE.pdfHandler.renderPage(p);

        const imageData = APP_STATE.pdfHandler.getImageData();
        const binImage = APP_STATE.ocrEngine.binarize(imageData, APP_STATE.config.threshold);

        // ID
        const idRes = APP_STATE.ocrEngine.detectMarks(binImage, APP_STATE.config.studentIdRegion, APP_STATE.config.studentIdGrid, APP_STATE.config.sensitivity);
        const studentId = APP_STATE.ocrEngine.readVerticalID(idRes.matrix);

        // Answers - process active blocks only
        const numBlocks = parseInt(document.getElementById('numBlocks').value) || 4;
        const studentAns = [];
        for (let i = 0; i < numBlocks; i++) {
            const block = APP_STATE.config.answerBlocks[i];
            const grid = { rows: APP_STATE.config.questionsPerBlock, cols: 10 };

            const blockRes = APP_STATE.ocrEngine.detectMarks(binImage, block, grid, APP_STATE.config.sensitivity);
            const blockAnswers = APP_STATE.ocrEngine.readHorizontalAnswers(blockRes.matrix);

            studentAns.push(...blockAnswers);
        }

        // Grade
        APP_STATE.grader.gradeStudent(studentId, studentAns, p);

        // Optional: Update progress UI
        document.getElementById('pageIndicator').textContent = `Processing ${p} / ${APP_STATE.totalPages}`;
    }

    document.getElementById('loader').style.display = 'none';
    document.getElementById('startGradingBtn').disabled = false;
    document.getElementById('exportCsvBtn').style.display = 'inline-block';

    alert('全ページの採点が完了しました。CSVエクスポートが可能です。');

    // Go to first student page
    changePage(2 - APP_STATE.currentPage);
}

function displayResults(page, studentId, markCount, scoreMsg) {
    document.getElementById('resultsPreview').style.display = 'block';
    document.getElementById('previewPageNum').textContent = `Page ${page}`;
    document.getElementById('resStudentId').textContent = studentId;
    document.getElementById('resMarkCount').textContent = markCount;
    document.getElementById('resScore').textContent = scoreMsg;
}

function displayResultsWithAnswers(page, studentId, answers) {
    document.getElementById('resultsPreview').style.display = 'block';
    document.getElementById('previewPageNum').textContent = `Page ${page}`;
    document.getElementById('resStudentId').textContent = studentId;
    document.getElementById('resMarkCount').textContent = answers.filter(a => a !== null).length;
    document.getElementById('resScore').textContent = page === 1 ? '正答キー' : '-';

    // Show answer values
    const grid = document.getElementById('resAnswersGrid');
    grid.innerHTML = '';
    answers.forEach((ans, idx) => {
        const div = document.createElement('div');
        div.className = 'ans-item';
        div.textContent = `${idx + 1}:${ans ?? '-'}`;
        grid.appendChild(div);
    });
}


function showResultsForCurrentPage() {
    // Check if we have processed results
    const res = APP_STATE.grader.results.find(r => r.page === APP_STATE.currentPage);
    if (res) {
        displayResults(res.page, res.studentId, res.details.length, `${res.score} / ${res.maxScore}`);

        // Render grid visualization
        const grid = document.getElementById('resAnswersGrid');
        grid.innerHTML = '';
        res.details.forEach(d => {
            const div = document.createElement('div');
            div.className = `ans-item ${d.isCorrect ? 'correct' : 'incorrect'}`;
            div.textContent = `${d.question}:${d.student ?? 'X'}`;
            grid.appendChild(div);
        });
    } else {
        document.getElementById('resultsPreview').style.display = 'none';
    }
}

function exportExcel() {
    const data = APP_STATE.grader.getExcelData();
    if (!data) {
        alert('エクスポートするデータがありません。');
        return;
    }

    const wb = XLSX.utils.book_new();
    const ws_data = [];

    // 1. Headers
    ws_data.push(data.headers);

    // 2. Points (Row 2)
    ws_data.push(data.pointsRow);

    // Create Base Sheet
    const ws = XLSX.utils.aoa_to_sheet(ws_data);

    // 3. Add Student Rows with Formulas
    const numQ = data.headers.length - 2; // Exclude ID and Total
    // Points Range: $B$2:$[LastCol]$2
    // SheetJS uses 0-indexed columns. Q1 is Col 1 ('B'). Last Q is Col numQ.
    const startColStr = XLSX.utils.encode_col(1); // 'B'
    const endColStr = XLSX.utils.encode_col(numQ);
    const pointsRange = `$${startColStr}$2:$${endColStr}$2`;

    data.dataRows.forEach((row, idx) => {
        const rowNum = idx + 3; // Excel Row Number (1-based)

        // Note: row array contains [ID, 1, 0, ..., null/dummy]
        // We write it to the sheet
        XLSX.utils.sheet_add_aoa(ws, [row], { origin: -1 });

        // Overwrite the last cell (Total Score) with Formula
        // Total Score Column is index numQ + 1
        const totalCellAddr = XLSX.utils.encode_cell({ r: rowNum - 1, c: numQ + 1 });

        // Student Answers Range: B[Row]:[LastCol][Row]
        const studentRange = `${startColStr}${rowNum}:${endColStr}${rowNum}`;
        const formula = `SUMPRODUCT(${pointsRange},${studentRange})`;

        ws[totalCellAddr] = { t: 'n', f: formula };
    });

    XLSX.utils.book_append_sheet(wb, ws, "Grading Results");
    XLSX.writeFile(wb, "grading_results.xlsx");
}

function updateConfigOutput() {
    const json = JSON.stringify(APP_STATE.config, null, 2);
    const textarea = document.getElementById('configOutput');
    if (textarea) textarea.value = json;
}

function copyConfigToClipboard() {
    const textarea = document.getElementById('configOutput');
    if (!textarea) return;
    textarea.select();
    document.execCommand('copy');
    alert('設定値をクリップボードにコピーしました。');
}

