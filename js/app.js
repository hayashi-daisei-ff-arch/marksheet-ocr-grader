/**
 * Main Application Logic
 */

const APP_STATE = {
    pdfHandler: new PDFHandler(),
    ocrEngine: new OCREngine(),
    grader: new Grader(),

    // Config
    config: {
        threshold: 200,
        sensitivity: 0.2,
        contrast: 0,
        studentIdRegion: { x: 100, y: 229, w: 177, h: 272 },
        studentIdGrid: { rows: 10, cols: 7 },

        // 4 separate answer blocks
        questionsPerBlock: 25,
        answerBlocks: [
            { x: 348.5, y: 174, w: 184, h: 677 }, // Block 1
            { x: 569.5, y: 176, w: 182, h: 675 }, // Block 2
            { x: 790.5, y: 177, w: 182, h: 674 }, // Block 3
            { x: 1011.5, y: 177, w: 180, h: 677 }  // Block 4
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
    document.getElementById('contrastSlider').value = APP_STATE.config.contrast;
    document.getElementById('contrastValue').textContent = APP_STATE.config.contrast;

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

    const sensitivitySlider = document.getElementById('sensitivitySlider');
    const sensitivityValue = document.getElementById('sensitivityValue');
    sensitivitySlider.addEventListener('input', (e) => {
        sensitivityValue.textContent = Math.round(e.target.value * 100) + '%';
        APP_STATE.config.sensitivity = parseFloat(e.target.value);
        if (document.getElementById('resultsPreview').style.display !== 'none') {
            requestRender();
        }
    });

    const contrastSlider = document.getElementById('contrastSlider');
    const contrastValue = document.getElementById('contrastValue');
    contrastSlider.addEventListener('input', (e) => {
        contrastValue.textContent = e.target.value;
        APP_STATE.config.contrast = parseInt(e.target.value);
    });
    contrastSlider.addEventListener('change', (e) => {
        requestRender(); // Render only when slider is released to avoid race condition
    });

    // Student List
    const studentListBtn = document.getElementById('studentListBtn');
    if (studentListBtn) {
        studentListBtn.addEventListener('click', toggleStudentList);
    }
    const closeModal = document.querySelector('.close-modal');
    if (closeModal) {
        closeModal.addEventListener('click', () => {
            document.getElementById('studentListModal').style.display = 'none';
        });
    }
    window.addEventListener('click', (e) => {
        const m = document.getElementById('studentListModal');
        if (e.target === m) m.style.display = 'none';
    });

    // Actions
    const reanalyzeBtn = document.getElementById('reanalyzePageBtn');
    if (reanalyzeBtn) reanalyzeBtn.addEventListener('click', reanalyzeCurrentPage);

    const exportPdfBtn = document.getElementById('exportPdfBtn');
    if (exportPdfBtn) exportPdfBtn.addEventListener('click', exportPdf);

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

    // Help
    document.getElementById('helpBtn').addEventListener('click', () => {
        window.open('https://github.com/hayashi-daisei-ff-arch/marksheet-ocr-grader/blob/main/README.md', '_blank');
    });

    // Results Preview Toggle
    const resPanel = document.getElementById('resultsPreview');
    resPanel.addEventListener('click', (e) => {
        // Only toggle if clicking the header H3 or its text
        if (e.target.tagName === 'H3' || e.target.closest('h3')) {
            resPanel.classList.toggle('collapsed');
        }
    });

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

        // Show overlay and results if available
        if (APP_STATE.grader.results.some(r => r.page === APP_STATE.currentPage)) {
            await visualizeCurrentPage();
        } else if (APP_STATE.currentPage === 1 && APP_STATE.answerKey.length > 0) {
            // Visualize Answer Key Page? 
            // Usually we don't have stored results for Page 1 in 'grader.results'.
            // But we can show the key in the panel.
            showResultsForCurrentPage();
        } else {
            // Clear overlay if no results
            const canvas = document.getElementById('overlayCanvas');
            const ctx = canvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            drawOverlay(); // Draw boxes setup
            document.getElementById('resultsPreview').style.display = 'none';
        }
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

    // Apply Contrast if needed
    const imageData = APP_STATE.pdfHandler.getImageData();
    if (imageData && APP_STATE.config.contrast !== 0) {
        APP_STATE.ocrEngine.applyContrast(imageData, APP_STATE.config.contrast);
        // Reflect contrast change to canvas immediately
        APP_STATE.pdfHandler.ctx.putImageData(imageData, 0, 0);
    }

    // Binarize preview if checked
    const showBin = document.getElementById('showBinarized').checked;
    if (showBin && imageData) {
        // Use the (possibly contrast-adjusted) imageData
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

function getBlockOptions(blockIndex) {
    const input = document.getElementById(`block${blockIndex + 1}Options`);
    if (input) {
        return parseInt(input.value) || 10;
    }
    return 10;
}

function resetStudentIdArea() {
    APP_STATE.config.studentIdRegion = { x: 100, y: 229, w: 177, h: 272 };
    drawOverlay();
    alert('学籍番号エリアをリセットしました。');
}

function resetAllBlocks() {
    APP_STATE.config.answerBlocks = [{}, {}, {}, {}];
    for (let i = 1; i <= 4; i++) {
        updateBlockStatus(i, false);
    }
    drawOverlay();
    alert('全ブロックをリセットしました。');
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
        const blockOptions = getBlockOptions(i);

        const blockRes = APP_STATE.ocrEngine.detectMarks(binImage, block, grid, APP_STATE.config.sensitivity);
        const blockAnswers = APP_STATE.ocrEngine.readHorizontalAnswers(blockRes.matrix, blockOptions);

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

    const numBlocks = parseInt(document.getElementById('numBlocks').value) || 4;
    const qPerBlock = APP_STATE.config.questionsPerBlock;

    // First, draw grey marks for invalid options
    allDebugData.forEach((blockDebug, blockIdx) => {
        const block = APP_STATE.config.answerBlocks[blockIdx];
        if (!block) return;

        const maxOptions = getBlockOptions(blockIdx);
        const cellW = block.w / 10;

        blockDebug.forEach(cell => {
            if (cell.isMarked) {
                // Calculate which option this cell represents
                const colIdx = Math.floor((cell.x - block.x) / cellW);
                const optionValue = colIdx === 9 ? 0 : colIdx + 1;

                // Check if invalid
                const isInvalid = optionValue > maxOptions && optionValue !== 0;

                if (isInvalid) {
                    ctx.fillStyle = 'rgba(150, 150, 150, 0.4)'; // Grey
                    ctx.fillRect(cell.x, cell.y, cell.w, cell.h);
                }
            }
        });
    });

    // Then, draw answer marks from all blocks (valid options with colors)
    const colors = [
        'rgba(16, 185, 129, 0.4)',  // Green
        'rgba(59, 130, 246, 0.4)',  // Blue
        'rgba(245, 158, 11, 0.4)',  // Orange
        'rgba(139, 92, 246, 0.4)'   // Purple
    ];

    allDebugData.forEach((blockDebug, blockIdx) => {
        const block = APP_STATE.config.answerBlocks[blockIdx];
        if (!block) return;

        const maxOptions = getBlockOptions(blockIdx);
        const cellW = block.w / 10;

        blockDebug.forEach(cell => {
            if (cell.isMarked) {
                // Calculate which option this cell represents
                const colIdx = Math.floor((cell.x - block.x) / cellW);
                const optionValue = colIdx === 9 ? 0 : colIdx + 1;

                // Only draw colored overlay for valid options
                const isValid = optionValue <= maxOptions || optionValue === 0;

                if (isValid) {
                    ctx.fillStyle = colors[blockIdx];
                    ctx.fillRect(cell.x, cell.y, cell.w, cell.h);
                }
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

    // Loop through all pages starting from 2 (assuming page 1 is key)
    // Actually, usually Page 1 is also a student sheet if we keyed it separately, 
    // but here the flow implies Page 1 was the Answer Key source.
    // If the user wants to Grade Page 1 as well? 
    // The current flow uses Page 1 ONLY as Key. So we start from 2.

    // However, we want to visualize progress.
    for (let p = 2; p <= APP_STATE.totalPages; p++) {
        // Update current page state so renderPage works as expected contextually
        APP_STATE.currentPage = p;

        // Render to canvas
        await APP_STATE.pdfHandler.renderPage(p);

        const imageData = APP_STATE.pdfHandler.getImageData();
        const binImage = APP_STATE.ocrEngine.binarize(imageData, APP_STATE.config.threshold);

        // ID
        const idRes = APP_STATE.ocrEngine.detectMarks(binImage, APP_STATE.config.studentIdRegion, APP_STATE.config.studentIdGrid, APP_STATE.config.sensitivity);
        const studentId = APP_STATE.ocrEngine.readVerticalID(idRes.matrix);

        // Answers - process active blocks only
        const numBlocks = parseInt(document.getElementById('numBlocks').value) || 4;
        const studentAns = [];
        const allDebugData = []; // Store debug data for visualization

        for (let i = 0; i < numBlocks; i++) {
            const block = APP_STATE.config.answerBlocks[i];
            const grid = { rows: APP_STATE.config.questionsPerBlock, cols: 10 };
            const blockOptions = getBlockOptions(i);

            const blockRes = APP_STATE.ocrEngine.detectMarks(binImage, block, grid, APP_STATE.config.sensitivity);
            const blockAnswers = APP_STATE.ocrEngine.readHorizontalAnswers(blockRes.matrix, blockOptions);

            studentAns.push(...blockAnswers);
            allDebugData.push(blockRes.debug);
        }

        // Visualize!
        drawDetectionResultsBlocks(idRes, allDebugData);

        // Grade
        APP_STATE.grader.gradeStudent(studentId, studentAns, p);

        // Update progress UI
        document.getElementById('pageIndicator').textContent = `Processing ${p} / ${APP_STATE.totalPages}`;

        // Short pause to allow browser to render the canvas update
        await new Promise(resolve => setTimeout(resolve, 200));
    }

    document.getElementById('loader').style.display = 'none';
    document.getElementById('startGradingBtn').disabled = false;
    document.getElementById('exportExcelBtn').style.display = 'inline-block';
    document.getElementById('studentListBtn').disabled = false;
    document.getElementById('reanalyzePageBtn').style.display = 'inline-block';
    document.getElementById('exportPdfBtn').style.display = 'inline-block';

    alert('全ページの採点が完了しました。Excelエクスポートが可能です。');

    // Go back to the first student page (Page 2) and visualize
    if (APP_STATE.totalPages >= 2) {
        APP_STATE.currentPage = 2;
        updateNav();
        await requestRender(); // Wait for clean PDF render

        // Re-analyze just for overlay
        await visualizeCurrentPage();
    }
}

async function visualizeCurrentPage() {
    // Stabilization wait
    await new Promise(resolve => setTimeout(resolve, 100));

    const imageData = APP_STATE.pdfHandler.getImageData();
    if (!imageData) return;

    // 1. Draw ID (fresh detection for ID area overlay)
    // Contrast is already applied by requestRender if it was called before
    // If we just need to re-visualize without re-render, we assume canvas has correct state.

    // However, if we do NOT call requestRender before this, we might be using raw PDF render?
    // visualizeCurrentPage is often called after requestRender.

    const binImage = APP_STATE.ocrEngine.binarize(imageData, APP_STATE.config.threshold);
    const idRes = APP_STATE.ocrEngine.detectMarks(binImage, APP_STATE.config.studentIdRegion, APP_STATE.config.studentIdGrid, APP_STATE.config.sensitivity);

    // Clear & ID
    drawDetectionResultsBlocks(idRes, []);

    // 2. Re-detect answer blocks to show ALL marks (including invalid options)
    const numBlocks = parseInt(document.getElementById('numBlocks').value) || 4;
    const qPerBlock = APP_STATE.config.questionsPerBlock;
    const allBlockDetections = [];

    for (let i = 0; i < numBlocks; i++) {
        const blockConfig = APP_STATE.config.answerBlocks[i];
        if (blockConfig && blockConfig.x) {
            const grid = { rows: qPerBlock, cols: 10 };
            const blockRes = APP_STATE.ocrEngine.detectMarks(binImage, blockConfig, grid, APP_STATE.config.sensitivity);
            const maxOptions = getBlockOptions(i);
            allBlockDetections.push({ debug: blockRes.debug, maxOptions: maxOptions, blockIdx: i });
        }
    }

    // 3. Draw Answers from Grading Results with detection data
    drawGradingResultOverlay(allBlockDetections);

    // 4. Show Panel
    showResultsForCurrentPage();
}

function drawGradingResultOverlay(allBlockDetections = []) {
    const canvas = document.getElementById('overlayCanvas');
    const ctx = canvas.getContext('2d');

    const numBlocks = parseInt(document.getElementById('numBlocks').value) || 4;
    const qPerBlock = APP_STATE.config.questionsPerBlock;

    // First, ALWAYS draw all detected marks (including invalid ones in grey)
    allBlockDetections.forEach(({ debug, maxOptions, blockIdx }) => {
        const block = APP_STATE.config.answerBlocks[blockIdx];
        if (!block) return;

        const cellW = block.w / 10;
        const cellH = block.h / qPerBlock;

        debug.forEach(cell => {
            if (cell.isMarked) {
                // Calculate which option this cell represents (1-9,0)
                const colIdx = Math.floor((cell.x - block.x) / cellW);
                const optionValue = colIdx === 9 ? 0 : colIdx + 1;

                // Check if this option is beyond maxOptions
                const isInvalid = optionValue > maxOptions && optionValue !== 0;

                if (isInvalid) {
                    // Draw invalid options in grey
                    ctx.fillStyle = 'rgba(150, 150, 150, 0.4)'; // Grey
                    ctx.fillRect(cell.x + 2, cell.y + 2, cell.w - 4, cell.h - 4);
                }
            }
        });
    });

    // Then, if grading results exist, draw them on top
    const res = APP_STATE.grader.results.find(r => r.page === APP_STATE.currentPage);
    if (!res) return; // No grading results yet

    // Do not clear here, ID overlay and grey marks are already drawn

    res.details.forEach((d, globalIdx) => {
        const blockIdx = Math.floor(globalIdx / qPerBlock);
        if (blockIdx >= numBlocks) return;

        const inBlockIdx = globalIdx % qPerBlock;
        const block = APP_STATE.config.answerBlocks[blockIdx];
        if (!block) return;

        const cellW = block.w / 10;
        const cellH = block.h / qPerBlock;
        const y = block.y + inBlockIdx * cellH;

        // Color Logic per User Request:
        // Correct (Green), Incorrect (Red), No Answer/X (Yellow)
        let color = 'rgba(239, 68, 68, 0.4)'; // Default Red (Incorrect)

        if (d.student === null) {
            color = 'rgba(255, 193, 7, 0.4)'; // Yellow (No Answer)
        } else if (d.isCorrect) {
            color = 'rgba(16, 185, 129, 0.4)'; // Green (Correct)
        }

        if (d.student === null) {
            // Fill row for empty answer
            ctx.fillStyle = color;
            ctx.fillRect(block.x, y, block.w, cellH);
            return;
        }

        // Determine columns to draw (only valid answers)
        let indicesToDraw = [];
        if (typeof d.student === 'number') {
            indicesToDraw.push(d.student === 0 ? 9 : d.student - 1);
        } else if (typeof d.student === 'string') {
            // "1,2" -> parse
            d.student.split(',').forEach(v => {
                const n = parseInt(v); // 0-9
                if (!isNaN(n)) indicesToDraw.push(n === 0 ? 9 : n - 1);
            });
        }

        ctx.fillStyle = color;
        indicesToDraw.forEach(colIdx => {
            const x = block.x + colIdx * cellW;
            ctx.fillRect(x + 2, y + 2, cellW - 4, cellH - 4);
        });
    });
}


function displayResultsWithAnswers(page, studentId, answers, score, maxScore) {
    document.getElementById('resultsPreview').style.display = 'block';
    document.getElementById('previewPageNum').textContent = `Page ${page}`;

    // Student ID with edit capability
    const sidEl = document.getElementById('resStudentId');
    sidEl.textContent = studentId;

    if (page > 1) {
        sidEl.style.cursor = 'pointer';
        sidEl.style.textDecoration = 'underline';
        sidEl.title = 'クリックして学籍番号を修正';
        sidEl.onclick = () => editStudentId(page, studentId);
    } else {
        sidEl.style.cursor = 'default';
        sidEl.style.textDecoration = 'none';
        sidEl.onclick = null;
    }

    document.getElementById('resMarkCount').textContent = answers.filter(a => a !== null).length;

    // Score display: "Score / Max"
    if (score !== null && maxScore !== null && score !== undefined) {
        document.getElementById('resScore').textContent = `${score} / ${maxScore}`;
    } else {
        document.getElementById('resScore').textContent = page === 1 ? '正答キー' : '-';
    }

    // Show answer values
    const grid = document.getElementById('resAnswersGrid');
    grid.innerHTML = '';
    answers.forEach((ans, idx) => {
        const div = document.createElement('div');
        div.className = 'ans-item';

        let text = '-';
        if (ans !== null) {
            text = ans;
        } else {
            div.classList.add('empty');
        }

        div.textContent = `${idx + 1}:${text}`;
        div.title = "Click to edit";
        div.style.cursor = "pointer";
        div.onclick = () => editAnswer(page, idx, ans);

        grid.appendChild(div);
    });
}

function editStudentId(page, currentId) {
    const newId = prompt("学籍番号を修正:", currentId);
    if (newId === null) return;
    const trimmed = newId.trim();
    if (trimmed === currentId) return;

    const res = APP_STATE.grader.results.find(r => r.page === page);
    if (res) {
        res.studentId = trimmed;
        showResultsForCurrentPage();
    }
}

function editAnswer(page, qIdx, currentVal) {
    const defaultVal = currentVal === null ? '' : currentVal;
    const newValStr = prompt(`修正後の値を入力してください (Q${qIdx + 1})\n(0-9: 回答, '1,2': 複数, 空白: 未回答)`, defaultVal);
    if (newValStr === null) return; // Cancel

    let newVal = null;
    const trimmed = newValStr.trim();

    if (trimmed === '') {
        newVal = null;
    } else if (trimmed.includes(',')) {
        newVal = trimmed;
    } else {
        const n = parseInt(trimmed);
        if (!isNaN(n) && n >= 0 && n <= 9) {
            newVal = n;
        } else {
            if (trimmed.length > 0) newVal = trimmed;
        }
    }

    // Update Logic
    if (page === 1) {
        if (APP_STATE.answerKey[qIdx] !== undefined) APP_STATE.answerKey[qIdx] = newVal;
        if (APP_STATE.grader.correctAnswers && APP_STATE.grader.correctAnswers[qIdx] !== undefined) {
            APP_STATE.grader.correctAnswers[qIdx] = newVal;
        }
        showResultsForCurrentPage();
        alert('正答キーを更新しました。全ページ採点を行うと反映されます。');
        return;
    }

    const res = APP_STATE.grader.results.find(r => r.page === page);
    if (!res) return;

    const currentAnswers = res.details.map(d => d.student);
    currentAnswers[qIdx] = newVal;

    // Re-grade (overwrite)
    APP_STATE.grader.gradeStudent(res.studentId, currentAnswers, page);

    // Refresh UI
    showResultsForCurrentPage();

    // Refresh Overlay
    const canvas = document.getElementById('overlayCanvas');
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const imageData = APP_STATE.pdfHandler.getImageData();
    if (imageData) {
        const binImage = APP_STATE.ocrEngine.binarize(imageData, APP_STATE.config.threshold);
        const idRes = APP_STATE.ocrEngine.detectMarks(binImage, APP_STATE.config.studentIdRegion, APP_STATE.config.studentIdGrid, APP_STATE.config.sensitivity);
        drawDetectionResultsBlocks(idRes, []);
    }

    drawGradingResultOverlay();
}


function showResultsForCurrentPage() {
    // Check if we have processed results
    const res = APP_STATE.grader.results.find(r => r.page === APP_STATE.currentPage);
    if (res) {
        // Map details back to simple answer array for display
        const answers = res.details.map(d => d.student);
        displayResultsWithAnswers(res.page, res.studentId, answers, res.score, res.maxScore);
    } else if (APP_STATE.currentPage === 1 && APP_STATE.answerKey.length > 0) {
        // Show Answer Key for Page 1
        displayResultsWithAnswers(1, "正答キー", APP_STATE.answerKey, null, null);
    } else {
        document.getElementById('resultsPreview').style.display = 'none';
    }
}

function exportExcel() {
    const rawArgs = APP_STATE.grader.getRawData();
    if (!rawArgs.results.length) {
        alert('エクスポートするデータがありません。');
        return;
    }

    const { correctAnswers, results } = rawArgs;
    const numQ = correctAnswers.length;

    const wb = XLSX.utils.book_new();
    const ws_data = [];

    // Row 1: Headers
    const headers = ["Student ID"];
    for (let i = 1; i <= numQ; i++) headers.push(`Q${i}`);
    headers.push("Total Score");
    ws_data.push(headers);

    // Row 2: Points (Default 0, user sets this)
    const pointsRow = ["Points (Set values here)"];
    for (let i = 0; i < numQ; i++) pointsRow.push(0);
    pointsRow.push(null);
    ws_data.push(pointsRow);

    // Row 3: Correct Answers
    const keyRow = ["Correct Answer"];
    keyRow.push(...correctAnswers.map(a => a === null ? "" : a));
    keyRow.push(null);
    ws_data.push(keyRow);

    // Prepare Stats
    const correctCounts = new Array(numQ).fill(0);
    const validStudentCount = results.length;

    // Add Student Rows and Collect Stats
    results.forEach((r) => {
        const rowData = [r.studentId];
        r.details.forEach((d, qIdx) => {
            rowData.push(d.student === null ? "" : d.student);
            if (d.isCorrect) {
                if (correctCounts[qIdx] !== undefined) correctCounts[qIdx]++;
            }
        });
        ws_data.push(rowData);
    });

    // Create Sheet
    const ws = XLSX.utils.aoa_to_sheet(ws_data);

    // Column Index Utility
    if (!XLSX.utils.encode_col) {
        alert("SheetJS utility error.");
        return;
    }

    const startCol = XLSX.utils.encode_col(1); // 'B'
    const endCol = XLSX.utils.encode_col(numQ);
    const pointsRange = `$${startCol}$2:$${endCol}$2`;
    const keyRange = `$${startCol}$3:$${endCol}$3`;

    // Add Formulas (Overwriting total score cells)
    results.forEach((r, idx) => {
        const rowNum = idx + 4; // Data starts at Row 4
        const totalCellAddr = XLSX.utils.encode_cell({ r: rowNum - 1, c: numQ + 1 });
        const studentRange = `${startCol}${rowNum}:${endCol}${rowNum}`;
        const formula = `SUMPRODUCT((${studentRange}=${keyRange})*${pointsRange})`;
        ws[totalCellAddr] = { t: 'n', f: formula };
    });

    // Add Accuracy Summary Row
    const summaryRowIndex = 4 + results.length; // 1-based Row Index for Excel (logic: 3 header rows + N students + 1) -> Actually ws_data has 3+N rows. So next is 3+N. In 0-indexed array it is index 3+N.
    // Wait, aoa_to_sheet uses 0-indexed row/col for encode_cell.
    // ws_data length is (3 + results.length). So next row index is (3 + results.length).
    const nextRowIdx = 3 + results.length;

    const summaryLabelCell = XLSX.utils.encode_cell({ r: nextRowIdx, c: 0 });
    XLSX.utils.sheet_add_aoa(ws, [["正答率"]], { origin: summaryLabelCell });

    // Calculate Rates
    const statsRow = [];
    for (let i = 0; i < numQ; i++) {
        const rate = validStudentCount > 0 ? correctCounts[i] / validStudentCount : 0;
        statsRow.push(rate);
    }

    const statsStartCell = XLSX.utils.encode_cell({ r: nextRowIdx, c: 1 });
    XLSX.utils.sheet_add_aoa(ws, [statsRow], { origin: statsStartCell });

    // Format as Percentage
    for (let i = 0; i < numQ; i++) {
        const cellRef = XLSX.utils.encode_cell({ r: nextRowIdx, c: i + 1 });
        if (ws[cellRef]) {
            ws[cellRef].z = '0%';
        }
    }

    XLSX.utils.book_append_sheet(wb, ws, "Grading Results");
    XLSX.writeFile(wb, "grading_results.xlsx");
}

/* Student List Logic */
function toggleStudentList() {
    const modal = document.getElementById('studentListModal');

    if (!APP_STATE.grader.results.length) {
        alert('解析結果がありません。全ページ採点を行ってください。');
        return;
    }

    renderStudentList(); // Call the new rendering function
    modal.style.display = 'block';
}

function renderStudentList() {
    const tbody = document.querySelector('#studentListTable tbody');
    tbody.innerHTML = '';

    if (!APP_STATE.grader.results.length) return;

    // Get total expected marks from correct answer key
    const expectedMarkCount = APP_STATE.grader.correctAnswers.filter(a => a !== null).length;

    APP_STATE.grader.results.forEach(res => {
        const tr = document.createElement('tr');

        const answers = res.details.map(d => d.student);
        const hasMulti = answers.some(a => typeof a === 'string' && a.includes(','));

        // 1. Total Mark Count (including multiple marks)
        let totalMarks = 0;
        answers.forEach(a => {
            if (a === null) return;
            if (typeof a === 'string' && a.includes(',')) {
                totalMarks += a.split(',').length;
            } else {
                totalMarks += 1;
            }
        });

        // 2. Flags Logic
        // "未回答（検出）": Check if answered count matches expected count (e.g. 30/30)
        const answeredCount = answers.filter(a => a !== null).length;
        const hasUnanswered = answeredCount !== expectedMarkCount;

        // Discrepancy: Same as Unanswered effectively, or strict check
        // We'll treat them similarly based on user request
        const hasDiscrepancy = hasUnanswered;

        // Color-coded flags: red = problem, green = OK
        const multiFlag = hasMulti
            ? '<span style="color: red; font-weight: bold;">●</span>'
            : '<span style="color: green;">●</span>';

        // Unanswered: Red if count mismatch
        const emptyFlag = hasUnanswered
            ? '<span style="color: red; font-weight: bold;">●</span>'
            : '<span style="color: green;">●</span>';

        // Discrepancy: Red if count mismatch
        const discrepancyFlag = hasDiscrepancy
            ? '<span style="color: red; font-weight: bold;">●</span>'
            : '<span style="color: green;">●</span>';

        tr.innerHTML = `
            <td>${res.page}</td>
            <td>${res.studentId}</td>
            <td>${totalMarks}</td>
            <td>${multiFlag}</td>
            <td>${emptyFlag}</td>
            <td>${discrepancyFlag}</td>
        `;

        tr.onclick = () => {
            APP_STATE.currentPage = res.page;
            updateNav();
            requestRender().then(() => visualizeCurrentPage());
            document.getElementById('studentListModal').style.display = 'none';
        };

        tbody.appendChild(tr);
    });
}

/* Re-analyze Logic */
async function reanalyzeCurrentPage() {
    if (!APP_STATE.pdfHandler.pdfDoc) return;

    const btn = document.getElementById('reanalyzePageBtn');
    btn.disabled = true;
    document.body.style.cursor = 'wait';

    try {
        // Stabilization wait to ensure canvas is fully rendered
        await new Promise(resolve => setTimeout(resolve, 100));

        const imageData = APP_STATE.pdfHandler.getImageData(); // Get current cached image
        if (!imageData) throw new Error("No image data");

        // Apply contrast: Removed because imageData comes from canvas which already has contrast applied by requestRender
        // if (APP_STATE.config.contrast !== 0) {
        //    APP_STATE.ocrEngine.applyContrast(imageData, APP_STATE.config.contrast);
        // }

        // Binarize
        const binImage = APP_STATE.ocrEngine.binarize(imageData, APP_STATE.config.threshold);

        // Detect ID
        const idRes = APP_STATE.ocrEngine.detectMarks(binImage, APP_STATE.config.studentIdRegion, APP_STATE.config.studentIdGrid, APP_STATE.config.sensitivity);
        const studentId = APP_STATE.ocrEngine.readVerticalID(idRes.matrix);

        // Detect Answers
        const answers = [];
        const numBlocks = parseInt(document.getElementById('numBlocks').value);
        const qPerBlock = APP_STATE.config.questionsPerBlock;

        for (let i = 0; i < numBlocks; i++) {
            const blockConfig = APP_STATE.config.answerBlocks[i];
            if (!blockConfig) continue;

            const grid = { rows: qPerBlock, cols: 10 };
            const blockOptions = getBlockOptions(i);

            const ansRes = APP_STATE.ocrEngine.detectMarks(binImage, blockConfig, grid, APP_STATE.config.sensitivity);
            const blockAnswers = APP_STATE.ocrEngine.readHorizontalAnswers(ansRes.matrix, blockOptions);
            answers.push(...blockAnswers);
        }

        // Update Grader
        // Update Grader
        APP_STATE.grader.gradeStudent(studentId, answers, APP_STATE.currentPage);

        // Update View
        await visualizeCurrentPage();

        // Update student list if modal is open
        const modal = document.getElementById('studentListModal');
        if (modal && modal.style.display === 'block') {
            renderStudentList();
        }

        alert('再解析が完了しました。');

    } catch (e) {
        console.error(e);
        alert('再解析に失敗しました: ' + e.message);
    } finally {
        btn.disabled = false;
        document.body.style.cursor = 'default';
    }
}

/* PDF Export Logic */
async function exportPdf() {
    if (!APP_STATE.grader.results.length) {
        alert('エクスポートするデータがありません。');
        return;
    }

    try {
        const fileInput = document.getElementById('fileInput');
        if (!fileInput.files.length) return;

        const fileBuffer = await fileInput.files[0].arrayBuffer();
        const pdfDoc = await PDFLib.PDFDocument.load(fileBuffer);
        const pages = pdfDoc.getPages();

        const numBlocks = parseInt(document.getElementById('numBlocks').value);
        const qPerBlock = APP_STATE.config.questionsPerBlock;
        const scale = APP_STATE.pdfHandler.scale;

        APP_STATE.grader.results.forEach(res => {
            const pageIndex = res.page - 1;
            if (pageIndex < 0 || pageIndex >= pages.length) return;

            const page = pages[pageIndex];
            const { width, height } = page.getSize();

            const toPdfX = (x) => x / scale;
            const toPdfY = (y, h) => height - (y / scale) - (h / scale);
            const toPdfW = (w) => w / scale;
            const toPdfH = (h) => h / scale;

            res.details.forEach((d, globalIdx) => {
                const blockIdx = Math.floor(globalIdx / qPerBlock);
                if (blockIdx >= numBlocks) return;

                const inBlockIdx = globalIdx % qPerBlock;
                const block = APP_STATE.config.answerBlocks[blockIdx];
                if (!block) return;

                const cellW = block.w / 10;
                const cellH = block.h / qPerBlock;
                const y = block.y + inBlockIdx * cellH;

                let color = { r: 0.93, g: 0.26, b: 0.26 }; // Red
                let opacity = 0.3;

                if (d.student === null) {
                    color = { r: 1, g: 0.75, b: 0.03 }; // Yellow
                } else if (d.isCorrect) {
                    color = { r: 0.06, g: 0.72, b: 0.4 }; // Green
                }

                if (d.student === null) {
                    page.drawRectangle({
                        x: toPdfX(block.x),
                        y: toPdfY(y, cellH),
                        width: toPdfW(block.w),
                        height: toPdfH(cellH),
                        color: PDFLib.rgb(color.r, color.g, color.b),
                        opacity: opacity
                    });
                    return;
                }
                let indicesToDraw = [];
                if (typeof d.student === 'number') {
                    indicesToDraw.push(d.student === 0 ? 9 : d.student - 1);
                } else if (typeof d.student === 'string') {
                    d.student.split(',').forEach(v => {
                        const n = parseInt(v);
                        if (!isNaN(n)) indicesToDraw.push(n === 0 ? 9 : n - 1);
                    });
                }

                indicesToDraw.forEach(colIdx => {
                    const x = block.x + colIdx * cellW;
                    const drawX = x + 2;
                    const drawY = y + 2;
                    const drawW = cellW - 4;
                    const drawH = cellH - 4;

                    page.drawRectangle({
                        x: toPdfX(drawX),
                        y: toPdfY(drawY, drawH),
                        width: toPdfW(drawW),
                        height: toPdfH(drawH),
                        color: PDFLib.rgb(color.r, color.g, color.b),
                        opacity: opacity
                    });
                });
            });
        });

        const pdfBytes = await pdfDoc.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'graded_with_overlay.pdf';
        link.click();

    } catch (e) {
        console.error(e);
        alert('PDF保存に失敗しました: ' + e.message);
    }
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

