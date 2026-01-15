/**
 * OCR Engine for processing mark sheet images
 */
class OCREngine {
    constructor() {
        this.binarizedData = null; // Cache binarized image data
    }

    /**
     * Binarize image data with improved pink/red line filtering
     * @param {ImageData} imageData 
     * @param {number} threshold 0-255
     * @returns {ImageData} Binarized (Black/White) ImageData
     */
    binarize(imageData, threshold = 128) {
        const w = imageData.width;
        const h = imageData.height;
        const data = imageData.data;
        const newData = new Uint8ClampedArray(data.length);

        for (let i = 0; i < data.length; i += 4) {
            const r = data[i];
            const g = data[i + 1];
            const b = data[i + 2];

            // Detect pink/red lines (high R, low B, moderate G)
            // Pink grid lines typically have R > 200, G > 150, B < 150
            const isPinkish = (r > 180 && g > 120 && b < 160 && r > b + 30);

            // Calculate grayscale
            const brightness = 0.299 * r + 0.587 * g + 0.114 * b;

            // If it's pinkish, treat as white (background)
            // Otherwise use brightness threshold
            let val;
            if (isPinkish) {
                val = 255; // White (ignore pink lines)
            } else {
                val = brightness < threshold ? 0 : 255;
            }

            newData[i] = val;     // R
            newData[i + 1] = val;   // G
            newData[i + 2] = val;   // B
            newData[i + 3] = 255;   // Alpha
        }

        return new ImageData(newData, w, h);
    }

    /**
     * Detect marks in a specific region
     * @param {ImageData} imageData Source image (can be binarized or raw)
     * @param {Object} region {x, y, w, h}
     * @param {Object} grid {rows, cols}
     * @param {number} sensitivity 0.0-1.0 (Fill ratio threshold)
     * @param {number} threshold 0-255 (Binarization threshold if not already binarized)
     */
    detectMarks(imageData, region, grid, sensitivity = 0.3, threshold = 128) {
        const { x, y, w, h } = region;
        const { rows, cols } = grid;
        const cellW = w / cols;
        const cellH = h / rows;

        const results = [];
        const debugGrids = [];

        const data = imageData.data;
        const width = imageData.width;

        for (let r = 0; r < rows; r++) {
            const rowResults = [];
            for (let c = 0; c < cols; c++) {
                const cellX = Math.floor(x + c * cellW);
                const cellY = Math.floor(y + r * cellH);
                const cellW_int = Math.floor(cellW);
                const cellH_int = Math.floor(cellH);

                // Analyze inner 80% of the cell (10% margin on each side)
                const marginX = Math.floor(cellW_int * 0.1);
                const marginY = Math.floor(cellH_int * 0.1);

                let blackPixels = 0;
                let totalPixels = 0;

                for (let py = cellY + marginY; py < cellY + cellH_int - marginY; py++) {
                    for (let px = cellX + marginX; px < cellX + cellW_int - marginX; px++) {
                        // Bounds check
                        if (px < 0 || px >= width || py < 0 || py >= imageData.height) continue;

                        const idx = (py * width + px) * 4;

                        // For binarized images: R channel is either 0 (black) or 255 (white)
                        // Count black pixels (value close to 0)
                        if (data[idx] < 128) {
                            blackPixels++;
                        }
                        totalPixels++;
                    }
                }

                const ratio = totalPixels > 0 ? blackPixels / totalPixels : 0;
                const isMarked = ratio > sensitivity;

                rowResults.push({
                    value: c,
                    isMarked: isMarked,
                    ratio: ratio,
                    rect: { x: cellX, y: cellY, w: cellW_int, h: cellH_int }
                });

                debugGrids.push({
                    x: cellX, y: cellY, w: cellW_int, h: cellH_int,
                    isMarked,
                    ratio
                });
            }
            results.push(rowResults);
        }

        return { matrix: results, debug: debugGrids };
    }

    /**
     * Interpret grid results for Student ID (Vertical columns)
     * @param {Array} matrix Results from detectMarks
     * @returns {String} Student ID
     */
    readVerticalID(matrix) {
        // Matrix is [Rows][Cols]
        // But for vertical ID, each column represents a digit, and the row represents the value 0-9
        // Wait, usually ID is written horizontally? 
        // User spec: "1. 0~9のマーク（縦方向）で，それが7桁です"
        // This implies 7 COLUMNS, and in each column, rows 0-9 exist vertically.
        // So we scan each COLUMN to find the marked ROW.

        if (!matrix || matrix.length === 0) return "";

        const numCols = matrix[0].length;
        const numRows = matrix.length;
        let id = "";

        for (let c = 0; c < numCols; c++) {
            let foundDigit = "?";
            // Find which row is marked in this column
            for (let r = 0; r < numRows; r++) {
                if (matrix[r][c].isMarked) {
                    foundDigit = r.toString(); // Assuming row 0 is '0', row 1 is '1'...
                    break;
                }
            }
            id += foundDigit;
        }
        return id;
    }

    /**
     * Interpret grid results for Answers (Horizontal rows)
     * @param {Array} matrix Results from detectMarks
     * @returns {Array} Array of answers
     */
    readHorizontalAnswers(matrix) {
        // Standard mark sheet: Each row is a question. Columns are options.
        // IMPORTANT: Physical layout is 1,2,3,4,5,6,7,8,9,0
        // So column 0 = value 1, column 1 = value 2, ..., column 8 = value 9, column 9 = value 0
        return matrix.map(row => {
            const markedIndices = [];

            row.forEach((cell, colIdx) => {
                if (cell.isMarked) {
                    // Map column index to actual value (1-9,0 order)
                    const value = colIdx === 9 ? 0 : colIdx + 1;
                    markedIndices.push(value);
                }
            });

            if (markedIndices.length === 0) return null; // No answer
            if (markedIndices.length === 1) return markedIndices[0];
            return "MULTIPLE"; // Multiple marks error
        });
    }
}
