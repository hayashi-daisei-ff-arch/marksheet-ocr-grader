/**
 * Handles PDF file loading and rendering using PDF.js
 */
class PDFHandler {
    constructor() {
        this.pdfDoc = null;
        this.pageNum = 1;
        this.pageRendering = false;
        this.pageNumPending = null;
        this.scale = 1.5; // High resolution rendering
        this.canvas = document.getElementById('pdfCanvas');
        this.ctx = this.canvas.getContext('2d');
        this.overlayCanvas = document.getElementById('overlayCanvas');
    }

    /**
     * Load a PDF file from input
     * @param {File} file 
     */
    async load(file) {
        if (!file) return;

        try {
            const arrayBuffer = await file.arrayBuffer();
            this.pdfDoc = await pdfjsLib.getDocument(arrayBuffer).promise;
            return this.pdfDoc.numPages;
        } catch (error) {
            console.error('Error loading PDF:', error);
            throw error;
        }
    }

    /**
     * Render a specific page
     * @param {number} num Page number (1-based)
     */
    async renderPage(num) {
        if (!this.pdfDoc) return;

        this.pageRendering = true;
        
        try {
            const page = await this.pdfDoc.getPage(num);
            
            // Calculate scale to fit width if needed, or stick to fixed high res
            // For OMR, fixed high res is better, but visual fit is important too.
            // Let's use a scale that gives good pixel density (e.g. 2.0)
            const viewport = page.getViewport({scale: this.scale});
            
            this.canvas.height = viewport.height;
            this.canvas.width = viewport.width;
            
            // Sync overlay canvas size
            this.overlayCanvas.height = viewport.height;
            this.overlayCanvas.width = viewport.width;

            // Render PDF page into canvas context
            const renderContext = {
                canvasContext: this.ctx,
                viewport: viewport
            };
            
            await page.render(renderContext).promise;
            
            this.pageRendering = false;
            this.pageNum = num;

            if (this.pageNumPending !== null) {
                // New page rendering is pending
                this.renderPage(this.pageNumPending);
                this.pageNumPending = null;
            }

            return viewport; // Return viewport to help with coordinate mapping
        } catch (error) {
            console.error('Page render error:', error);
            this.pageRendering = false;
        }
    }

    /**
     * Queue a page for rendering
     * @param {number} num 
     */
    queueRenderPage(num) {
        if (this.pageRendering) {
            this.pageNumPending = num;
        } else {
            this.renderPage(num);
        }
    }

    /**
     * Get ImageData from the current canvas
     * Used by OCR engine
     */
    getImageData() {
        return this.ctx.getImageData(0, 0, this.canvas.width, this.canvas.height);
    }
}
