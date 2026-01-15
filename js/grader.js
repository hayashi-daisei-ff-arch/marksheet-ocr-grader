/**
 * Handles grading logic and data management
 */
class Grader {
    constructor() {
        this.correctAnswers = []; // Array of answers [1, 2, 0, ...]
        this.results = []; // Array of student results
    }

    /**
     * Set the answer key from the first page analysis
     * @param {Array} answers 
     */
    setAnswerKey(answers) {
        this.correctAnswers = answers;
        console.log("Answer Key set:", this.correctAnswers);
    }

    /**
     * Grade a single student's answers
     * @param {string} studentId 
     * @param {Array} studentAnswers 
     * @param {number} pageNum
     */
    gradeStudent(studentId, studentAnswers, pageNum) {
        // Compare with correctAnswers
        const gradedDetails = [];
        let score = 0;
        let maxScore = 0;

        // Use the length of correct answers as the basis
        for (let i = 0; i < this.correctAnswers.length; i++) {
            const correct = this.correctAnswers[i];
            const student = studentAnswers[i];

            // Skip if question doesn't exist in student sheet (e.g. partial scan)
            // But usually we expect full match.

            const isCorrect = (student === correct) && (correct !== null) && (correct !== undefined);
            if (isCorrect) score++;
            if (correct !== null && correct !== undefined) maxScore++;

            gradedDetails.push({
                question: i + 1,
                correct: correct,
                student: student,
                isCorrect: isCorrect
            });
        }

        const result = {
            page: pageNum,
            studentId,
            score,
            maxScore,
            details: gradedDetails,
            timestamp: new Date().toISOString()
        };

        this.results.push(result);
        return result;
    }

    /**
     * Export all results to CSV
     */
    exportCSV() {
        if (this.results.length === 0) return null;

        // Headers
        const headers = ["Page", "Student ID", "Score", "Max Score"];
        // Dynamic headers for each question
        const numQuestions = this.correctAnswers.length;
        for (let i = 1; i <= numQuestions; i++) {
            headers.push(`Q${i}`);
        }

        // Rows
        const rows = this.results.map(r => {
            const basicInfo = [r.page, r.studentId, r.score, r.maxScore];
            const answers = r.details.map(d => {
                if (d.student === null) return "";
                return d.student;
            });
            return [...basicInfo, ...answers].join(",");
        });

        return [headers.join(","), ...rows].join("\n");
    }

    /**
     * Get data object for Excel export
     * Structure:
     * - Headers: ["Student ID", "Q1", "Q2"..., "Total"]
     * - Points: ["Points", 0, 0..., (Formula)]
     * - Rows: [ID, 1/0, 1/0..., (Formula)]
     */
    getExcelData() {
        if (this.results.length === 0) return null;

        const numQuestions = this.correctAnswers.length;

        // 1. Header Row
        const headers = ["Student ID"];
        for (let i = 1; i <= numQuestions; i++) headers.push(`Q${i}`);
        headers.push("Total Score");

        // 2. Points Row (Default 0 per question)
        const pointsRow = ["Points (Set values here)"];
        for (let i = 0; i < numQuestions; i++) pointsRow.push(0);
        pointsRow.push(null); // Last cell reserved for potential total check or empty

        // 3. Student Data Rows (Binary 1/0 for Correct/Incorrect)
        const dataRows = this.results.map((r, rIdx) => {
            const rowData = [r.studentId];

            // Add 1 for correct, 0 for incorrect
            r.details.forEach(d => {
                rowData.push(d.isCorrect ? 1 : 0);
            });

            // Add Formula for Total Score
            // Formula: =SUMPRODUCT(PointsRow, ThisRowQuestions)
            // Points are in Row 2 (Index 2 in 1-based Excel). Columns B to (B+N).
            // Current Row is 3 + rIdx.
            const rowNum = 3 + rIdx; // Header is 1, Points is 2, Data starts 3
            // Columns: StudentID is A. Q1 is B.
            // Start Col: B (ASCII 66)
            // End Col: B + numQuestions - 1

            // Using R1C1 or A1 notation? SheetJS utility helps, but for raw strings we need logic.
            // Let's rely on app.js to construct the sheet with proper types.
            // Here we just return the raw data and let app.js handle SheetJS specifics.
            return rowData;
        });

        return {
            headers,
            pointsRow,
            dataRows
        };
    }

    /**
     * Get raw data for custom export formatting
     */
    getRawData() {
        return {
            correctAnswers: this.correctAnswers,
            results: this.results
        };
    }

    reset() {
        this.results = [];
    }
}
