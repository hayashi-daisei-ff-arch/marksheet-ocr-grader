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

    reset() {
        this.results = [];
    }
}
