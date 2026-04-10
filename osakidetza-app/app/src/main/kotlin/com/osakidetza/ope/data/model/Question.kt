package com.osakidetza.ope.data.model

data class Question(
    val id: Int,
    val text: String,
    val options: Map<String, String>,
    val correct: String,
    val explanation: String = ""
)

data class QuestionBank(
    val categoryId: String,
    val questions: List<Question>
)

data class Category(
    val id: String,
    val name: String,
    val description: String,
    val questionFile: String,
    val iconEmoji: String = "📋"
)

data class QuizSession(
    val categoryId: String,
    val questions: List<Question>,
    val mode: QuizMode,
    val timeLimitSeconds: Int = 0
)

enum class QuizMode {
    STUDY,   // Shows answer immediately after selection
    EXAM     // Shows all results at the end (timed)
}

data class QuizResult(
    val totalQuestions: Int,
    val correctAnswers: Int,
    val wrongAnswers: Int,
    val skippedAnswers: Int,
    val answers: List<AnswerRecord>,
    val timeTakenSeconds: Int = 0
) {
    val score: Float get() = if (totalQuestions > 0) correctAnswers.toFloat() / totalQuestions else 0f
    val scorePercent: Int get() = (score * 100).toInt()
    val passed: Boolean get() = scorePercent >= 50
}

data class AnswerRecord(
    val question: Question,
    val selectedAnswer: String?,
    val isCorrect: Boolean
)
