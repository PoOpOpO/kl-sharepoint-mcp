package com.osakidetza.ope.viewmodel

import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.osakidetza.ope.data.QuestionRepository
import com.osakidetza.ope.data.model.AnswerRecord
import com.osakidetza.ope.data.model.Question
import com.osakidetza.ope.data.model.QuizMode
import com.osakidetza.ope.data.model.QuizResult
import kotlinx.coroutines.Job
import kotlinx.coroutines.delay
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.launch

sealed class QuizUiState {
    object Loading : QuizUiState()
    data class Ready(
        val currentIndex: Int,
        val questions: List<Question>,
        val selectedAnswer: String?,
        val showResult: Boolean,
        val mode: QuizMode,
        val timeRemainingSeconds: Int,
        val isBookmarked: Boolean,
        val answers: Map<Int, String?>
    ) : QuizUiState()
    data class Finished(val result: QuizResult) : QuizUiState()
    object Error : QuizUiState()
}

class QuizViewModel(application: Application) : AndroidViewModel(application) {

    private val repository = QuestionRepository(application)

    private val _uiState = MutableStateFlow<QuizUiState>(QuizUiState.Loading)
    val uiState: StateFlow<QuizUiState> = _uiState.asStateFlow()

    private var questions: List<Question> = emptyList()
    private var currentIndex = 0
    private var answers: MutableMap<Int, String?> = mutableMapOf()
    private var mode: QuizMode = QuizMode.STUDY
    private var categoryId: String = ""
    private var timerJob: Job? = null
    private var timeRemaining = 0
    private var startTime = 0L
    private var bookmarkedIds: Set<String> = emptySet()

    fun startQuiz(
        categoryId: String,
        mode: QuizMode,
        questionCount: Int = 20,
        shuffled: Boolean = true
    ) {
        this.categoryId = categoryId
        this.mode = mode
        viewModelScope.launch {
            _uiState.value = QuizUiState.Loading
            val allQuestions = repository.loadQuestions(categoryId)
            if (allQuestions.isEmpty()) {
                _uiState.value = QuizUiState.Error
                return@launch
            }
            val selected = if (shuffled) allQuestions.shuffled() else allQuestions
            questions = selected.take(minOf(questionCount, allQuestions.size))
            currentIndex = 0
            answers = mutableMapOf()
            startTime = System.currentTimeMillis()

            // Time limit: 1.5 minutes per question in exam mode
            timeRemaining = if (mode == QuizMode.EXAM) questions.size * 90 else 0

            // Load bookmarks
            repository.getBookmarkedIds().collect { ids ->
                bookmarkedIds = ids
            }

            if (mode == QuizMode.EXAM && timeRemaining > 0) {
                startTimer()
            }
            updateState()
        }
    }

    fun selectAnswer(answer: String) {
        val state = _uiState.value as? QuizUiState.Ready ?: return
        if (state.selectedAnswer != null && mode == QuizMode.STUDY) return
        answers[currentIndex] = answer
        updateState(selectedAnswer = answer, showResult = mode == QuizMode.STUDY)
    }

    fun nextQuestion() {
        if (currentIndex < questions.size - 1) {
            currentIndex++
            updateState()
        } else {
            finishQuiz()
        }
    }

    fun previousQuestion() {
        if (currentIndex > 0) {
            currentIndex--
            updateState()
        }
    }

    fun skipQuestion() {
        answers[currentIndex] = null
        nextQuestion()
    }

    fun toggleBookmark() {
        val question = questions.getOrNull(currentIndex) ?: return
        val questionKey = "${categoryId}_${question.id}"
        viewModelScope.launch {
            repository.toggleBookmark(questionKey)
        }
    }

    fun finishQuiz() {
        timerJob?.cancel()
        val timeTaken = ((System.currentTimeMillis() - startTime) / 1000).toInt()
        val answerRecords = questions.mapIndexed { index, question ->
            val selected = answers[index]
            AnswerRecord(
                question = question,
                selectedAnswer = selected,
                isCorrect = selected == question.correct
            )
        }
        val result = QuizResult(
            totalQuestions = questions.size,
            correctAnswers = answerRecords.count { it.isCorrect },
            wrongAnswers = answerRecords.count { !it.isCorrect && it.selectedAnswer != null },
            skippedAnswers = answerRecords.count { it.selectedAnswer == null },
            answers = answerRecords,
            timeTakenSeconds = timeTaken
        )
        viewModelScope.launch {
            repository.saveQuizResult(categoryId, result)
        }
        _uiState.value = QuizUiState.Finished(result)
    }

    private fun startTimer() {
        timerJob?.cancel()
        timerJob = viewModelScope.launch {
            while (timeRemaining > 0) {
                delay(1000)
                timeRemaining--
                updateState()
                if (timeRemaining == 0) {
                    finishQuiz()
                }
            }
        }
    }

    private fun updateState(
        selectedAnswer: String? = answers[currentIndex],
        showResult: Boolean = answers.containsKey(currentIndex) && mode == QuizMode.STUDY
    ) {
        val question = questions.getOrNull(currentIndex) ?: return
        val questionKey = "${categoryId}_${question.id}"
        _uiState.value = QuizUiState.Ready(
            currentIndex = currentIndex,
            questions = questions,
            selectedAnswer = selectedAnswer,
            showResult = showResult,
            mode = mode,
            timeRemainingSeconds = timeRemaining,
            isBookmarked = bookmarkedIds.contains(questionKey),
            answers = answers.toMap()
        )
    }

    override fun onCleared() {
        super.onCleared()
        timerJob?.cancel()
    }
}
