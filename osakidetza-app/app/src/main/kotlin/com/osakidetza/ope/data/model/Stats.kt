package com.osakidetza.ope.data.model

import android.content.Context
import androidx.datastore.preferences.core.floatPreferencesKey
import androidx.datastore.preferences.core.intPreferencesKey
import androidx.datastore.preferences.core.stringSetPreferencesKey
import androidx.datastore.preferences.preferencesDataStore

val Context.dataStore by preferencesDataStore(name = "ope_stats")

object StatsKeys {
    fun totalSessions(categoryId: String) = intPreferencesKey("total_sessions_$categoryId")
    fun totalCorrect(categoryId: String) = intPreferencesKey("total_correct_$categoryId")
    fun totalQuestions(categoryId: String) = intPreferencesKey("total_questions_$categoryId")
    fun bestScore(categoryId: String) = floatPreferencesKey("best_score_$categoryId")
    val bookmarkedQuestions = stringSetPreferencesKey("bookmarked_questions")
}

data class CategoryStats(
    val categoryId: String,
    val totalSessions: Int = 0,
    val totalCorrect: Int = 0,
    val totalQuestions: Int = 0,
    val bestScore: Float = 0f
) {
    val averageScore: Float get() = if (totalQuestions > 0) totalCorrect.toFloat() / totalQuestions else 0f
    val averageScorePercent: Int get() = (averageScore * 100).toInt()
    val bestScorePercent: Int get() = (bestScore * 100).toInt()
}
