package com.osakidetza.ope.data

import android.content.Context
import androidx.datastore.preferences.core.edit
import com.google.gson.Gson
import com.google.gson.reflect.TypeToken
import com.osakidetza.ope.data.model.AnswerRecord
import com.osakidetza.ope.data.model.Category
import com.osakidetza.ope.data.model.CategoryStats
import com.osakidetza.ope.data.model.Question
import com.osakidetza.ope.data.model.QuizResult
import com.osakidetza.ope.data.model.StatsKeys
import com.osakidetza.ope.data.model.dataStore
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.first
import kotlinx.coroutines.flow.map
import kotlinx.coroutines.withContext

class QuestionRepository(private val context: Context) {

    private val gson = Gson()

    val categories = listOf(
        Category(
            id = "comun_abc1",
            name = "Común A, B y C1",
            description = "Temario común para categorías A, B y C1. Organización sanitaria, legislación laboral, igualdad, protección de datos.",
            questionFile = "questions/comun_abc1.json",
            iconEmoji = "🏥"
        ),
        Category(
            id = "comun_c2c3de",
            name = "Común C2, C3, D y E",
            description = "Temario común para categorías C2, C3, D y E. Organización básica, derechos laborales.",
            questionFile = "questions/comun_c2c3de.json",
            iconEmoji = "📋"
        ),
        Category(
            id = "enfermeria",
            name = "Enfermero/a",
            description = "Temario específico de Enfermería. Anatomía, fisiología, enfermería médico-quirúrgica, farmacología.",
            questionFile = "questions/enfermeria.json",
            iconEmoji = "💉"
        ),
        Category(
            id = "tcae",
            name = "TCAE / Auxiliar de Enfermería",
            description = "Técnico en Cuidados Auxiliares de Enfermería. Cuidados básicos, higiene, constantes vitales.",
            questionFile = "questions/tcae.json",
            iconEmoji = "🩺"
        ),
        Category(
            id = "celador",
            name = "Celador/a",
            description = "Transporte de pacientes, auxiliar en urgencias, mantenimiento del orden hospitalario.",
            questionFile = "questions/celador.json",
            iconEmoji = "🦺"
        ),
        Category(
            id = "aux_admin",
            name = "Auxiliar Administrativo/a",
            description = "Gestión de historias clínicas, atención al paciente, ofimática, citas.",
            questionFile = "questions/aux_admin.json",
            iconEmoji = "💼"
        ),
        Category(
            id = "administrativo",
            name = "Administrativo/a",
            description = "Administración pública, derecho administrativo, contratación pública, ofimática avanzada.",
            questionFile = "questions/administrativo.json",
            iconEmoji = "📁"
        ),
        Category(
            id = "laboratorio",
            name = "Técnico Especialista Laboratorio",
            description = "Bioquímica, hematología, microbiología, inmunología, seguridad en laboratorio.",
            questionFile = "questions/laboratorio.json",
            iconEmoji = "🔬"
        ),
        Category(
            id = "fisioterapeuta",
            name = "Fisioterapeuta",
            description = "Técnicas fisioterapéuticas, anatomía y biomecánica, rehabilitación neurológica y musculoesquelética.",
            questionFile = "questions/fisioterapeuta.json",
            iconEmoji = "🏃"
        ),
        Category(
            id = "tecnico_superior",
            name = "Técnico/a Superior Administración",
            description = "Administración y gestión sanitaria, derecho administrativo avanzado, supuestos prácticos.",
            questionFile = "questions/tecnico_superior.json",
            iconEmoji = "⚖️"
        )
    )

    suspend fun loadQuestions(categoryId: String): List<Question> = withContext(Dispatchers.IO) {
        val category = categories.find { it.id == categoryId } ?: return@withContext emptyList()
        try {
            val json = context.assets.open(category.questionFile).bufferedReader().use { it.readText() }
            val type = object : TypeToken<List<QuestionData>>() {}.type
            val questionDataList: List<QuestionData> = gson.fromJson(json, type)
            questionDataList.map { qd ->
                Question(
                    id = qd.id,
                    text = qd.text,
                    options = qd.options ?: emptyMap(),
                    correct = qd.correct ?: "",
                    explanation = qd.explanation ?: ""
                )
            }.filter { it.correct.isNotEmpty() && it.options.isNotEmpty() }
        } catch (e: Exception) {
            emptyList()
        }
    }

    fun getCategoryStats(categoryId: String): Flow<CategoryStats> {
        return context.dataStore.data.map { prefs ->
            CategoryStats(
                categoryId = categoryId,
                totalSessions = prefs[StatsKeys.totalSessions(categoryId)] ?: 0,
                totalCorrect = prefs[StatsKeys.totalCorrect(categoryId)] ?: 0,
                totalQuestions = prefs[StatsKeys.totalQuestions(categoryId)] ?: 0,
                bestScore = prefs[StatsKeys.bestScore(categoryId)] ?: 0f
            )
        }
    }

    suspend fun saveQuizResult(categoryId: String, result: QuizResult) {
        context.dataStore.edit { prefs ->
            val prevSessions = prefs[StatsKeys.totalSessions(categoryId)] ?: 0
            val prevCorrect = prefs[StatsKeys.totalCorrect(categoryId)] ?: 0
            val prevTotal = prefs[StatsKeys.totalQuestions(categoryId)] ?: 0
            val prevBest = prefs[StatsKeys.bestScore(categoryId)] ?: 0f

            prefs[StatsKeys.totalSessions(categoryId)] = prevSessions + 1
            prefs[StatsKeys.totalCorrect(categoryId)] = prevCorrect + result.correctAnswers
            prefs[StatsKeys.totalQuestions(categoryId)] = prevTotal + result.totalQuestions
            if (result.score > prevBest) {
                prefs[StatsKeys.bestScore(categoryId)] = result.score
            }
        }
    }

    fun getBookmarkedIds(): Flow<Set<String>> {
        return context.dataStore.data.map { prefs ->
            prefs[StatsKeys.bookmarkedQuestions] ?: emptySet()
        }
    }

    suspend fun toggleBookmark(questionId: String) {
        context.dataStore.edit { prefs ->
            val current = prefs[StatsKeys.bookmarkedQuestions]?.toMutableSet() ?: mutableSetOf()
            if (current.contains(questionId)) {
                current.remove(questionId)
            } else {
                current.add(questionId)
            }
            prefs[StatsKeys.bookmarkedQuestions] = current
        }
    }

    private data class QuestionData(
        val id: Int,
        val text: String,
        val options: Map<String, String>,
        val correct: String,
        val explanation: String?
    )
}
