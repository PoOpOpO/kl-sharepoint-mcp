package com.osakidetza.ope.ui.screens

import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.itemsIndexed
import androidx.compose.foundation.shape.CircleShape
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Check
import androidx.compose.material.icons.filled.Close
import androidx.compose.material.icons.filled.Home
import androidx.compose.material.icons.filled.Refresh
import androidx.compose.material.icons.filled.Remove
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Brush
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import com.osakidetza.ope.data.model.AnswerRecord
import com.osakidetza.ope.data.model.QuizResult
import com.osakidetza.ope.ui.theme.CorrectGreen
import com.osakidetza.ope.ui.theme.CorrectGreenLight
import com.osakidetza.ope.ui.theme.OsakidetзaBlue
import com.osakidetza.ope.ui.theme.OsakidetзaRed
import com.osakidetza.ope.ui.theme.WrongRed
import com.osakidetza.ope.ui.theme.WrongRedLight
import com.osakidetza.ope.viewmodel.QuizUiState
import com.osakidetza.ope.viewmodel.QuizViewModel

@Composable
fun ResultScreen(
    onHome: () -> Unit,
    onRetry: () -> Unit,
    viewModel: QuizViewModel
) {
    val uiState by viewModel.uiState.collectAsState()
    val result = (uiState as? QuizUiState.Finished)?.result ?: return

    var showDetails by remember { mutableStateOf(false) }

    LazyColumn(
        modifier = Modifier.fillMaxSize(),
        contentPadding = PaddingValues(bottom = 32.dp)
    ) {
        item {
            // Header
            Box(
                modifier = Modifier
                    .fillMaxWidth()
                    .background(
                        Brush.verticalGradient(
                            colors = if (result.passed)
                                listOf(CorrectGreen, Color(0xFF1B5E20))
                            else listOf(WrongRed, Color(0xFF7F0000))
                        )
                    )
                    .padding(top = 56.dp, bottom = 40.dp, start = 24.dp, end = 24.dp),
                contentAlignment = Alignment.Center
            ) {
                Column(horizontalAlignment = Alignment.CenterHorizontally) {
                    // Score circle
                    Box(
                        modifier = Modifier
                            .size(120.dp)
                            .clip(CircleShape)
                            .background(Color.White.copy(0.2f)),
                        contentAlignment = Alignment.Center
                    ) {
                        Column(horizontalAlignment = Alignment.CenterHorizontally) {
                            Text(
                                text = "${result.scorePercent}%",
                                style = MaterialTheme.typography.displaySmall,
                                fontWeight = FontWeight.ExtraBold,
                                color = Color.White
                            )
                        }
                    }
                    Spacer(modifier = Modifier.height(16.dp))
                    Text(
                        text = if (result.passed) "¡Aprobado!" else "Suspendido",
                        style = MaterialTheme.typography.headlineMedium,
                        fontWeight = FontWeight.Bold,
                        color = Color.White
                    )
                    Text(
                        text = if (result.passed) "Sigue practicando para mejorar" else "No te rindas, sigue practicando",
                        style = MaterialTheme.typography.bodyMedium,
                        color = Color.White.copy(0.85f)
                    )
                }
            }
        }

        item {
            Spacer(modifier = Modifier.height(16.dp))
            // Score breakdown
            Row(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(horizontal = 16.dp),
                horizontalArrangement = Arrangement.spacedBy(12.dp)
            ) {
                ScoreCard(
                    label = "Correctas",
                    value = result.correctAnswers.toString(),
                    color = CorrectGreen,
                    bgColor = CorrectGreenLight,
                    modifier = Modifier.weight(1f)
                )
                ScoreCard(
                    label = "Incorrectas",
                    value = result.wrongAnswers.toString(),
                    color = WrongRed,
                    bgColor = WrongRedLight,
                    modifier = Modifier.weight(1f)
                )
                ScoreCard(
                    label = "Sin contestar",
                    value = result.skippedAnswers.toString(),
                    color = MaterialTheme.colorScheme.onSurfaceVariant,
                    bgColor = MaterialTheme.colorScheme.surfaceVariant,
                    modifier = Modifier.weight(1f)
                )
            }
        }

        item {
            Spacer(modifier = Modifier.height(8.dp))
            if (result.timeTakenSeconds > 0) {
                val mins = result.timeTakenSeconds / 60
                val secs = result.timeTakenSeconds % 60
                Text(
                    text = "Tiempo empleado: %d min %02d seg".format(mins, secs),
                    style = MaterialTheme.typography.bodySmall,
                    color = MaterialTheme.colorScheme.onSurfaceVariant,
                    modifier = Modifier.padding(horizontal = 16.dp)
                )
            }
        }

        item {
            Spacer(modifier = Modifier.height(16.dp))
            // Action buttons
            Row(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(horizontal = 16.dp),
                horizontalArrangement = Arrangement.spacedBy(12.dp)
            ) {
                OutlinedButton(
                    onClick = onHome,
                    modifier = Modifier.weight(1f),
                    shape = RoundedCornerShape(12.dp)
                ) {
                    Icon(Icons.Filled.Home, null)
                    Spacer(Modifier.width(8.dp))
                    Text("Inicio")
                }
                Button(
                    onClick = onRetry,
                    modifier = Modifier.weight(1f),
                    shape = RoundedCornerShape(12.dp),
                    colors = ButtonDefaults.buttonColors(containerColor = OsakidetзaRed)
                ) {
                    Icon(Icons.Filled.Refresh, null)
                    Spacer(Modifier.width(8.dp))
                    Text("Repetir")
                }
            }
        }

        item {
            Spacer(modifier = Modifier.height(16.dp))
            // Details toggle
            TextButton(
                onClick = { showDetails = !showDetails },
                modifier = Modifier.padding(horizontal = 16.dp)
            ) {
                Text(if (showDetails) "Ocultar respuestas" else "Ver respuestas detalladas")
            }
        }

        if (showDetails) {
            itemsIndexed(result.answers) { index, record ->
                AnswerReviewItem(index + 1, record)
                HorizontalDivider(modifier = Modifier.padding(horizontal = 16.dp))
            }
        }
    }
}

@Composable
private fun ScoreCard(
    label: String,
    value: String,
    color: Color,
    bgColor: Color,
    modifier: Modifier = Modifier
) {
    Card(
        modifier = modifier,
        shape = RoundedCornerShape(12.dp),
        colors = CardDefaults.cardColors(containerColor = bgColor)
    ) {
        Column(
            modifier = Modifier.padding(12.dp),
            horizontalAlignment = Alignment.CenterHorizontally
        ) {
            Text(
                text = value,
                style = MaterialTheme.typography.headlineSmall,
                fontWeight = FontWeight.Bold,
                color = color
            )
            Text(
                text = label,
                style = MaterialTheme.typography.labelSmall,
                color = color.copy(alpha = 0.8f)
            )
        }
    }
}

@Composable
private fun AnswerReviewItem(index: Int, record: AnswerRecord) {
    val bgColor = when {
        record.isCorrect -> CorrectGreenLight
        record.selectedAnswer != null -> WrongRedLight
        else -> MaterialTheme.colorScheme.surfaceVariant
    }
    val icon = when {
        record.isCorrect -> Icons.Filled.Check
        record.selectedAnswer != null -> Icons.Filled.Close
        else -> Icons.Filled.Remove
    }
    val iconTint = when {
        record.isCorrect -> CorrectGreen
        record.selectedAnswer != null -> WrongRed
        else -> MaterialTheme.colorScheme.onSurfaceVariant
    }

    Row(
        modifier = Modifier
            .fillMaxWidth()
            .background(bgColor)
            .padding(16.dp),
        verticalAlignment = Alignment.Top
    ) {
        Box(
            modifier = Modifier
                .size(28.dp)
                .clip(CircleShape)
                .background(iconTint),
            contentAlignment = Alignment.Center
        ) {
            Icon(icon, null, tint = Color.White, modifier = Modifier.size(16.dp))
        }
        Spacer(Modifier.width(12.dp))
        Column(modifier = Modifier.weight(1f)) {
            Text(
                text = "$index. ${record.question.text}",
                style = MaterialTheme.typography.bodySmall,
                fontWeight = FontWeight.Medium
            )
            Spacer(Modifier.height(4.dp))
            Text(
                text = "Correcta: ${record.question.correct}) ${record.question.options[record.question.correct] ?: ""}",
                style = MaterialTheme.typography.bodySmall,
                color = CorrectGreen,
                fontWeight = FontWeight.SemiBold
            )
            if (!record.isCorrect && record.selectedAnswer != null) {
                Text(
                    text = "Tu respuesta: ${record.selectedAnswer}) ${record.question.options[record.selectedAnswer] ?: ""}",
                    style = MaterialTheme.typography.bodySmall,
                    color = WrongRed
                )
            }
            if (record.question.explanation.isNotBlank()) {
                Spacer(Modifier.height(4.dp))
                Text(
                    text = record.question.explanation,
                    style = MaterialTheme.typography.bodySmall,
                    color = MaterialTheme.colorScheme.onSurfaceVariant
                )
            }
        }
    }
}
