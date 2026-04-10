package com.osakidetza.ope.ui.screens

import androidx.activity.compose.BackHandler
import androidx.compose.animation.AnimatedContent
import androidx.compose.animation.core.tween
import androidx.compose.animation.fadeIn
import androidx.compose.animation.fadeOut
import androidx.compose.animation.togetherWith
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.automirrored.filled.ArrowForward
import androidx.compose.material.icons.filled.BookmarkAdd
import androidx.compose.material.icons.filled.BookmarkAdded
import androidx.compose.material.icons.filled.Check
import androidx.compose.material.icons.filled.Close
import androidx.compose.material.icons.filled.Flag
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.lifecycle.viewmodel.compose.viewModel
import com.osakidetza.ope.data.model.QuizMode
import com.osakidetza.ope.data.model.QuizResult
import com.osakidetza.ope.ui.theme.CorrectGreen
import com.osakidetza.ope.ui.theme.CorrectGreenLight
import com.osakidetza.ope.ui.theme.OsakidetзaRed
import com.osakidetza.ope.ui.theme.WrongRed
import com.osakidetza.ope.ui.theme.WrongRedLight
import com.osakidetza.ope.viewmodel.QuizUiState
import com.osakidetza.ope.viewmodel.QuizViewModel

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun QuizScreen(
    categoryId: String,
    mode: QuizMode,
    questionCount: Int,
    onFinished: (QuizResult) -> Unit,
    onBack: () -> Unit,
    viewModel: QuizViewModel = viewModel()
) {
    var showExitDialog by remember { mutableStateOf(false) }
    val uiState by viewModel.uiState.collectAsState()

    LaunchedEffect(categoryId, mode, questionCount) {
        viewModel.startQuiz(categoryId, mode, questionCount)
    }

    BackHandler {
        showExitDialog = true
    }

    if (showExitDialog) {
        AlertDialog(
            onDismissRequest = { showExitDialog = false },
            title = { Text("¿Salir del examen?") },
            text = { Text("Tu progreso en este ejercicio se perderá.") },
            confirmButton = {
                TextButton(onClick = onBack) { Text("Salir", color = WrongRed) }
            },
            dismissButton = {
                TextButton(onClick = { showExitDialog = false }) { Text("Continuar") }
            }
        )
    }

    when (val state = uiState) {
        is QuizUiState.Loading -> {
            Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                Column(horizontalAlignment = Alignment.CenterHorizontally) {
                    CircularProgressIndicator(color = OsakidetзaRed)
                    Spacer(Modifier.height(16.dp))
                    Text("Cargando preguntas...")
                }
            }
        }

        is QuizUiState.Error -> {
            Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                Column(horizontalAlignment = Alignment.CenterHorizontally, modifier = Modifier.padding(32.dp)) {
                    Text("Error al cargar preguntas", style = MaterialTheme.typography.titleMedium)
                    Spacer(Modifier.height(16.dp))
                    Button(onClick = onBack) { Text("Volver") }
                }
            }
        }

        is QuizUiState.Ready -> {
            val question = state.questions[state.currentIndex]
            Scaffold(
                topBar = {
                    Column {
                        TopAppBar(
                            title = {
                                Text(
                                    "${state.currentIndex + 1} / ${state.questions.size}",
                                    style = MaterialTheme.typography.titleMedium
                                )
                            },
                            navigationIcon = {
                                IconButton(onClick = { showExitDialog = true }) {
                                    Icon(Icons.AutoMirrored.Filled.ArrowBack, "Salir")
                                }
                            },
                            actions = {
                                if (state.mode == QuizMode.EXAM && state.timeRemainingSeconds > 0) {
                                    val mins = state.timeRemainingSeconds / 60
                                    val secs = state.timeRemainingSeconds % 60
                                    val isUrgent = state.timeRemainingSeconds < 60
                                    Text(
                                        text = "%02d:%02d".format(mins, secs),
                                        color = if (isUrgent) WrongRed else Color.White,
                                        fontWeight = if (isUrgent) FontWeight.Bold else FontWeight.Normal,
                                        modifier = Modifier.padding(end = 8.dp)
                                    )
                                }
                                IconButton(onClick = viewModel::toggleBookmark) {
                                    Icon(
                                        if (state.isBookmarked) Icons.Filled.BookmarkAdded else Icons.Filled.BookmarkAdd,
                                        contentDescription = "Marcar",
                                        tint = Color.White
                                    )
                                }
                            },
                            colors = TopAppBarDefaults.topAppBarColors(
                                containerColor = OsakidetзaRed,
                                titleContentColor = Color.White,
                                navigationIconContentColor = Color.White,
                                actionIconContentColor = Color.White
                            )
                        )
                        LinearProgressIndicator(
                            progress = { (state.currentIndex + 1).toFloat() / state.questions.size },
                            modifier = Modifier.fillMaxWidth(),
                            color = Color.White.copy(0.8f),
                            trackColor = OsakidetзaRed.copy(alpha = 0.3f)
                        )
                    }
                },
                bottomBar = {
                    Surface(shadowElevation = 8.dp) {
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .padding(16.dp),
                            horizontalArrangement = Arrangement.spacedBy(8.dp)
                        ) {
                            if (state.currentIndex > 0) {
                                OutlinedButton(
                                    onClick = viewModel::previousQuestion,
                                    modifier = Modifier.weight(1f)
                                ) {
                                    Icon(Icons.AutoMirrored.Filled.ArrowBack, null)
                                    Spacer(Modifier.width(4.dp))
                                    Text("Anterior")
                                }
                            }

                            if (state.mode == QuizMode.EXAM && state.selectedAnswer == null) {
                                OutlinedButton(
                                    onClick = viewModel::skipQuestion,
                                    modifier = Modifier.weight(1f)
                                ) {
                                    Text("Omitir")
                                }
                            }

                            val isLast = state.currentIndex == state.questions.size - 1
                            if (state.selectedAnswer != null || state.mode == QuizMode.EXAM) {
                                Button(
                                    onClick = {
                                        if (isLast) viewModel.finishQuiz()
                                        else viewModel.nextQuestion()
                                    },
                                    modifier = Modifier.weight(1f),
                                    colors = ButtonDefaults.buttonColors(
                                        containerColor = if (isLast) MaterialTheme.colorScheme.tertiary
                                        else OsakidetзaRed
                                    ),
                                    enabled = state.selectedAnswer != null || state.mode == QuizMode.EXAM
                                ) {
                                    if (isLast) {
                                        Icon(Icons.Filled.Flag, null)
                                        Spacer(Modifier.width(4.dp))
                                        Text("Finalizar")
                                    } else {
                                        Text("Siguiente")
                                        Spacer(Modifier.width(4.dp))
                                        Icon(Icons.AutoMirrored.Filled.ArrowForward, null)
                                    }
                                }
                            }
                        }
                    }
                }
            ) { padding ->
                AnimatedContent(
                    targetState = state.currentIndex,
                    transitionSpec = {
                        fadeIn(tween(200)) togetherWith fadeOut(tween(200))
                    },
                    label = "question"
                ) { _ ->
                    Column(
                        modifier = Modifier
                            .padding(padding)
                            .fillMaxSize()
                            .verticalScroll(rememberScrollState())
                            .padding(16.dp),
                        verticalArrangement = Arrangement.spacedBy(12.dp)
                    ) {
                        // Question card
                        Card(
                            shape = RoundedCornerShape(16.dp),
                            colors = CardDefaults.cardColors(
                                containerColor = MaterialTheme.colorScheme.surfaceVariant
                            )
                        ) {
                            Column(modifier = Modifier.padding(20.dp)) {
                                Text(
                                    text = "Pregunta ${state.currentIndex + 1}",
                                    style = MaterialTheme.typography.labelMedium,
                                    color = MaterialTheme.colorScheme.onSurfaceVariant
                                )
                                Spacer(modifier = Modifier.height(8.dp))
                                Text(
                                    text = question.text,
                                    style = MaterialTheme.typography.bodyLarge,
                                    fontWeight = FontWeight.Medium,
                                    lineHeight = MaterialTheme.typography.bodyLarge.fontSize * 1.5
                                )
                            }
                        }

                        // Options
                        question.options.entries.sortedBy { it.key }.forEach { (key, value) ->
                            OptionItem(
                                optionKey = key,
                                optionText = value,
                                isSelected = state.selectedAnswer == key,
                                showResult = state.showResult,
                                isCorrect = question.correct == key,
                                isWrong = state.selectedAnswer == key && question.correct != key,
                                onClick = {
                                    if (state.selectedAnswer == null || state.mode == QuizMode.EXAM) {
                                        viewModel.selectAnswer(key)
                                    }
                                }
                            )
                        }

                        // Explanation (study mode)
                        if (state.showResult && question.explanation.isNotBlank()) {
                            Card(
                                shape = RoundedCornerShape(16.dp),
                                colors = CardDefaults.cardColors(
                                    containerColor = MaterialTheme.colorScheme.tertiaryContainer
                                )
                            ) {
                                Column(modifier = Modifier.padding(16.dp)) {
                                    Text(
                                        text = "Explicación",
                                        style = MaterialTheme.typography.labelMedium,
                                        color = MaterialTheme.colorScheme.onTertiaryContainer,
                                        fontWeight = FontWeight.SemiBold
                                    )
                                    Spacer(Modifier.height(4.dp))
                                    Text(
                                        text = question.explanation,
                                        style = MaterialTheme.typography.bodyMedium,
                                        color = MaterialTheme.colorScheme.onTertiaryContainer
                                    )
                                }
                            }
                        }
                    }
                }
            }
        }

        is QuizUiState.Finished -> {
            LaunchedEffect(state.result) {
                onFinished(state.result)
            }
        }
    }
}

@Composable
private fun OptionItem(
    optionKey: String,
    optionText: String,
    isSelected: Boolean,
    showResult: Boolean,
    isCorrect: Boolean,
    isWrong: Boolean,
    onClick: () -> Unit
) {
    val bgColor = when {
        showResult && isCorrect -> CorrectGreenLight
        showResult && isWrong -> WrongRedLight
        isSelected -> MaterialTheme.colorScheme.primaryContainer
        else -> MaterialTheme.colorScheme.surface
    }
    val borderColor = when {
        showResult && isCorrect -> CorrectGreen
        showResult && isWrong -> WrongRed
        isSelected -> MaterialTheme.colorScheme.primary
        else -> MaterialTheme.colorScheme.outlineVariant
    }
    val textColor = when {
        showResult && isCorrect -> CorrectGreen
        showResult && isWrong -> WrongRed
        isSelected -> MaterialTheme.colorScheme.primary
        else -> MaterialTheme.colorScheme.onSurface
    }

    Row(
        modifier = Modifier
            .fillMaxWidth()
            .clip(RoundedCornerShape(12.dp))
            .background(bgColor)
            .border(
                width = if (isSelected || (showResult && isCorrect)) 2.dp else 1.dp,
                color = borderColor,
                shape = RoundedCornerShape(12.dp)
            )
            .clickable(onClick = onClick)
            .padding(16.dp),
        verticalAlignment = Alignment.CenterVertically
    ) {
        Box(
            modifier = Modifier
                .size(32.dp)
                .clip(RoundedCornerShape(8.dp))
                .background(borderColor),
            contentAlignment = Alignment.Center
        ) {
            if (showResult && isCorrect) {
                Icon(Icons.Filled.Check, null, tint = Color.White, modifier = Modifier.size(18.dp))
            } else if (showResult && isWrong) {
                Icon(Icons.Filled.Close, null, tint = Color.White, modifier = Modifier.size(18.dp))
            } else {
                Text(
                    text = optionKey,
                    color = Color.White,
                    fontWeight = FontWeight.Bold,
                    style = MaterialTheme.typography.labelLarge
                )
            }
        }
        Spacer(modifier = Modifier.width(12.dp))
        Text(
            text = optionText,
            style = MaterialTheme.typography.bodyMedium,
            color = textColor,
            modifier = Modifier.weight(1f)
        )
    }
}
