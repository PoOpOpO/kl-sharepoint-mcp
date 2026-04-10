package com.osakidetza.ope.ui.navigation

import androidx.compose.runtime.Composable
import androidx.compose.runtime.remember
import androidx.lifecycle.viewmodel.compose.viewModel
import androidx.navigation.NavType
import androidx.navigation.compose.NavHost
import androidx.navigation.compose.composable
import androidx.navigation.compose.rememberNavController
import androidx.navigation.navArgument
import com.osakidetza.ope.data.model.QuizMode
import com.osakidetza.ope.ui.screens.CategoryDetailScreen
import com.osakidetza.ope.ui.screens.HomeScreen
import com.osakidetza.ope.ui.screens.QuizScreen
import com.osakidetza.ope.ui.screens.ResultScreen
import com.osakidetza.ope.ui.screens.StatsScreen
import com.osakidetza.ope.viewmodel.QuizViewModel

sealed class Screen(val route: String) {
    object Home : Screen("home")
    object CategoryDetail : Screen("category/{categoryId}") {
        fun createRoute(categoryId: String) = "category/$categoryId"
    }
    object Quiz : Screen("quiz/{categoryId}/{mode}/{count}") {
        fun createRoute(categoryId: String, mode: QuizMode, count: Int) =
            "quiz/$categoryId/${mode.name}/$count"
    }
    object Result : Screen("result")
    object Stats : Screen("stats")
}

@Composable
fun AppNavigation() {
    val navController = rememberNavController()

    NavHost(navController = navController, startDestination = Screen.Home.route) {
        composable(Screen.Home.route) {
            HomeScreen(
                onCategoryClick = { categoryId ->
                    navController.navigate(Screen.CategoryDetail.createRoute(categoryId))
                },
                onStatsClick = {
                    navController.navigate(Screen.Stats.route)
                }
            )
        }

        composable(
            route = Screen.CategoryDetail.route,
            arguments = listOf(navArgument("categoryId") { type = NavType.StringType })
        ) { backStackEntry ->
            val categoryId = backStackEntry.arguments?.getString("categoryId") ?: return@composable
            CategoryDetailScreen(
                categoryId = categoryId,
                onStartStudy = { count ->
                    navController.navigate(Screen.Quiz.createRoute(categoryId, QuizMode.STUDY, count))
                },
                onStartExam = { count ->
                    navController.navigate(Screen.Quiz.createRoute(categoryId, QuizMode.EXAM, count))
                },
                onBack = { navController.popBackStack() }
            )
        }

        composable(
            route = Screen.Quiz.route,
            arguments = listOf(
                navArgument("categoryId") { type = NavType.StringType },
                navArgument("mode") { type = NavType.StringType },
                navArgument("count") { type = NavType.IntType }
            )
        ) { backStackEntry ->
            val categoryId = backStackEntry.arguments?.getString("categoryId") ?: return@composable
            val modeStr = backStackEntry.arguments?.getString("mode") ?: "STUDY"
            val mode = QuizMode.valueOf(modeStr)
            val count = backStackEntry.arguments?.getInt("count") ?: 20
            QuizScreen(
                categoryId = categoryId,
                mode = mode,
                questionCount = count,
                onFinished = { _ ->
                    // Navigate without popping quiz entry so its ViewModel stays alive
                    navController.navigate(Screen.Result.route)
                },
                onBack = { navController.popBackStack() }
            )
        }

        composable(Screen.Result.route) { backStackEntry ->
            // Share the QuizViewModel from the quiz back stack entry so we see the Finished state
            val quizBackStackEntry = remember(backStackEntry) {
                navController.previousBackStackEntry
            }
            val sharedViewModel: QuizViewModel = viewModel(quizBackStackEntry!!)
            ResultScreen(
                viewModel = sharedViewModel,
                onHome = {
                    navController.navigate(Screen.Home.route) {
                        popUpTo(Screen.Home.route) { inclusive = true }
                    }
                },
                onRetry = {
                    // Pop result and quiz, returning to CategoryDetail
                    navController.popBackStack()
                    navController.popBackStack()
                }
            )
        }

        composable(Screen.Stats.route) {
            StatsScreen(onBack = { navController.popBackStack() })
        }
    }
}
