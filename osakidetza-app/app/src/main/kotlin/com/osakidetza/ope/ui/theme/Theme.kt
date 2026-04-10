package com.osakidetza.ope.ui.theme

import android.app.Activity
import androidx.compose.foundation.isSystemInDarkTheme
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.darkColorScheme
import androidx.compose.material3.lightColorScheme
import androidx.compose.runtime.Composable
import androidx.compose.runtime.SideEffect
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.toArgb
import androidx.compose.ui.platform.LocalView
import androidx.core.view.WindowCompat

private val LightColorScheme = lightColorScheme(
    primary = OsakidetзaRed,
    onPrimary = Color.White,
    primaryContainer = Color(0xFFFFDAD6),
    onPrimaryContainer = OsakidetзaRedDark,
    secondary = OsakidetзaBlue,
    onSecondary = Color.White,
    secondaryContainer = Color(0xFFD4E3FF),
    onSecondaryContainer = OsakidetзaBlue,
    tertiary = OsakidetзaGreen,
    onTertiary = Color.White,
    tertiaryContainer = Color(0xFFD7F5E4),
    onTertiaryContainer = OsakidetзaGreen,
    background = Color(0xFFFFFBFF),
    onBackground = Color(0xFF1C1B1F),
    surface = Color(0xFFFFFBFF),
    onSurface = Color(0xFF1C1B1F),
    surfaceVariant = Color(0xFFF4F0F4),
    onSurfaceVariant = Color(0xFF49454F),
    error = WrongRed,
    onError = Color.White
)

private val DarkColorScheme = darkColorScheme(
    primary = Color(0xFFFFB3AD),
    onPrimary = Color(0xFF680003),
    primaryContainer = OsakidetзaRedDark,
    onPrimaryContainer = Color(0xFFFFDAD6),
    secondary = Color(0xFFAAC7FF),
    onSecondary = Color(0xFF002E6F),
    secondaryContainer = OsakidetзaBlue,
    onSecondaryContainer = Color(0xFFD4E3FF),
    tertiary = Color(0xFF6EDB95),
    onTertiary = Color(0xFF003919),
    tertiaryContainer = OsakidetзaGreen,
    onTertiaryContainer = Color(0xFFD7F5E4),
    background = Color(0xFF1C1B1F),
    onBackground = Color(0xFFE6E1E5),
    surface = Color(0xFF1C1B1F),
    onSurface = Color(0xFFE6E1E5),
    error = Color(0xFFFFB4AB),
    onError = Color(0xFF690005)
)

@Composable
fun OsakidetзaOPETheme(
    darkTheme: Boolean = isSystemInDarkTheme(),
    content: @Composable () -> Unit
) {
    val colorScheme = if (darkTheme) DarkColorScheme else LightColorScheme

    val view = LocalView.current
    if (!view.isInEditMode) {
        SideEffect {
            val window = (view.context as Activity).window
            window.statusBarColor = colorScheme.primary.toArgb()
            WindowCompat.getInsetsController(window, view).isAppearanceLightStatusBars = false
        }
    }

    MaterialTheme(
        colorScheme = colorScheme,
        content = content
    )
}
