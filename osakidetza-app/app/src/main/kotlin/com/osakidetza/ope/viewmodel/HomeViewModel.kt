package com.osakidetza.ope.viewmodel

import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.osakidetza.ope.data.QuestionRepository
import com.osakidetza.ope.data.model.Category
import com.osakidetza.ope.data.model.CategoryStats
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.SharingStarted
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.flow.combine
import kotlinx.coroutines.flow.stateIn
import kotlinx.coroutines.launch

class HomeViewModel(application: Application) : AndroidViewModel(application) {

    private val repository = QuestionRepository(application)

    val categories: List<Category> = repository.categories

    private val _searchQuery = MutableStateFlow("")
    val searchQuery: StateFlow<String> = _searchQuery.asStateFlow()

    val filteredCategories: StateFlow<List<Category>> = _searchQuery
        .combine(MutableStateFlow(categories)) { query, cats ->
            if (query.isBlank()) cats
            else cats.filter {
                it.name.contains(query, ignoreCase = true) ||
                it.description.contains(query, ignoreCase = true)
            }
        }
        .stateIn(viewModelScope, SharingStarted.Lazily, categories)

    // Question counts loaded asynchronously on IO dispatcher to avoid main-thread I/O
    private val _questionCounts = MutableStateFlow<Map<String, Int>>(emptyMap())
    val questionCounts: StateFlow<Map<String, Int>> = _questionCounts.asStateFlow()

    init {
        viewModelScope.launch {
            val counts = categories.associate { cat ->
                cat.id to repository.loadQuestions(cat.id).size
            }
            _questionCounts.value = counts
        }
    }

    fun onSearchQueryChange(query: String) {
        _searchQuery.value = query
    }

    fun getCategoryStats(categoryId: String): StateFlow<CategoryStats> {
        return repository.getCategoryStats(categoryId)
            .stateIn(viewModelScope, SharingStarted.Lazily, CategoryStats(categoryId))
    }
}
