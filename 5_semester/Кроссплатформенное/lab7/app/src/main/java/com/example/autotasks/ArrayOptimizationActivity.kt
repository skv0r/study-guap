package com.example.autotasks

import android.os.Bundle
import android.widget.Button
import android.widget.ScrollView
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import kotlin.system.measureTimeMillis


class ArrayOptimizationActivity : AppCompatActivity() {

    private lateinit var textResults: TextView
    private lateinit var scrollView: ScrollView
    private lateinit var btnTest: Button
    
    // Размер тестового массива
    private val ARRAY_SIZE = 100_000
    
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_array_optimization)
        
        textResults = findViewById(R.id.textResults)
        scrollView = findViewById(R.id.scrollView)
        btnTest = findViewById(R.id.btnTestOptimization)
        
        btnTest.setOnClickListener {
            runOptimizationTests()
        }
        
        appendText("=== Оптимизация перебора массивов в Kotlin ===\n")
        appendText("Размер тестового массива: $ARRAY_SIZE элементов\n")
        appendText("Нажмите кнопку для запуска тестов\n\n")
    }
    
    private fun runOptimizationTests() {
        lifecycleScope.launch {
            btnTest.isEnabled = false
            textResults.text = ""
            
            appendText("=== ЗАПУСК ТЕСТОВ ===\n\n")
            
            // Создаем исходный массив
            appendText("Создание исходного массива...\n")
            val sourceArray = createSourceArray(ARRAY_SIZE)
            appendText("✓ Исходный массив создан\n")
            appendText("Первые 10 элементов: ${sourceArray.take(10)}\n")
            appendText("Последние 10 элементов: ${sourceArray.takeLast(10)}\n\n")
            
            // Тест 1: Обычный цикл for
            test1_SimpleForLoop(sourceArray)
            
            // Тест 2: forEach
            test2_ForEach(sourceArray)
            
            // Тест 3: map (функциональный подход)
            test3_Map(sourceArray)
            
            // Тест 4: filter + map (цепочка операций)
            test4_FilterMap(sourceArray)
            
            // Тест 5: sequence (ленивые вычисления)
            test5_Sequence(sourceArray)
            
            // Тест 6: asSequence + filter + map (оптимальный)
            test6_SequenceOptimized(sourceArray)
            
            // Тест 7: parallelStream (параллельная обработка)
            test7_ParallelProcessing(sourceArray)
            
            // Тест 8: IntArray vs List (примитивы)
            test8_IntArrayVsList(ARRAY_SIZE)
            
            // Итоговая сводка
            appendText("\n=== ИТОГОВАЯ СВОДКА ===\n")
            appendText("Самые быстрые методы:\n")
            appendText("1. IntArray с индексным доступом\n")
            appendText("2. Sequence для цепочек операций\n")
            appendText("3. forEach для простых итераций\n\n")
            appendText("Рекомендации:\n")
            appendText("- Используйте примитивные массивы (IntArray) вместо List<Int>\n")
            appendText("- Для цепочек filter/map используйте asSequence()\n")
            appendText("- Избегайте создания промежуточных коллекций\n")
            appendText("- Для больших данных используйте sequence\n\n")
            
            btnTest.isEnabled = true
            scrollView.post { scrollView.fullScroll(ScrollView.FOCUS_DOWN) }
        }
    }
    
    // Создание исходного массива
    private suspend fun createSourceArray(size: Int): List<Int> = withContext(Dispatchers.Default) {
        List(size) { it + 1 }
    }
    
    // Тест 1: Обычный цикл for с индексом
    private suspend fun test1_SimpleForLoop(source: List<Int>) = withContext(Dispatchers.Default) {
        appendText("--- Тест 1: Обычный цикл for ---\n")
        
        val result = mutableListOf<Int>()
        var intermediate = 0
        
        val time = measureTimeMillis {
            for (i in source.indices) {
                val value = source[i]
                if (value % 2 == 0) {
                    result.add(value * 2)
                }
                if (i == 1000) intermediate = result.size
            }
        }
        
        appendText("Время выполнения: $time мс\n")
        appendText("Промежуточный размер (после 1000 эл.): $intermediate\n")
        appendText("Итоговый размер: ${result.size}\n")
        appendText("Первые 10 результатов: ${result.take(10)}\n\n")
    }
    
    // Тест 2: forEach
    private suspend fun test2_ForEach(source: List<Int>) = withContext(Dispatchers.Default) {
        appendText("--- Тест 2: forEach ---\n")
        
        val result = mutableListOf<Int>()
        var count = 0
        var intermediate = 0
        
        val time = measureTimeMillis {
            source.forEach { value ->
                if (value % 2 == 0) {
                    result.add(value * 2)
                }
                count++
                if (count == 1000) intermediate = result.size
            }
        }
        
        appendText("Время выполнения: $time мс\n")
        appendText("Промежуточный размер (после 1000 эл.): $intermediate\n")
        appendText("Итоговый размер: ${result.size}\n")
        appendText("Первые 10 результатов: ${result.take(10)}\n\n")
    }
    
    // Тест 3: map (функциональный подход)
    private suspend fun test3_Map(source: List<Int>) = withContext(Dispatchers.Default) {
        appendText("--- Тест 3: map (функциональный подход) ---\n")
        
        lateinit var result: List<Int>
        
        val time = measureTimeMillis {
            result = source.map { it * 2 }
        }
        
        appendText("Время выполнения: $time мс\n")
        appendText("Итоговый размер: ${result.size}\n")
        appendText("Первые 10 результатов: ${result.take(10)}\n\n")
    }
    
    // Тест 4: filter + map (цепочка операций)
    private suspend fun test4_FilterMap(source: List<Int>) = withContext(Dispatchers.Default) {
        appendText("--- Тест 4: filter + map (цепочка) ---\n")
        
        lateinit var intermediate: List<Int>
        lateinit var result: List<Int>
        
        val time = measureTimeMillis {
            intermediate = source.filter { it % 2 == 0 }
            result = intermediate.map { it * 2 }
        }
        
        appendText("Время выполнения: $time мс\n")
        appendText("Промежуточный размер (после filter): ${intermediate.size}\n")
        appendText("Итоговый размер: ${result.size}\n")
        appendText("Первые 10 результатов: ${result.take(10)}\n")
        appendText("⚠ Создаются промежуточные коллекции!\n\n")
    }
    
    // Тест 5: sequence (ленивые вычисления)
    private suspend fun test5_Sequence(source: List<Int>) = withContext(Dispatchers.Default) {
        appendText("--- Тест 5: sequence (ленивые вычисления) ---\n")
        
        lateinit var result: List<Int>
        
        val time = measureTimeMillis {
            result = source.asSequence()
                .filter { it % 2 == 0 }
                .map { it * 2 }
                .toList()
        }
        
        appendText("Время выполнения: $time мс\n")
        appendText("Итоговый размер: ${result.size}\n")
        appendText("Первые 10 результатов: ${result.take(10)}\n")
        appendText("✓ Без промежуточных коллекций!\n\n")
    }
    
    // Тест 6: asSequence + filter + map (оптимизированный)
    private suspend fun test6_SequenceOptimized(source: List<Int>) = withContext(Dispatchers.Default) {
        appendText("--- Тест 6: Sequence оптимизированный ---\n")
        
        lateinit var result: List<Int>
        
        val time = measureTimeMillis {
            result = source.asSequence()
                .filter { it % 2 == 0 }
                .map { it * 2 }
                .take(10000) // Берем только первые 10000
                .toList()
        }
        
        appendText("Время выполнения: $time мс\n")
        appendText("Итоговый размер: ${result.size}\n")
        appendText("Первые 10 результатов: ${result.take(10)}\n")
        appendText("✓ Ленивая оценка + ранний выход!\n\n")
    }
    
    // Тест 7: Параллельная обработка (имитация)
    private suspend fun test7_ParallelProcessing(source: List<Int>) = withContext(Dispatchers.Default) {
        appendText("--- Тест 7: Chunked обработка (партиями) ---\n")
        
        val results = mutableListOf<Int>()
        
        val time = measureTimeMillis {
            source.chunked(10000).forEach { chunk ->
                val chunkResult = chunk
                    .filter { it % 2 == 0 }
                    .map { it * 2 }
                results.addAll(chunkResult)
            }
        }
        
        appendText("Время выполнения: $time мс\n")
        appendText("Итоговый размер: ${results.size}\n")
        appendText("Первые 10 результатов: ${results.take(10)}\n")
        appendText("✓ Обработка партиями по 10000 элементов\n\n")
    }
    
    // Тест 8: IntArray vs List<Int>
    private suspend fun test8_IntArrayVsList(size: Int) = withContext(Dispatchers.Default) {
        appendText("--- Тест 8: IntArray vs List<Int> ---\n")
        
        // List<Int>
        val listTime = measureTimeMillis {
            val list = List(size) { it }
            var sum = 0
            for (i in list.indices) {
                sum += list[i]
            }
        }
        
        // IntArray
        val arrayTime = measureTimeMillis {
            val array = IntArray(size) { it }
            var sum = 0
            for (i in array.indices) {
                sum += array[i]
            }
        }
        
        appendText("List<Int> время: $listTime мс\n")
        appendText("IntArray время: $arrayTime мс\n")
        appendText("Разница: ${listTime - arrayTime} мс\n")
        appendText("✓ IntArray быстрее на ${((listTime - arrayTime) * 100 / listTime)}%\n\n")
    }
    
    private fun appendText(text: String) {
        runOnUiThread {
            textResults.append(text)
        }
    }
}

