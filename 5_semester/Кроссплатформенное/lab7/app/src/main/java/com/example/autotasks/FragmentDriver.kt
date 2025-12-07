package com.example.autotasks

import android.os.Bundle
import android.util.Log
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import androidx.fragment.app.Fragment
import androidx.lifecycle.lifecycleScope
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import com.example.autotasks.database.Driver
import kotlinx.coroutines.launch

class FragmentDriver(
    private val drivers: MutableList<Driver>,
    private val loadDriverInfo: suspend (Driver) -> String,
    private val showDriverActionsDialog: (Driver) -> Unit,
    private val refreshDrivers: () -> Unit,
    private val getLastLoadTime: () -> Long
) : Fragment() {

    private var currentDriverIndex = 0
    private lateinit var adapter: DriversAdapter
    private lateinit var textResult: TextView
    private lateinit var textLastUpdate: TextView

    companion object {
        fun newInstance(
            drivers: MutableList<Driver>,
            loadDriverInfo: suspend (Driver) -> String,
            showDriverActionsDialog: (Driver) -> Unit,
            refreshDrivers: () -> Unit,
            getLastLoadTime: () -> Long
        ): FragmentDriver {
            return FragmentDriver(drivers, loadDriverInfo, showDriverActionsDialog, refreshDrivers, getLastLoadTime)
        }
    }

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View? {
        val view = inflater.inflate(R.layout.fragment_driver, container, false)

        val recyclerDrivers = view.findViewById<RecyclerView>(R.id.recyclerDrivers)
        textResult = view.findViewById<TextView>(R.id.textResult)
        textLastUpdate = view.findViewById<TextView>(R.id.textLastUpdate)

        recyclerDrivers.layoutManager = LinearLayoutManager(requireContext())
        adapter = DriversAdapter(
            drivers,
            onItemClick = { driver ->
                currentDriverIndex = drivers.indexOf(driver)
                lifecycleScope.launch {
                    textResult.text = loadDriverInfo(driver)
                    (activity as MainActivity).currentDriverIndex = currentDriverIndex
                }
            },
            onItemLongClick = { driver ->
                showDriverActionsDialog(driver)
            }
        )
        recyclerDrivers.adapter = adapter

        // Обновляем adapter сразу после создания, чтобы показать текущие данные
        adapter.notifyDataSetChanged()
        
        // Обновляем время последней загрузки
        val lastLoadTime = getLastLoadTime()
        if (lastLoadTime > 0) {
            val timeFormat = java.text.SimpleDateFormat("HH:mm:ss", java.util.Locale.getDefault())
            val timeString = timeFormat.format(java.util.Date(lastLoadTime))
            textLastUpdate.text = "Последняя загрузка: $timeString (${drivers.size} гонщиков)"
        } else {
            textLastUpdate.text = "Последняя загрузка: не загружено"
        }

        if (drivers.isNotEmpty()) {
            lifecycleScope.launch {
                textResult.text = loadDriverInfo(drivers[currentDriverIndex])
            }
        }

        return view
    }
    
    fun refreshDrivers() {
        if (!::adapter.isInitialized) {
            android.util.Log.w("FragmentDriver", "refreshDrivers: adapter не инициализирован")
            return
        }
        
        try {
            // Получаем актуальные данные из MainActivity
            val mainActivity = activity as? MainActivity
            val currentDrivers = if (mainActivity != null) {
                android.util.Log.d("FragmentDriver", "refreshDrivers: получение данных из MainActivity, размер: ${mainActivity.drivers.size}")
                mainActivity.drivers
            } else {
                android.util.Log.w("FragmentDriver", "refreshDrivers: MainActivity null, используем локальный список, размер: ${drivers.size}")
                drivers
            }
            
            android.util.Log.d("FragmentDriver", "refreshDrivers: обновление с ${currentDrivers.size} гонщиками")
            android.util.Log.d("FragmentDriver", "refreshDrivers: ID гонщиков: ${currentDrivers.map { it.id }}")
            
            if (currentDrivers.isEmpty()) {
                android.util.Log.w("FragmentDriver", "refreshDrivers: ВНИМАНИЕ - список пустой!")
            }
            
            // Обновляем adapter с актуальными данными
            // Важно: обновляем в главном потоке
            adapter.updateDrivers(currentDrivers.toList()) // Создаем копию списка
            
            // Принудительно обновляем RecyclerView
            view?.findViewById<androidx.recyclerview.widget.RecyclerView>(R.id.recyclerDrivers)?.adapter?.notifyDataSetChanged()
            
                    // Обновляем время последней загрузки
                    if (::textLastUpdate.isInitialized) {
                        val lastLoadTime = getLastLoadTime()
                        Log.d("FragmentDriver", "refreshDrivers: lastLoadTime = $lastLoadTime")
                        if (lastLoadTime > 0) {
                            val timeFormat = java.text.SimpleDateFormat("HH:mm:ss", java.util.Locale.getDefault())
                            val timeString = timeFormat.format(java.util.Date(lastLoadTime))
                            val updateText = "Последняя загрузка: $timeString (${currentDrivers.size} гонщиков)"
                            textLastUpdate.text = updateText
                            Log.d("FragmentDriver", "refreshDrivers: установлен текст: $updateText")
                        } else {
                            textLastUpdate.text = "Последняя загрузка: не загружено"
                            Log.w("FragmentDriver", "refreshDrivers: lastLoadTime = 0, время не установлено")
                        }
                    } else {
                        Log.w("FragmentDriver", "refreshDrivers: textLastUpdate не инициализирован")
                    }
            
            if (currentDrivers.isNotEmpty()) {
                // Проверяем и корректируем индекс
                if (currentDriverIndex >= currentDrivers.size) {
                    currentDriverIndex = 0
                }
                lifecycleScope.launch {
                    try {
                        textResult.text = loadDriverInfo(currentDrivers[currentDriverIndex])
                    } catch (e: Exception) {
                        e.printStackTrace()
                        textResult.text = "Ошибка загрузки информации"
                    }
                }
            } else {
                // Если список пустой, очищаем текст
                textResult.text = "Нет гонщиков"
            }
        } catch (e: Exception) {
            e.printStackTrace()
            android.util.Log.e("FragmentDriver", "Ошибка в refreshDrivers: ${e.message}")
        }
    }
}
