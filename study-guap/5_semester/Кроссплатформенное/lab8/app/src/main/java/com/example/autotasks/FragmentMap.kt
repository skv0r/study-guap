package com.example.autotasks

import android.Manifest
import android.content.Context
import android.content.pm.PackageManager
import android.net.ConnectivityManager
import android.net.NetworkCapabilities
import android.os.Build
import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.Button
import android.widget.EditText
import android.widget.LinearLayout
import android.widget.Toast
import androidx.appcompat.app.AlertDialog
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import androidx.fragment.app.Fragment
import androidx.lifecycle.lifecycleScope
import com.example.autotasks.database.AppDatabase
import com.example.autotasks.database.MapMarker
import com.google.android.material.floatingactionbutton.FloatingActionButton
import com.yandex.mapkit.Animation
import com.yandex.mapkit.MapKitFactory
import com.yandex.mapkit.geometry.Point
import com.yandex.mapkit.map.CameraPosition
import com.yandex.mapkit.map.InputListener
import com.yandex.mapkit.map.Map
import com.yandex.mapkit.map.MapObjectTapListener
import com.yandex.mapkit.mapview.MapView
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext

class FragmentMap : Fragment() {

    private lateinit var mapView: MapView
    private lateinit var database: AppDatabase
    private val markers = mutableListOf<com.yandex.mapkit.map.PlacemarkMapObject>()
    
    // Координаты центра Москвы (Красная площадь)
    private val MOSCOW_CENTER = Point(55.753544, 37.621211)
    
    companion object {
        private const val PERMISSIONS_REQUEST_CODE = 100
        
        fun newInstance() = FragmentMap()
    }
    
    private fun isInternetAvailable(): Boolean {
        val connectivityManager = requireContext().getSystemService(Context.CONNECTIVITY_SERVICE) as ConnectivityManager
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            val network = connectivityManager.activeNetwork ?: return false
            val capabilities = connectivityManager.getNetworkCapabilities(network) ?: return false
            return capabilities.hasCapability(NetworkCapabilities.NET_CAPABILITY_INTERNET)
        } else {
            @Suppress("DEPRECATION")
            val networkInfo = connectivityManager.activeNetworkInfo
            @Suppress("DEPRECATION")
            return networkInfo?.isConnected == true
        }
    }

    override fun onCreateView(
        inflater: LayoutInflater,
        container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View? {
        return inflater.inflate(R.layout.fragment_map, container, false)
    }

    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
        super.onViewCreated(view, savedInstanceState)
        
        database = AppDatabase.getDatabase(requireContext())
        
        // Проверка интернета
        if (!isInternetAvailable()) {
            Toast.makeText(requireContext(), "Нет подключения к интернету. Карта может не загрузиться.", Toast.LENGTH_LONG).show()
            android.util.Log.w("FragmentMap", "Нет подключения к интернету!")
        } else {
            android.util.Log.d("FragmentMap", "Интернет доступен")
        }
        
        // Инициализация карты
        mapView = view.findViewById(R.id.mapview)
        
        android.util.Log.d("FragmentMap", "MapView инициализирована, начинаем настройку карты")
        
        // ВАЖНО: запускаем MapView (MapKit уже запущен в MainActivity.onStart())
        try {
            mapView.onStart()
            android.util.Log.d("FragmentMap", "MapView запущена")
        } catch (e: Exception) {
            android.util.Log.e("FragmentMap", "Ошибка запуска MapView: ${e.message}")
            e.printStackTrace()
            Toast.makeText(requireContext(), "Ошибка запуска карты: ${e.message}", Toast.LENGTH_LONG).show()
        }
        
        // Настраиваем кнопки сразу
        setupButtons(view)
        
        // Небольшая задержка для инициализации карты и загрузки тайлов
        view.postDelayed({
            try {
                // Устанавливаем начальную позицию камеры на Москву с увеличенным зумом
                mapView.map.move(
                    CameraPosition(MOSCOW_CENTER, 12.0f, 0.0f, 0.0f),
                    Animation(Animation.Type.SMOOTH, 1.0f),
                    null
                )
                
                android.util.Log.d("FragmentMap", "Камера установлена на Москву (зум 12)")
                
                // Проверяем разрешения
                checkPermissions()
                
                // Загружаем сохраненные метки
                loadMarkersFromDatabase()
                
                // Добавляем слушатель долгого нажатия на карту
                setupMapListener()
                
            } catch (e: Exception) {
                android.util.Log.e("FragmentMap", "Ошибка настройки карты: ${e.message}")
                e.printStackTrace()
            }
        }, 300)
    }
    
    private fun checkPermissions() {
        if (ContextCompat.checkSelfPermission(
                requireContext(),
                Manifest.permission.ACCESS_FINE_LOCATION
            ) != PackageManager.PERMISSION_GRANTED
        ) {
            ActivityCompat.requestPermissions(
                requireActivity(),
                arrayOf(
                    Manifest.permission.ACCESS_FINE_LOCATION,
                    Manifest.permission.ACCESS_COARSE_LOCATION
                ),
                PERMISSIONS_REQUEST_CODE
            )
        }
    }
    
    private fun setupButtons(view: View) {
        val btnClearMarkers = view.findViewById<Button>(R.id.btnClearMarkers)
        val fabMyLocation = view.findViewById<FloatingActionButton>(R.id.fabMyLocation)
        
        btnClearMarkers.setOnClickListener {
            confirmClearMarkers()
        }
        
        fabMyLocation.setOnClickListener {
            // Возврат к центру Москвы
            mapView.map.move(
                CameraPosition(MOSCOW_CENTER, 12.0f, 0.0f, 0.0f),
                Animation(Animation.Type.SMOOTH, 1.0f),
                null
            )
        }
    }
    
    private fun setupMapListener() {
        mapView.map.addInputListener(object : InputListener {
            override fun onMapTap(map: Map, point: Point) {
                // Короткое нажатие - ничего не делаем
            }
            
            override fun onMapLongTap(map: Map, point: Point) {
                // Долгое нажатие - добавляем метку
                showAddMarkerDialogAtPoint(point)
            }
        })
    }
    
    private fun showAddMarkerDialogAtPoint(point: Point) {
        val layout = LinearLayout(requireContext()).apply {
            orientation = LinearLayout.VERTICAL
            setPadding(60, 40, 60, 20)
        }
        
        // Текст с координатами
        val coordsText = android.widget.TextView(requireContext()).apply {
            text = "Координаты:\nШирота: ${String.format("%.6f", point.latitude)}\nДолгота: ${String.format("%.6f", point.longitude)}"
            setPadding(0, 0, 0, 20)
            textSize = 12f
        }
        
        val titleInput = EditText(requireContext()).apply {
            hint = "Название места"
        }
        
        val descriptionInput = EditText(requireContext()).apply {
            hint = "Описание (необязательно)"
        }
        
        layout.addView(coordsText)
        layout.addView(titleInput)
        layout.addView(descriptionInput)
        
        AlertDialog.Builder(requireContext())
            .setTitle("Добавить метку на карту")
            .setMessage("Долгое нажатие на карту позволяет добавить метку")
            .setView(layout)
            .setPositiveButton("Добавить") { _, _ ->
                val title = titleInput.text.toString().trim()
                val description = descriptionInput.text.toString().trim()
                
                if (title.isEmpty()) {
                    Toast.makeText(requireContext(), "Введите название места", Toast.LENGTH_SHORT).show()
                    return@setPositiveButton
                }
                
                addMarker(title, description.ifEmpty { "Нет описания" }, point)
            }
            .setNegativeButton("Отмена", null)
            .show()
    }
    
    private fun addMarker(title: String, description: String, point: Point) {
        lifecycleScope.launch {
            try {
                // Сохраняем в БД
                val marker = MapMarker(
                    title = title,
                    description = description,
                    latitude = point.latitude,
                    longitude = point.longitude
                )
                
                val markerId = withContext(Dispatchers.IO) {
                    database.mapMarkerDao().insertMarker(marker)
                }
                
                // Добавляем метку на карту
                val mapMarker = mapView.map.mapObjects.addPlacemark(point)
                mapMarker.setText(title)
                
                // Добавляем слушатель нажатий на метку
                mapMarker.addTapListener(MapObjectTapListener { mapObject, _ ->
                    showMarkerInfo(marker.copy(id = markerId))
                    true
                })
                
                markers.add(mapMarker)
                
                Toast.makeText(requireContext(), "Метка добавлена: $title", Toast.LENGTH_SHORT).show()
                
            } catch (e: Exception) {
                e.printStackTrace()
                Toast.makeText(requireContext(), "Ошибка: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }
    
    private fun showMarkerInfo(marker: MapMarker) {
        val message = """
            Название: ${marker.title}
            
            Описание: ${marker.description}
            
            Координаты:
            Широта: ${String.format("%.6f", marker.latitude)}
            Долгота: ${String.format("%.6f", marker.longitude)}
        """.trimIndent()
        
        AlertDialog.Builder(requireContext())
            .setTitle("Информация о месте")
            .setMessage(message)
            .setPositiveButton("Удалить") { _, _ ->
                deleteMarker(marker)
            }
            .setNegativeButton("Закрыть", null)
            .show()
    }
    
    private fun deleteMarker(marker: MapMarker) {
        lifecycleScope.launch {
            try {
                withContext(Dispatchers.IO) {
                    database.mapMarkerDao().deleteMarker(marker)
                }
                
                // Перезагружаем метки
                loadMarkersFromDatabase()
                
                Toast.makeText(requireContext(), "Метка удалена", Toast.LENGTH_SHORT).show()
                
            } catch (e: Exception) {
                e.printStackTrace()
                Toast.makeText(requireContext(), "Ошибка удаления: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }
    
    private fun confirmClearMarkers() {
        if (markers.isEmpty()) {
            Toast.makeText(requireContext(), "Нет меток для удаления", Toast.LENGTH_SHORT).show()
            return
        }
        
        AlertDialog.Builder(requireContext())
            .setTitle("Очистить все метки?")
            .setMessage("Вы уверены, что хотите удалить все метки с карты?")
            .setPositiveButton("Да") { _, _ ->
                clearAllMarkers()
            }
            .setNegativeButton("Нет", null)
            .show()
    }
    
    private fun clearAllMarkers() {
        lifecycleScope.launch {
            try {
                withContext(Dispatchers.IO) {
                    database.mapMarkerDao().deleteAllMarkers()
                }
                
                // Очищаем метки с карты
                mapView.map.mapObjects.clear()
                markers.clear()
                
                Toast.makeText(requireContext(), "Все метки удалены", Toast.LENGTH_SHORT).show()
                
            } catch (e: Exception) {
                e.printStackTrace()
                Toast.makeText(requireContext(), "Ошибка: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }
    
    private fun loadMarkersFromDatabase() {
        lifecycleScope.launch {
            try {
                val savedMarkers = withContext(Dispatchers.IO) {
                    database.mapMarkerDao().getAllMarkers()
                }
                
                // Очищаем текущие метки
                mapView.map.mapObjects.clear()
                markers.clear()
                
                // Добавляем сохраненные метки
                savedMarkers.forEach { marker ->
                    val point = Point(marker.latitude, marker.longitude)
                    val mapMarker = mapView.map.mapObjects.addPlacemark(point)
                    mapMarker.setText(marker.title)
                    
                    mapMarker.addTapListener(MapObjectTapListener { _, _ ->
                        showMarkerInfo(marker)
                        true
                    })
                    
                    markers.add(mapMarker)
                }
                
                if (savedMarkers.isNotEmpty()) {
                    Toast.makeText(requireContext(), "Загружено меток: ${savedMarkers.size}", Toast.LENGTH_SHORT).show()
                }
                
            } catch (e: Exception) {
                e.printStackTrace()
                Toast.makeText(requireContext(), "Ошибка загрузки меток: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }
    
    override fun onStart() {
        super.onStart()
        // onStart уже вызывается в onViewCreated, здесь не нужно
    }
    
    override fun onStop() {
        super.onStop()
        try {
            if (::mapView.isInitialized) {
                mapView.onStop()
                android.util.Log.d("FragmentMap", "MapView остановлена")
            }
        } catch (e: Exception) {
            android.util.Log.e("FragmentMap", "Ошибка остановки MapView: ${e.message}")
            e.printStackTrace()
        }
    }
    
    override fun onDestroyView() {
        super.onDestroyView()
        try {
            if (::mapView.isInitialized) {
                mapView.onStop()
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }
}

