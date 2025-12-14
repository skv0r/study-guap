package com.example.autotasks

import android.Manifest
import android.content.pm.PackageManager
import android.os.Bundle
import android.widget.Button
import android.widget.Toast
import androidx.appcompat.app.AlertDialog
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import androidx.lifecycle.lifecycleScope
import com.google.android.material.floatingactionbutton.FloatingActionButton
import com.yandex.mapkit.Animation
import com.yandex.mapkit.MapKit
import com.yandex.mapkit.MapKitFactory
import com.yandex.mapkit.geometry.Point
import com.yandex.mapkit.map.CameraPosition
import com.yandex.mapkit.map.InputListener
import com.yandex.mapkit.map.Map
import com.yandex.mapkit.map.MapObjectTapListener
import com.yandex.mapkit.mapview.MapView
import com.yandex.runtime.image.ImageProvider
import com.example.autotasks.database.AppDatabase
import com.example.autotasks.database.MapMarker
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import android.widget.EditText
import android.view.ViewGroup
import android.widget.LinearLayout

class MapActivity : AppCompatActivity() {

    private lateinit var mapView: MapView
    private lateinit var database: AppDatabase
    private val markers = mutableListOf<com.yandex.mapkit.map.PlacemarkMapObject>()
    
    // Координаты центра Москвы (Красная площадь)
    private val MOSCOW_CENTER = Point(55.753544, 37.621211)
    
    // Известные места в Москве для примера
    private val FAMOUS_PLACES = listOf(
        Triple("Красная площадь", 55.753544, 37.621211),
        Triple("Кремль", 55.752004, 37.617734),
        Triple("Храм Василия Блаженного", 55.752516, 37.623147),
        Triple("ГУМ", 55.754491, 37.621481),
        Triple("Большой театр", 55.760376, 37.618423),
        Triple("МГУ", 55.703370, 37.530663),
        Triple("Останкинская башня", 55.819694, 37.611916),
        Triple("Парк Горького", 55.730899, 37.601179),
        Triple("ВДНХ", 55.823921, 37.636856),
        Triple("Третьяковская галерея", 55.741425, 37.620498)
    )
    
    companion object {
        private const val PERMISSIONS_REQUEST_CODE = 100
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        
        // Инициализация MapKit должна быть ПЕРЕД setContentView
        MapKitFactory.setApiKey("8975e7ac-e7be-47ef-8f5d-4a0a08c3af04")
        MapKitFactory.initialize(this)
        
        setContentView(R.layout.activity_map)
        
        database = AppDatabase.getDatabase(this)
        
        // Инициализация карты
        mapView = findViewById(R.id.mapview)
        
        // Устанавливаем начальную позицию камеры на Москву
        mapView.map.move(
            CameraPosition(MOSCOW_CENTER, 11.0f, 0.0f, 0.0f),
            Animation(Animation.Type.SMOOTH, 0f),
            null
        )
        
        // Проверяем разрешения
        checkPermissions()
        
        // Настраиваем кнопки
        setupButtons()
        
        // Загружаем сохраненные метки
        loadMarkersFromDatabase()
        
        // Добавляем слушатель долгого нажатия на карту
        setupMapListener()
    }
    
    private fun checkPermissions() {
        if (ContextCompat.checkSelfPermission(
                this,
                Manifest.permission.ACCESS_FINE_LOCATION
            ) != PackageManager.PERMISSION_GRANTED
        ) {
            ActivityCompat.requestPermissions(
                this,
                arrayOf(
                    Manifest.permission.ACCESS_FINE_LOCATION,
                    Manifest.permission.ACCESS_COARSE_LOCATION
                ),
                PERMISSIONS_REQUEST_CODE
            )
        }
    }
    
    override fun onRequestPermissionsResult(
        requestCode: Int,
        permissions: Array<out String>,
        grantResults: IntArray
    ) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        if (requestCode == PERMISSIONS_REQUEST_CODE) {
            if (grantResults.isNotEmpty() && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                Toast.makeText(this, "Разрешения получены", Toast.LENGTH_SHORT).show()
            } else {
                Toast.makeText(this, "Разрешения не получены", Toast.LENGTH_SHORT).show()
            }
        }
    }
    
    private fun setupButtons() {
        val btnAddMarker = findViewById<Button>(R.id.btnAddMarker)
        val btnClearMarkers = findViewById<Button>(R.id.btnClearMarkers)
        val fabMyLocation = findViewById<FloatingActionButton>(R.id.fabMyLocation)
        
        btnAddMarker.setOnClickListener {
            showAddMarkerDialog()
        }
        
        btnClearMarkers.setOnClickListener {
            confirmClearMarkers()
        }
        
        fabMyLocation.setOnClickListener {
            // Возврат к центру Москвы
            mapView.map.move(
                CameraPosition(MOSCOW_CENTER, 11.0f, 0.0f, 0.0f),
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
    
    private fun showAddMarkerDialog() {
        val places = FAMOUS_PLACES.map { it.first }.toTypedArray()
        
        AlertDialog.Builder(this)
            .setTitle("Выберите место в Москве")
            .setItems(places) { _, which ->
                val place = FAMOUS_PLACES[which]
                showMarkerDetailsDialog(place.first, Point(place.second, place.third))
            }
            .setNegativeButton("Отмена", null)
            .show()
    }
    
    private fun showAddMarkerDialogAtPoint(point: Point) {
        val layout = LinearLayout(this).apply {
            orientation = LinearLayout.VERTICAL
            setPadding(60, 40, 60, 20)
        }
        
        val titleInput = EditText(this).apply {
            hint = "Название места"
        }
        
        val descriptionInput = EditText(this).apply {
            hint = "Описание"
        }
        
        layout.addView(titleInput)
        layout.addView(descriptionInput)
        
        AlertDialog.Builder(this)
            .setTitle("Добавить метку")
            .setView(layout)
            .setPositiveButton("Добавить") { _, _ ->
                val title = titleInput.text.toString().trim()
                val description = descriptionInput.text.toString().trim()
                
                if (title.isEmpty()) {
                    Toast.makeText(this, "Введите название", Toast.LENGTH_SHORT).show()
                    return@setPositiveButton
                }
                
                addMarker(title, description.ifEmpty { "Нет описания" }, point)
            }
            .setNegativeButton("Отмена", null)
            .show()
    }
    
    private fun showMarkerDetailsDialog(name: String, point: Point) {
        val layout = LinearLayout(this).apply {
            orientation = LinearLayout.VERTICAL
            setPadding(60, 40, 60, 20)
        }
        
        val descriptionInput = EditText(this).apply {
            hint = "Описание места"
        }
        
        layout.addView(descriptionInput)
        
        AlertDialog.Builder(this)
            .setTitle(name)
            .setView(layout)
            .setPositiveButton("Добавить метку") { _, _ ->
                val description = descriptionInput.text.toString().trim()
                addMarker(name, description.ifEmpty { "Известное место в Москве" }, point)
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
                
                Toast.makeText(this@MapActivity, "Метка добавлена: $title", Toast.LENGTH_SHORT).show()
                
            } catch (e: Exception) {
                e.printStackTrace()
                Toast.makeText(this@MapActivity, "Ошибка: ${e.message}", Toast.LENGTH_SHORT).show()
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
        
        AlertDialog.Builder(this)
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
                
                Toast.makeText(this@MapActivity, "Метка удалена", Toast.LENGTH_SHORT).show()
                
            } catch (e: Exception) {
                e.printStackTrace()
                Toast.makeText(this@MapActivity, "Ошибка удаления: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }
    
    private fun confirmClearMarkers() {
        if (markers.isEmpty()) {
            Toast.makeText(this, "Нет меток для удаления", Toast.LENGTH_SHORT).show()
            return
        }
        
        AlertDialog.Builder(this)
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
                
                Toast.makeText(this@MapActivity, "Все метки удалены", Toast.LENGTH_SHORT).show()
                
            } catch (e: Exception) {
                e.printStackTrace()
                Toast.makeText(this@MapActivity, "Ошибка: ${e.message}", Toast.LENGTH_SHORT).show()
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
                    Toast.makeText(this@MapActivity, "Загружено меток: ${savedMarkers.size}", Toast.LENGTH_SHORT).show()
                }
                
            } catch (e: Exception) {
                e.printStackTrace()
                Toast.makeText(this@MapActivity, "Ошибка загрузки меток: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }
    
    override fun onStart() {
        super.onStart()
        MapKitFactory.getInstance().onStart()
        mapView.onStart()
    }
    
    override fun onStop() {
        mapView.onStop()
        MapKitFactory.getInstance().onStop()
        super.onStop()
    }
}

