package com.example.autotasks

import android.os.Bundle
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import androidx.fragment.app.commit
import com.google.android.material.bottomnavigation.BottomNavigationView
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.json.JSONArray
import java.net.HttpURLConnection
import java.net.URL
import android.net.ConnectivityManager
import android.net.NetworkInfo
import android.widget.Toast

@Suppress("DEPRECATION")
class MainActivity : AppCompatActivity() {

    val drivers = mutableListOf<Pair<String, Int>>() // Pair<fullName, driver_number>
    var currentDriverIndex = 0

    private lateinit var fragmentDriver: FragmentDriver

    private fun isNetworkAvailable(): Boolean {
        val connectivityManager = getSystemService(CONNECTIVITY_SERVICE) as ConnectivityManager
        val networkInfo: NetworkInfo? = connectivityManager.activeNetworkInfo
        return networkInfo?.isConnected == true
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        if (!isNetworkAvailable()) {
            Toast.makeText(this, "Нет интернета", Toast.LENGTH_SHORT).show()
        }

        val bottomNav = findViewById<BottomNavigationView>(R.id.bottomNavigation)

        lifecycleScope.launch {
            loadDriversList()

            // Создаём фрагмент гонщиков
            fragmentDriver = FragmentDriver.newInstance(drivers, ::loadDriverInfo)
            supportFragmentManager.commit {
                replace(R.id.fragmentContainer, fragmentDriver)
            }
        }

        bottomNav.setOnItemSelectedListener { item ->
            when (item.itemId) {
                R.id.menu_driver -> {
                    // Показываем фрагмент гонщиков
                    supportFragmentManager.commit {
                        replace(R.id.fragmentContainer, fragmentDriver)
                    }
                    true
                }
                R.id.menu_car -> {
                    // Показываем фрагмент авто текущего гонщика
                    if (drivers.isNotEmpty()) {
                        val driverName = drivers[currentDriverIndex].first
                        supportFragmentManager.commit {
                            replace(
                                R.id.fragmentContainer,
                                FragmentCar.newInstance(::loadCarInfo, currentDriverIndex, driverName)
                            )
                        }
                    }
                    true
                }
                else -> false
            }
        }
    }



    // Загружаем всех гонщиков из API
    private suspend fun loadDriversList() {
        withContext(Dispatchers.IO) {
            try {
                val url = URL("https://api.openf1.org/v1/drivers?session_key=9158")
                val conn = url.openConnection() as HttpURLConnection
                conn.requestMethod = "GET"
                conn.connectTimeout = 5000
                conn.readTimeout = 5000

                val code = conn.responseCode
                if (code != 200) return@withContext

                val data = conn.inputStream.bufferedReader().readText()
                val json = JSONArray(data)

                drivers.clear()
                for (i in 0 until json.length()) {
                    val obj = json.getJSONObject(i)
                    val fullName = obj.optString("full_name")
                    val driverNum = obj.optInt("driver_number")
                    drivers.add(Pair(fullName, driverNum))
                }
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }
    }

    // Получаем информацию о выбранном гонщике
    private suspend fun loadDriverInfo(driverNumber: Int): String {
        return withContext(Dispatchers.IO) {
            try {
                val url =
                    URL("https://api.openf1.org/v1/drivers?driver_number=$driverNumber&session_key=9158")
                val conn = url.openConnection() as HttpURLConnection
                conn.requestMethod = "GET"
                conn.connectTimeout = 5000
                conn.readTimeout = 5000

                val code = conn.responseCode
                if (code != 200) return@withContext "Ошибка запроса: $code"

                val data = conn.inputStream.bufferedReader().readText()
                val json = JSONArray(data)
                if (json.length() == 0) return@withContext "Гонщик не найден"

                val obj = json.getJSONObject(0)
                val broadcastName = obj.optString("broadcast_name")
                val countryCode = obj.optString("country_code")
                val driverNum = obj.optInt("driver_number")
                val firstName = obj.optString("first_name")
                val fullName = obj.optString("full_name")
                val lastName = obj.optString("last_name")
                val teamName = obj.optString("team_name")
                val teamColour = obj.optString("team_colour")
                val nameAcronym = obj.optString("name_acronym")

                return@withContext """
                    Broadcast Name: $broadcastName
                    Полное имя: $fullName
                    Имя: $firstName
                    Фамилия: $lastName
                    Номер гонщика: $driverNum
                    Команда: $teamName
                    Цвет команды: #$teamColour
                    Акроним: $nameAcronym
                    Страна: $countryCode
                """.trimIndent()

            } catch (e: Exception) {
                return@withContext "Ошибка: ${e.message}"
            }
        }
    }

    // Получаем данные автомобиля по номеру гонщика
    private suspend fun loadCarInfo(driverNumber: Int): String {
        return withContext(Dispatchers.IO) {
            try {
                val url =
                    URL("https://api.openf1.org/v1/car_data?driver_number=$driverNumber&session_key=9159")
                val conn = url.openConnection() as HttpURLConnection
                conn.requestMethod = "GET"
                conn.connectTimeout = 5000
                conn.readTimeout = 5000

                val code = conn.responseCode
                if (code != 200) return@withContext "Ошибка запроса: $code"

                val data = conn.inputStream.bufferedReader().readText()
                val json = JSONArray(data)
                if (json.length() == 0) return@withContext "Данные автомобиля не найдены"

                val obj = json.getJSONObject(0)
                val speed = obj.optInt("speed")
                val nGear = obj.optInt("n_gear")
                val drs = obj.optInt("drs")
                val throttle = obj.optInt("throttle")
                val brake = obj.optInt("brake")
                val rpm = obj.optInt("rpm")
                val date = obj.optString("date")

                return@withContext """
                    Дата: $date
                    Скорость: $speed км/ч
                    Передача: $nGear
                    DRS: $drs
                    Газ: $throttle%
                    Тормоз: $brake
                    RPM: $rpm
                """.trimIndent()

            } catch (e: Exception) {
                return@withContext "Ошибка: ${e.message}"
            }
        }
    }
}
