package com.example.autotasks

import android.os.Bundle
import android.widget.Button
import android.widget.EditText
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import com.example.autotasks.database.AppDatabase
import com.example.autotasks.database.Driver
import com.example.autotasks.network.SupabaseSync
import kotlinx.coroutines.launch

class EditDriverActivity : AppCompatActivity() {
    
    private lateinit var editFullName: EditText
    private lateinit var editDriverNumber: EditText
    private lateinit var editFirstName: EditText
    private lateinit var editLastName: EditText
    private lateinit var editTeamName: EditText
    private lateinit var editTeamColour: EditText
    private lateinit var editNameAcronym: EditText
    private lateinit var editCountryCode: EditText
    private lateinit var editBroadcastName: EditText
    private lateinit var btnUpdate: Button
    private lateinit var btnCancel: Button
    
    private lateinit var database: AppDatabase
    private var driverId: Int = -1
    
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_edit_driver)
        
        database = AppDatabase.getDatabase(this)
        driverId = intent.getIntExtra("DRIVER_ID", -1)
        
        if (driverId == -1) {
            Toast.makeText(this, "Ошибка: ID гонщика не найден", Toast.LENGTH_SHORT).show()
            finish()
            return
        }
        
        // Инициализация полей
        editFullName = findViewById(R.id.editFullName)
        editDriverNumber = findViewById(R.id.editDriverNumber)
        editFirstName = findViewById(R.id.editFirstName)
        editLastName = findViewById(R.id.editLastName)
        editTeamName = findViewById(R.id.editTeamName)
        editTeamColour = findViewById(R.id.editTeamColour)
        editNameAcronym = findViewById(R.id.editNameAcronym)
        editCountryCode = findViewById(R.id.editCountryCode)
        editBroadcastName = findViewById(R.id.editBroadcastName)
        btnUpdate = findViewById(R.id.btnUpdate)
        btnCancel = findViewById(R.id.btnCancel)
        
        // Загружаем данные гонщика
        loadDriverData()
        
        btnUpdate.setOnClickListener {
            updateDriver()
        }
        
        btnCancel.setOnClickListener {
            finish()
        }
    }
    
    private fun loadDriverData() {
        lifecycleScope.launch {
            try {
                val driver = database.driverDao().getDriverById(driverId)
                if (driver != null) {
                    editFullName.setText(driver.fullName)
                    editDriverNumber.setText(driver.driverNumber.toString())
                    editFirstName.setText(driver.firstName)
                    editLastName.setText(driver.lastName)
                    editTeamName.setText(driver.teamName)
                    editTeamColour.setText(driver.teamColour)
                    editNameAcronym.setText(driver.nameAcronym)
                    editCountryCode.setText(driver.countryCode)
                    editBroadcastName.setText(driver.broadcastName)
                } else {
                    Toast.makeText(this@EditDriverActivity, "Гонщик не найден", Toast.LENGTH_SHORT).show()
                    finish()
                }
            } catch (e: Exception) {
                Toast.makeText(this@EditDriverActivity, "Ошибка загрузки: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }
    
    private fun updateDriver() {
        val fullName = editFullName.text.toString().trim()
        val driverNumberStr = editDriverNumber.text.toString().trim()
        val firstName = editFirstName.text.toString().trim()
        val lastName = editLastName.text.toString().trim()
        val teamName = editTeamName.text.toString().trim()
        val teamColour = editTeamColour.text.toString().trim()
        val nameAcronym = editNameAcronym.text.toString().trim()
        val countryCode = editCountryCode.text.toString().trim()
        val broadcastName = editBroadcastName.text.toString().trim()
        
        if (fullName.isEmpty() || driverNumberStr.isEmpty()) {
            Toast.makeText(this, "Заполните обязательные поля (Имя и Номер)", Toast.LENGTH_SHORT).show()
            return
        }
        
        val driverNumber = driverNumberStr.toIntOrNull()
        if (driverNumber == null) {
            Toast.makeText(this, "Номер должен быть числом", Toast.LENGTH_SHORT).show()
            return
        }
        
        val driver = Driver(
            id = driverId,
            fullName = fullName,
            driverNumber = driverNumber,
            firstName = firstName,
            lastName = lastName,
            teamName = teamName,
            teamColour = teamColour,
            nameAcronym = nameAcronym,
            countryCode = countryCode,
            broadcastName = broadcastName
        )
        
        lifecycleScope.launch {
            try {
                // Обновляем в локальной БД
                database.driverDao().updateDriver(driver)
                
                // Проверяем, что данные действительно сохранились
                val updated = database.driverDao().getDriverById(driver.id)
                if (updated == null) {
                    Toast.makeText(this@EditDriverActivity, "Ошибка: данные не сохранились", Toast.LENGTH_SHORT).show()
                    return@launch
                }
                
                // Синхронизируем с Supabase в фоне
                SupabaseSync.syncUpdateDriver(driver)
                
                Toast.makeText(this@EditDriverActivity, "Гонщик обновлен", Toast.LENGTH_SHORT).show()
                setResult(RESULT_OK)
                finish()
            } catch (e: Exception) {
                e.printStackTrace()
                Toast.makeText(this@EditDriverActivity, "Ошибка: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }
}

