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

class AddDriverActivity : AppCompatActivity() {
    
    private lateinit var editFullName: EditText
    private lateinit var editDriverNumber: EditText
    private lateinit var editFirstName: EditText
    private lateinit var editLastName: EditText
    private lateinit var editTeamName: EditText
    private lateinit var editTeamColour: EditText
    private lateinit var editNameAcronym: EditText
    private lateinit var editCountryCode: EditText
    private lateinit var editBroadcastName: EditText
    private lateinit var btnSave: Button
    private lateinit var btnCancel: Button
    
    private lateinit var database: AppDatabase
    
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_add_driver)
        
        database = AppDatabase.getDatabase(this)
        
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
        btnSave = findViewById(R.id.btnSave)
        btnCancel = findViewById(R.id.btnCancel)
        
        btnSave.setOnClickListener {
            saveDriver()
        }
        
        btnCancel.setOnClickListener {
            finish()
        }
    }
    
    private fun saveDriver() {
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
                // Сохраняем в локальную БД
                database.driverDao().insertDriver(driver)
                
                // Синхронизируем с Supabase в фоне
                SupabaseSync.syncAddDriver(driver)
                
                Toast.makeText(this@AddDriverActivity, "Гонщик добавлен", Toast.LENGTH_SHORT).show()
                setResult(RESULT_OK)
                finish()
            } catch (e: Exception) {
                Toast.makeText(this@AddDriverActivity, "Ошибка: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }
}

