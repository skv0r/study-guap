package com.example.autotasks.network

import com.example.autotasks.database.Driver
import kotlinx.coroutines.runBlocking

object SupabaseSync {
    
    // Синхронизация добавления гонщика в Supabase
    fun syncAddDriver(driver: Driver) {
        Thread {
            try {
                val supabaseDriver = SupabaseDriver(
                    id = null,
                    fullName = driver.fullName,
                    driverNumber = driver.driverNumber,
                    firstName = driver.firstName,
                    lastName = driver.lastName,
                    teamName = driver.teamName,
                    teamColour = driver.teamColour,
                    nameAcronym = driver.nameAcronym,
                    countryCode = driver.countryCode,
                    broadcastName = driver.broadcastName
                )
                
                runBlocking {
                    val response = SupabaseService.api.insertDriver(
                        SupabaseService.API_KEY,
                        SupabaseService.getAuthorizationHeader(),
                        "application/json",
                        "return=representation",
                        supabaseDriver
                    )
                    
                    if (!response.isSuccessful) {
                        android.util.Log.e("SupabaseSync", "Ошибка добавления в Supabase: ${response.code()}")
                    }
                }
            } catch (e: Exception) {
                android.util.Log.e("SupabaseSync", "Ошибка синхронизации добавления: ${e.message}")
            }
        }.start()
    }
    
    // Синхронизация обновления гонщика в Supabase
    fun syncUpdateDriver(driver: Driver) {
        Thread {
            try {
                android.util.Log.d("SupabaseSync", "Начало синхронизации обновления гонщика: ${driver.fullName}, номер: ${driver.driverNumber}")
                
                // Сначала нужно получить ID из Supabase по driver_number
                runBlocking {
                    // Используем формат фильтра Supabase: driver_number=eq.1
                    val driverNumberFilter = "eq.${driver.driverNumber}"
                    android.util.Log.d("SupabaseSync", "Фильтр поиска: driver_number=$driverNumberFilter")
                    
                    val getResponse = SupabaseService.api.getDriverByFilter(
                        driverNumberFilter = driverNumberFilter,
                        select = "*",
                        apiKey = SupabaseService.API_KEY,
                        authorization = SupabaseService.getAuthorizationHeader()
                    )
                    
                    android.util.Log.d("SupabaseSync", "Поиск гонщика: код ответа = ${getResponse.code()}, успешно = ${getResponse.isSuccessful}")
                    
                    if (getResponse.isSuccessful && getResponse.body() != null) {
                        val drivers = getResponse.body()!!
                        android.util.Log.d("SupabaseSync", "Найдено гонщиков: ${drivers.size}")
                        
                        if (drivers.isNotEmpty()) {
                            val supabaseId = drivers[0].id
                            android.util.Log.d("SupabaseSync", "ID гонщика в Supabase: $supabaseId")
                            
                            if (supabaseId != null) {
                                // При обновлении не отправляем id в теле, только в фильтре
                                val supabaseDriver = SupabaseDriver(
                                    id = null, // Не отправляем id в теле запроса
                                    fullName = driver.fullName,
                                    driverNumber = driver.driverNumber,
                                    firstName = driver.firstName,
                                    lastName = driver.lastName,
                                    teamName = driver.teamName,
                                    teamColour = driver.teamColour,
                                    nameAcronym = driver.nameAcronym,
                                    countryCode = driver.countryCode,
                                    broadcastName = driver.broadcastName
                                )
                                
                                android.util.Log.d("SupabaseSync", "Отправка обновления для ID: $supabaseId")
                                
                                // Формат фильтра для Supabase: id=eq.123 (полный формат)
                                val filter = "eq.$supabaseId"
                                android.util.Log.d("SupabaseSync", "Используемый фильтр: id=$filter")
                                
                                val updateResponse = SupabaseService.api.updateDriver(
                                    filter = filter,
                                    apiKey = SupabaseService.API_KEY,
                                    authorization = SupabaseService.getAuthorizationHeader(),
                                    contentType = "application/json",
                                    prefer = "return=representation",
                                    driver = supabaseDriver
                                )
                                
                                android.util.Log.d("SupabaseSync", "Ответ обновления: код = ${updateResponse.code()}, успешно = ${updateResponse.isSuccessful}")
                                
                                if (updateResponse.isSuccessful) {
                                    android.util.Log.d("SupabaseSync", "Гонщик успешно обновлен в Supabase")
                                } else {
                                    android.util.Log.e("SupabaseSync", "Ошибка обновления в Supabase: код ${updateResponse.code()}, сообщение: ${updateResponse.message()}")
                                    if (updateResponse.errorBody() != null) {
                                        val errorBody = updateResponse.errorBody()?.string()
                                        android.util.Log.e("SupabaseSync", "Тело ошибки: $errorBody")
                                    }
                                }
                            } else {
                                android.util.Log.e("SupabaseSync", "ID гонщика в Supabase равен null")
                            }
                        } else {
                            android.util.Log.w("SupabaseSync", "Гонщик с номером ${driver.driverNumber} не найден в Supabase для обновления")
                        }
                    } else {
                        android.util.Log.e("SupabaseSync", "Ошибка поиска гонщика: код ${getResponse.code()}, сообщение: ${getResponse.message()}")
                        if (getResponse.errorBody() != null) {
                            val errorBody = getResponse.errorBody()?.string()
                            android.util.Log.e("SupabaseSync", "Тело ошибки поиска: $errorBody")
                        }
                    }
                }
            } catch (e: Exception) {
                android.util.Log.e("SupabaseSync", "Ошибка синхронизации обновления: ${e.message}", e)
                e.printStackTrace()
            }
        }.start()
    }
    
    // Синхронизация удаления гонщика из Supabase
    fun syncDeleteDriver(driverNumber: Int) {
        Thread {
            try {
                android.util.Log.d("SupabaseSync", "Начало синхронизации удаления гонщика с номером: $driverNumber")
                
                runBlocking {
                    // Находим ID в Supabase по driver_number
                    val driverNumberFilter = "eq.$driverNumber"
                    android.util.Log.d("SupabaseSync", "Поиск гонщика для удаления: driver_number=$driverNumberFilter")
                    
                    val getResponse = SupabaseService.api.getDriverByFilter(
                        driverNumberFilter = driverNumberFilter,
                        select = "*",
                        apiKey = SupabaseService.API_KEY,
                        authorization = SupabaseService.getAuthorizationHeader()
                    )
                    
                    android.util.Log.d("SupabaseSync", "Поиск для удаления: код = ${getResponse.code()}, успешно = ${getResponse.isSuccessful}")
                    
                    if (getResponse.isSuccessful && getResponse.body() != null) {
                        val drivers = getResponse.body()!!
                        android.util.Log.d("SupabaseSync", "Найдено гонщиков для удаления: ${drivers.size}")
                        
                        if (drivers.isNotEmpty()) {
                            val supabaseId = drivers[0].id
                            android.util.Log.d("SupabaseSync", "ID гонщика для удаления: $supabaseId")
                            
                            if (supabaseId != null) {
                                // Формат фильтра для Supabase: id=eq.123 (полный формат)
                                val filter = "eq.$supabaseId"
                                android.util.Log.d("SupabaseSync", "Используемый фильтр для удаления: id=$filter")
                                
                                val deleteResponse = SupabaseService.api.deleteDriver(
                                    filter = filter,
                                    apiKey = SupabaseService.API_KEY,
                                    authorization = SupabaseService.getAuthorizationHeader(),
                                    prefer = "return=representation"
                                )
                                
                                android.util.Log.d("SupabaseSync", "Ответ удаления: код = ${deleteResponse.code()}, успешно = ${deleteResponse.isSuccessful}")
                                
                                if (deleteResponse.isSuccessful) {
                                    android.util.Log.d("SupabaseSync", "Гонщик успешно удален из Supabase")
                                } else {
                                    android.util.Log.e("SupabaseSync", "Ошибка удаления из Supabase: код ${deleteResponse.code()}, сообщение: ${deleteResponse.message()}")
                                    if (deleteResponse.errorBody() != null) {
                                        val errorBody = deleteResponse.errorBody()?.string()
                                        android.util.Log.e("SupabaseSync", "Тело ошибки удаления: $errorBody")
                                    }
                                }
                            } else {
                                android.util.Log.e("SupabaseSync", "ID гонщика для удаления равен null")
                            }
                        } else {
                            android.util.Log.w("SupabaseSync", "Гонщик с номером $driverNumber не найден в Supabase для удаления")
                        }
                    } else {
                        android.util.Log.e("SupabaseSync", "Ошибка поиска гонщика для удаления: код ${getResponse.code()}, сообщение: ${getResponse.message()}")
                        if (getResponse.errorBody() != null) {
                            val errorBody = getResponse.errorBody()?.string()
                            android.util.Log.e("SupabaseSync", "Тело ошибки поиска для удаления: $errorBody")
                        }
                    }
                }
            } catch (e: Exception) {
                android.util.Log.e("SupabaseSync", "Ошибка синхронизации удаления: ${e.message}", e)
                e.printStackTrace()
            }
        }.start()
    }
}

