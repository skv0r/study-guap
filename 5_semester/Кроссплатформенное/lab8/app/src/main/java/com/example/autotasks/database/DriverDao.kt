package com.example.autotasks.database

import androidx.room.*

@Dao
interface DriverDao {
    @Query("SELECT * FROM drivers ORDER BY driver_number ASC")
    suspend fun getAllDrivers(): List<Driver>
    
    @Query("SELECT COUNT(*) FROM drivers")
    suspend fun getDriversCount(): Int

    @Query("SELECT * FROM drivers WHERE id = :id")
    suspend fun getDriverById(id: Int): Driver?

    @Query("SELECT * FROM drivers WHERE driver_number = :driverNumber")
    suspend fun getDriverByNumber(driverNumber: Int): Driver?

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun insertDriver(driver: Driver): Long

    @Update
    suspend fun updateDriver(driver: Driver)

    @Delete
    suspend fun deleteDriver(driver: Driver)

    @Query("DELETE FROM drivers")
    suspend fun deleteAllDrivers()
}

