package com.example.autotasks.database

import android.content.Context
import androidx.room.Database
import androidx.room.Room
import androidx.room.RoomDatabase

@Database(entities = [Driver::class, MapMarker::class], version = 3, exportSchema = false)
abstract class AppDatabase : RoomDatabase() {
    abstract fun driverDao(): DriverDao
    abstract fun mapMarkerDao(): MapMarkerDao

    companion object {
        @Volatile
        private var INSTANCE: AppDatabase? = null

        fun getDatabase(context: Context): AppDatabase {
            return INSTANCE ?: synchronized(this) {
                val instance = Room.databaseBuilder(
                    context.applicationContext,
                    AppDatabase::class.java,
                    "f1_drivers_database"
                )
                .fallbackToDestructiveMigration()
                .setQueryCallback({ sqlQuery, bindArgs ->
                    android.util.Log.d("Room", "SQL: $sqlQuery, Args: ${bindArgs?.joinToString()}")
                }, java.util.concurrent.Executors.newSingleThreadExecutor())
                .build()
                INSTANCE = instance
                instance
            }
        }
    }
}

