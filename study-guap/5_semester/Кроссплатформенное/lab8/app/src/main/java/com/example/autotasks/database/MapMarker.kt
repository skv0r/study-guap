package com.example.autotasks.database

import androidx.room.Entity
import androidx.room.PrimaryKey

@Entity(tableName = "map_markers")
data class MapMarker(
    @PrimaryKey(autoGenerate = true)
    val id: Long = 0,
    val title: String,
    val description: String,
    val latitude: Double,
    val longitude: Double,
    val timestamp: Long = System.currentTimeMillis()
)

