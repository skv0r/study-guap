package com.example.autotasks.database

import androidx.room.*

@Dao
interface MapMarkerDao {
    @Query("SELECT * FROM map_markers ORDER BY timestamp DESC")
    suspend fun getAllMarkers(): List<MapMarker>
    
    @Insert
    suspend fun insertMarker(marker: MapMarker): Long
    
    @Update
    suspend fun updateMarker(marker: MapMarker)
    
    @Delete
    suspend fun deleteMarker(marker: MapMarker)
    
    @Query("DELETE FROM map_markers")
    suspend fun deleteAllMarkers()
    
    @Query("SELECT * FROM map_markers WHERE id = :id")
    suspend fun getMarkerById(id: Long): MapMarker?
}

