package com.example.autotasks.database

import androidx.room.ColumnInfo
import androidx.room.Entity
import androidx.room.Index
import androidx.room.PrimaryKey

@Entity(
    tableName = "drivers",
    indices = [Index(value = ["driver_number"], unique = true)]
)
data class Driver(
    @PrimaryKey(autoGenerate = true)
    val id: Int = 0,
    @ColumnInfo(name = "full_name")
    val fullName: String,
    @ColumnInfo(name = "driver_number")
    val driverNumber: Int,
    @ColumnInfo(name = "first_name")
    val firstName: String,
    @ColumnInfo(name = "last_name")
    val lastName: String,
    @ColumnInfo(name = "team_name")
    val teamName: String,
    @ColumnInfo(name = "team_colour")
    val teamColour: String,
    @ColumnInfo(name = "name_acronym")
    val nameAcronym: String,
    @ColumnInfo(name = "country_code")
    val countryCode: String,
    @ColumnInfo(name = "broadcast_name")
    val broadcastName: String
)

