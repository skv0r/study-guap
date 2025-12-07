package com.example.autotasks.network

import com.google.gson.annotations.SerializedName

data class SupabaseDriver(
    @SerializedName("id")
    val id: Int? = null,
    
    @SerializedName("full_name")
    val fullName: String,
    
    @SerializedName("driver_number")
    val driverNumber: Int,
    
    @SerializedName("first_name")
    val firstName: String,
    
    @SerializedName("last_name")
    val lastName: String,
    
    @SerializedName("team_name")
    val teamName: String,
    
    @SerializedName("team_colour")
    val teamColour: String,
    
    @SerializedName("name_acronym")
    val nameAcronym: String,
    
    @SerializedName("country_code")
    val countryCode: String,
    
    @SerializedName("broadcast_name")
    val broadcastName: String
)

