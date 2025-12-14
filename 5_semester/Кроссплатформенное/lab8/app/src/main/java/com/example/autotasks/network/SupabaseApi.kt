package com.example.autotasks.network

import retrofit2.Response
import retrofit2.http.*

interface SupabaseApi {
    
    @GET("drivers")
    suspend fun getAllDrivers(
        @Header("apikey") apiKey: String,
        @Header("Authorization") authorization: String
    ): Response<List<SupabaseDriver>>
    
    @GET("drivers")
    suspend fun getDriverByFilter(
        @Query("driver_number") driverNumberFilter: String? = null,
        @Query("select") select: String = "*",
        @Header("apikey") apiKey: String,
        @Header("Authorization") authorization: String
    ): Response<List<SupabaseDriver>>
    
    @POST("drivers")
    suspend fun insertDriver(
        @Header("apikey") apiKey: String,
        @Header("Authorization") authorization: String,
        @Header("Content-Type") contentType: String = "application/json",
        @Header("Prefer") prefer: String = "return=representation",
        @Body driver: SupabaseDriver
    ): Response<List<SupabaseDriver>>
    
    @PATCH("drivers")
    suspend fun updateDriver(
        @Query("id") filter: String,
        @Header("apikey") apiKey: String,
        @Header("Authorization") authorization: String,
        @Header("Content-Type") contentType: String = "application/json",
        @Header("Prefer") prefer: String = "return=representation",
        @Body driver: SupabaseDriver
    ): Response<List<SupabaseDriver>>
    
    @DELETE("drivers")
    suspend fun deleteDriver(
        @Query("id") filter: String,
        @Header("apikey") apiKey: String,
        @Header("Authorization") authorization: String,
        @Header("Prefer") prefer: String = "return=representation"
    ): Response<Void>
}

