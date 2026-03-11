package com.example.autotasks.network

import okhttp3.OkHttpClient
import okhttp3.logging.HttpLoggingInterceptor
import retrofit2.Retrofit
import retrofit2.converter.gson.GsonConverterFactory
import java.util.concurrent.TimeUnit

object SupabaseService {
    private const val BASE_URL = "https://qbxlkrikikrjvlywvbmd.supabase.co/rest/v1/"
    const val API_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFieGxrcmlraWtyanZseXd2Ym1kIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjQ0MzAwNzAsImV4cCI6MjA4MDAwNjA3MH0.QMk6fj6-zuzbkCwEvwlom8mMGSR8cQRmD6bnKlhTf5I"
    
    private val loggingInterceptor = HttpLoggingInterceptor().apply {
        level = HttpLoggingInterceptor.Level.BODY
    }
    
    private val client = OkHttpClient.Builder()
        .addInterceptor(loggingInterceptor)
        .connectTimeout(30, TimeUnit.SECONDS)
        .readTimeout(30, TimeUnit.SECONDS)
        .writeTimeout(30, TimeUnit.SECONDS)
        .build()
    
    private val retrofit = Retrofit.Builder()
        .baseUrl(BASE_URL)
        .client(client)
        .addConverterFactory(GsonConverterFactory.create())
        .build()
    
    val api: SupabaseApi = retrofit.create(SupabaseApi::class.java)
    
    fun getAuthorizationHeader(): String {
        return "Bearer $API_KEY"
    }
}

