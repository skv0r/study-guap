package com.example.autotasks.network;

@kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000(\n\u0002\u0018\u0002\n\u0002\u0010\u0000\n\u0000\n\u0002\u0018\u0002\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010\u000e\n\u0002\b\u0005\n\u0002\u0010 \n\u0002\u0018\u0002\n\u0002\b\u000b\bf\u0018\u00002\u00020\u0001J<\u0010\u0002\u001a\b\u0012\u0004\u0012\u00020\u00040\u00032\b\b\u0001\u0010\u0005\u001a\u00020\u00062\b\b\u0001\u0010\u0007\u001a\u00020\u00062\b\b\u0001\u0010\b\u001a\u00020\u00062\b\b\u0003\u0010\t\u001a\u00020\u0006H\u00a7@\u00a2\u0006\u0002\u0010\nJ.\u0010\u000b\u001a\u000e\u0012\n\u0012\b\u0012\u0004\u0012\u00020\r0\f0\u00032\b\b\u0001\u0010\u0007\u001a\u00020\u00062\b\b\u0001\u0010\b\u001a\u00020\u0006H\u00a7@\u00a2\u0006\u0002\u0010\u000eJD\u0010\u000f\u001a\u000e\u0012\n\u0012\b\u0012\u0004\u0012\u00020\r0\f0\u00032\n\b\u0003\u0010\u0010\u001a\u0004\u0018\u00010\u00062\b\b\u0003\u0010\u0011\u001a\u00020\u00062\b\b\u0001\u0010\u0007\u001a\u00020\u00062\b\b\u0001\u0010\b\u001a\u00020\u0006H\u00a7@\u00a2\u0006\u0002\u0010\nJL\u0010\u0012\u001a\u000e\u0012\n\u0012\b\u0012\u0004\u0012\u00020\r0\f0\u00032\b\b\u0001\u0010\u0007\u001a\u00020\u00062\b\b\u0001\u0010\b\u001a\u00020\u00062\b\b\u0003\u0010\u0013\u001a\u00020\u00062\b\b\u0003\u0010\t\u001a\u00020\u00062\b\b\u0001\u0010\u0014\u001a\u00020\rH\u00a7@\u00a2\u0006\u0002\u0010\u0015JV\u0010\u0016\u001a\u000e\u0012\n\u0012\b\u0012\u0004\u0012\u00020\r0\f0\u00032\b\b\u0001\u0010\u0005\u001a\u00020\u00062\b\b\u0001\u0010\u0007\u001a\u00020\u00062\b\b\u0001\u0010\b\u001a\u00020\u00062\b\b\u0003\u0010\u0013\u001a\u00020\u00062\b\b\u0003\u0010\t\u001a\u00020\u00062\b\b\u0001\u0010\u0014\u001a\u00020\rH\u00a7@\u00a2\u0006\u0002\u0010\u0017\u00a8\u0006\u0018"}, d2 = {"Lcom/example/autotasks/network/SupabaseApi;", "", "deleteDriver", "Lretrofit2/Response;", "Ljava/lang/Void;", "filter", "", "apiKey", "authorization", "prefer", "(Ljava/lang/String;Ljava/lang/String;Ljava/lang/String;Ljava/lang/String;Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "getAllDrivers", "", "Lcom/example/autotasks/network/SupabaseDriver;", "(Ljava/lang/String;Ljava/lang/String;Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "getDriverByFilter", "driverNumberFilter", "select", "insertDriver", "contentType", "driver", "(Ljava/lang/String;Ljava/lang/String;Ljava/lang/String;Ljava/lang/String;Lcom/example/autotasks/network/SupabaseDriver;Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "updateDriver", "(Ljava/lang/String;Ljava/lang/String;Ljava/lang/String;Ljava/lang/String;Ljava/lang/String;Lcom/example/autotasks/network/SupabaseDriver;Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "app_debug"})
public abstract interface SupabaseApi {
    
    @retrofit2.http.GET(value = "drivers")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object getAllDrivers(@retrofit2.http.Header(value = "apikey")
    @org.jetbrains.annotations.NotNull()
    java.lang.String apiKey, @retrofit2.http.Header(value = "Authorization")
    @org.jetbrains.annotations.NotNull()
    java.lang.String authorization, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super retrofit2.Response<java.util.List<com.example.autotasks.network.SupabaseDriver>>> $completion);
    
    @retrofit2.http.GET(value = "drivers")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object getDriverByFilter(@retrofit2.http.Query(value = "driver_number")
    @org.jetbrains.annotations.Nullable()
    java.lang.String driverNumberFilter, @retrofit2.http.Query(value = "select")
    @org.jetbrains.annotations.NotNull()
    java.lang.String select, @retrofit2.http.Header(value = "apikey")
    @org.jetbrains.annotations.NotNull()
    java.lang.String apiKey, @retrofit2.http.Header(value = "Authorization")
    @org.jetbrains.annotations.NotNull()
    java.lang.String authorization, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super retrofit2.Response<java.util.List<com.example.autotasks.network.SupabaseDriver>>> $completion);
    
    @retrofit2.http.POST(value = "drivers")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object insertDriver(@retrofit2.http.Header(value = "apikey")
    @org.jetbrains.annotations.NotNull()
    java.lang.String apiKey, @retrofit2.http.Header(value = "Authorization")
    @org.jetbrains.annotations.NotNull()
    java.lang.String authorization, @retrofit2.http.Header(value = "Content-Type")
    @org.jetbrains.annotations.NotNull()
    java.lang.String contentType, @retrofit2.http.Header(value = "Prefer")
    @org.jetbrains.annotations.NotNull()
    java.lang.String prefer, @retrofit2.http.Body()
    @org.jetbrains.annotations.NotNull()
    com.example.autotasks.network.SupabaseDriver driver, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super retrofit2.Response<java.util.List<com.example.autotasks.network.SupabaseDriver>>> $completion);
    
    @retrofit2.http.PATCH(value = "drivers")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object updateDriver(@retrofit2.http.Query(value = "id")
    @org.jetbrains.annotations.NotNull()
    java.lang.String filter, @retrofit2.http.Header(value = "apikey")
    @org.jetbrains.annotations.NotNull()
    java.lang.String apiKey, @retrofit2.http.Header(value = "Authorization")
    @org.jetbrains.annotations.NotNull()
    java.lang.String authorization, @retrofit2.http.Header(value = "Content-Type")
    @org.jetbrains.annotations.NotNull()
    java.lang.String contentType, @retrofit2.http.Header(value = "Prefer")
    @org.jetbrains.annotations.NotNull()
    java.lang.String prefer, @retrofit2.http.Body()
    @org.jetbrains.annotations.NotNull()
    com.example.autotasks.network.SupabaseDriver driver, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super retrofit2.Response<java.util.List<com.example.autotasks.network.SupabaseDriver>>> $completion);
    
    @retrofit2.http.DELETE(value = "drivers")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object deleteDriver(@retrofit2.http.Query(value = "id")
    @org.jetbrains.annotations.NotNull()
    java.lang.String filter, @retrofit2.http.Header(value = "apikey")
    @org.jetbrains.annotations.NotNull()
    java.lang.String apiKey, @retrofit2.http.Header(value = "Authorization")
    @org.jetbrains.annotations.NotNull()
    java.lang.String authorization, @retrofit2.http.Header(value = "Prefer")
    @org.jetbrains.annotations.NotNull()
    java.lang.String prefer, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super retrofit2.Response<java.lang.Void>> $completion);
    
    @kotlin.Metadata(mv = {1, 9, 0}, k = 3, xi = 48)
    public static final class DefaultImpls {
    }
}