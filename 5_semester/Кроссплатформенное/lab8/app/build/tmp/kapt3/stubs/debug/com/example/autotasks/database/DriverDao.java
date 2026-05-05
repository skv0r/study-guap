package com.example.autotasks.database;

@kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u00002\n\u0002\u0018\u0002\n\u0002\u0010\u0000\n\u0000\n\u0002\u0010\u0002\n\u0002\b\u0003\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0010 \n\u0002\b\u0002\n\u0002\u0010\b\n\u0002\b\u0005\n\u0002\u0010\t\n\u0002\b\u0002\bg\u0018\u00002\u00020\u0001J\u000e\u0010\u0002\u001a\u00020\u0003H\u00a7@\u00a2\u0006\u0002\u0010\u0004J\u0016\u0010\u0005\u001a\u00020\u00032\u0006\u0010\u0006\u001a\u00020\u0007H\u00a7@\u00a2\u0006\u0002\u0010\bJ\u0014\u0010\t\u001a\b\u0012\u0004\u0012\u00020\u00070\nH\u00a7@\u00a2\u0006\u0002\u0010\u0004J\u0018\u0010\u000b\u001a\u0004\u0018\u00010\u00072\u0006\u0010\f\u001a\u00020\rH\u00a7@\u00a2\u0006\u0002\u0010\u000eJ\u0018\u0010\u000f\u001a\u0004\u0018\u00010\u00072\u0006\u0010\u0010\u001a\u00020\rH\u00a7@\u00a2\u0006\u0002\u0010\u000eJ\u000e\u0010\u0011\u001a\u00020\rH\u00a7@\u00a2\u0006\u0002\u0010\u0004J\u0016\u0010\u0012\u001a\u00020\u00132\u0006\u0010\u0006\u001a\u00020\u0007H\u00a7@\u00a2\u0006\u0002\u0010\bJ\u0016\u0010\u0014\u001a\u00020\u00032\u0006\u0010\u0006\u001a\u00020\u0007H\u00a7@\u00a2\u0006\u0002\u0010\b\u00a8\u0006\u0015"}, d2 = {"Lcom/example/autotasks/database/DriverDao;", "", "deleteAllDrivers", "", "(Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "deleteDriver", "driver", "Lcom/example/autotasks/database/Driver;", "(Lcom/example/autotasks/database/Driver;Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "getAllDrivers", "", "getDriverById", "id", "", "(ILkotlin/coroutines/Continuation;)Ljava/lang/Object;", "getDriverByNumber", "driverNumber", "getDriversCount", "insertDriver", "", "updateDriver", "app_debug"})
@androidx.room.Dao()
public abstract interface DriverDao {
    
    @androidx.room.Query(value = "SELECT * FROM drivers ORDER BY driver_number ASC")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object getAllDrivers(@org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super java.util.List<com.example.autotasks.database.Driver>> $completion);
    
    @androidx.room.Query(value = "SELECT COUNT(*) FROM drivers")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object getDriversCount(@org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super java.lang.Integer> $completion);
    
    @androidx.room.Query(value = "SELECT * FROM drivers WHERE id = :id")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object getDriverById(int id, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super com.example.autotasks.database.Driver> $completion);
    
    @androidx.room.Query(value = "SELECT * FROM drivers WHERE driver_number = :driverNumber")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object getDriverByNumber(int driverNumber, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super com.example.autotasks.database.Driver> $completion);
    
    @androidx.room.Insert(onConflict = 1)
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object insertDriver(@org.jetbrains.annotations.NotNull()
    com.example.autotasks.database.Driver driver, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super java.lang.Long> $completion);
    
    @androidx.room.Update()
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object updateDriver(@org.jetbrains.annotations.NotNull()
    com.example.autotasks.database.Driver driver, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super kotlin.Unit> $completion);
    
    @androidx.room.Delete()
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object deleteDriver(@org.jetbrains.annotations.NotNull()
    com.example.autotasks.database.Driver driver, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super kotlin.Unit> $completion);
    
    @androidx.room.Query(value = "DELETE FROM drivers")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object deleteAllDrivers(@org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super kotlin.Unit> $completion);
}