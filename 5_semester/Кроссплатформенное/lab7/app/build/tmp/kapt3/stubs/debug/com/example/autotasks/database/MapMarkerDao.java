package com.example.autotasks.database;

@kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000*\n\u0002\u0018\u0002\n\u0002\u0010\u0000\n\u0000\n\u0002\u0010\u0002\n\u0002\b\u0003\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0010 \n\u0002\b\u0002\n\u0002\u0010\t\n\u0002\b\u0004\bg\u0018\u00002\u00020\u0001J\u000e\u0010\u0002\u001a\u00020\u0003H\u00a7@\u00a2\u0006\u0002\u0010\u0004J\u0016\u0010\u0005\u001a\u00020\u00032\u0006\u0010\u0006\u001a\u00020\u0007H\u00a7@\u00a2\u0006\u0002\u0010\bJ\u0014\u0010\t\u001a\b\u0012\u0004\u0012\u00020\u00070\nH\u00a7@\u00a2\u0006\u0002\u0010\u0004J\u0018\u0010\u000b\u001a\u0004\u0018\u00010\u00072\u0006\u0010\f\u001a\u00020\rH\u00a7@\u00a2\u0006\u0002\u0010\u000eJ\u0016\u0010\u000f\u001a\u00020\r2\u0006\u0010\u0006\u001a\u00020\u0007H\u00a7@\u00a2\u0006\u0002\u0010\bJ\u0016\u0010\u0010\u001a\u00020\u00032\u0006\u0010\u0006\u001a\u00020\u0007H\u00a7@\u00a2\u0006\u0002\u0010\b\u00a8\u0006\u0011"}, d2 = {"Lcom/example/autotasks/database/MapMarkerDao;", "", "deleteAllMarkers", "", "(Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "deleteMarker", "marker", "Lcom/example/autotasks/database/MapMarker;", "(Lcom/example/autotasks/database/MapMarker;Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "getAllMarkers", "", "getMarkerById", "id", "", "(JLkotlin/coroutines/Continuation;)Ljava/lang/Object;", "insertMarker", "updateMarker", "app_debug"})
@androidx.room.Dao()
public abstract interface MapMarkerDao {
    
    @androidx.room.Query(value = "SELECT * FROM map_markers ORDER BY timestamp DESC")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object getAllMarkers(@org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super java.util.List<com.example.autotasks.database.MapMarker>> $completion);
    
    @androidx.room.Insert()
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object insertMarker(@org.jetbrains.annotations.NotNull()
    com.example.autotasks.database.MapMarker marker, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super java.lang.Long> $completion);
    
    @androidx.room.Update()
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object updateMarker(@org.jetbrains.annotations.NotNull()
    com.example.autotasks.database.MapMarker marker, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super kotlin.Unit> $completion);
    
    @androidx.room.Delete()
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object deleteMarker(@org.jetbrains.annotations.NotNull()
    com.example.autotasks.database.MapMarker marker, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super kotlin.Unit> $completion);
    
    @androidx.room.Query(value = "DELETE FROM map_markers")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object deleteAllMarkers(@org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super kotlin.Unit> $completion);
    
    @androidx.room.Query(value = "SELECT * FROM map_markers WHERE id = :id")
    @org.jetbrains.annotations.Nullable()
    public abstract java.lang.Object getMarkerById(long id, @org.jetbrains.annotations.NotNull()
    kotlin.coroutines.Continuation<? super com.example.autotasks.database.MapMarker> $completion);
}