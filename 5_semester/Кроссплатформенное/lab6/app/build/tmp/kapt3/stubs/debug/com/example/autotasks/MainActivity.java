package com.example.autotasks;

@kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000Z\n\u0002\u0018\u0002\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0010\b\n\u0002\b\u0005\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010!\n\u0002\u0018\u0002\n\u0002\b\u0003\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010\t\n\u0002\b\u0005\n\u0002\u0010\u0002\n\u0002\b\u0004\n\u0002\u0010\u000b\n\u0000\n\u0002\u0010\u000e\n\u0002\b\f\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0002\b\u0006\u0018\u0000 52\u00020\u0001:\u00015B\u0005\u00a2\u0006\u0002\u0010\u0002J\u0010\u0010\u0018\u001a\u00020\u00192\u0006\u0010\u001a\u001a\u00020\rH\u0002J\u0010\u0010\u001b\u001a\u00020\u00192\u0006\u0010\u001a\u001a\u00020\rH\u0002J\u0010\u0010\u001c\u001a\u00020\u00192\u0006\u0010\u001a\u001a\u00020\rH\u0002J\b\u0010\u001d\u001a\u00020\u001eH\u0002J\u0016\u0010\u001f\u001a\u00020 2\u0006\u0010!\u001a\u00020\u0004H\u0082@\u00a2\u0006\u0002\u0010\"J\u0016\u0010#\u001a\u00020 2\u0006\u0010\u001a\u001a\u00020\rH\u0082@\u00a2\u0006\u0002\u0010$J\u000e\u0010%\u001a\u00020\u0019H\u0082@\u00a2\u0006\u0002\u0010&J\b\u0010\'\u001a\u00020\u0019H\u0002J\u000e\u0010(\u001a\u00020\u0019H\u0082@\u00a2\u0006\u0002\u0010&J\"\u0010)\u001a\u00020\u00192\u0006\u0010*\u001a\u00020\u00042\u0006\u0010+\u001a\u00020\u00042\b\u0010,\u001a\u0004\u0018\u00010-H\u0014J\u0012\u0010.\u001a\u00020\u00192\b\u0010/\u001a\u0004\u0018\u000100H\u0014J\b\u00101\u001a\u00020\u0019H\u0002J\b\u00102\u001a\u00020\u0019H\u0002J\u0010\u00103\u001a\u00020\u00192\u0006\u0010\u001a\u001a\u00020\rH\u0002J\u0010\u00104\u001a\u00020\u00192\u0006\u0010\u001a\u001a\u00020\rH\u0002R\u001a\u0010\u0003\u001a\u00020\u0004X\u0086\u000e\u00a2\u0006\u000e\n\u0000\u001a\u0004\b\u0005\u0010\u0006\"\u0004\b\u0007\u0010\bR\u000e\u0010\t\u001a\u00020\nX\u0082.\u00a2\u0006\u0002\n\u0000R\u0017\u0010\u000b\u001a\b\u0012\u0004\u0012\u00020\r0\f\u00a2\u0006\b\n\u0000\u001a\u0004\b\u000e\u0010\u000fR\u000e\u0010\u0010\u001a\u00020\u0011X\u0082.\u00a2\u0006\u0002\n\u0000R\u001a\u0010\u0012\u001a\u00020\u0013X\u0086\u000e\u00a2\u0006\u000e\n\u0000\u001a\u0004\b\u0014\u0010\u0015\"\u0004\b\u0016\u0010\u0017\u00a8\u00066"}, d2 = {"Lcom/example/autotasks/MainActivity;", "Landroidx/appcompat/app/AppCompatActivity;", "()V", "currentDriverIndex", "", "getCurrentDriverIndex", "()I", "setCurrentDriverIndex", "(I)V", "database", "Lcom/example/autotasks/database/AppDatabase;", "drivers", "", "Lcom/example/autotasks/database/Driver;", "getDrivers", "()Ljava/util/List;", "fragmentDriver", "Lcom/example/autotasks/FragmentDriver;", "lastLoadTime", "", "getLastLoadTime", "()J", "setLastLoadTime", "(J)V", "confirmDeleteDriver", "", "driver", "deleteDriver", "editDriver", "isNetworkAvailable", "", "loadCarInfo", "", "driverNumber", "(ILkotlin/coroutines/Continuation;)Ljava/lang/Object;", "loadDriverInfo", "(Lcom/example/autotasks/database/Driver;Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "loadDriversFromDatabase", "(Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "loadDriversFromSupabase", "loadDriversFromSupabaseSync", "onActivityResult", "requestCode", "resultCode", "data", "Landroid/content/Intent;", "onCreate", "savedInstanceState", "Landroid/os/Bundle;", "refreshDrivers", "refreshFromSupabase", "showDriverActionsDialog", "showDriverDetails", "Companion", "app_debug"})
@kotlin.Suppress(names = {"DEPRECATION"})
public final class MainActivity extends androidx.appcompat.app.AppCompatActivity {
    @org.jetbrains.annotations.NotNull()
    private final java.util.List<com.example.autotasks.database.Driver> drivers = null;
    private int currentDriverIndex = 0;
    private long lastLoadTime = 0L;
    private com.example.autotasks.FragmentDriver fragmentDriver;
    private com.example.autotasks.database.AppDatabase database;
    public static final int REQUEST_ADD_DRIVER = 1;
    public static final int REQUEST_EDIT_DRIVER = 2;
    @org.jetbrains.annotations.NotNull()
    public static final com.example.autotasks.MainActivity.Companion Companion = null;
    
    public MainActivity() {
        super();
    }
    
    @org.jetbrains.annotations.NotNull()
    public final java.util.List<com.example.autotasks.database.Driver> getDrivers() {
        return null;
    }
    
    public final int getCurrentDriverIndex() {
        return 0;
    }
    
    public final void setCurrentDriverIndex(int p0) {
    }
    
    public final long getLastLoadTime() {
        return 0L;
    }
    
    public final void setLastLoadTime(long p0) {
    }
    
    private final boolean isNetworkAvailable() {
        return false;
    }
    
    @java.lang.Override()
    protected void onCreate(@org.jetbrains.annotations.Nullable()
    android.os.Bundle savedInstanceState) {
    }
    
    private final java.lang.Object loadDriversFromDatabase(kotlin.coroutines.Continuation<? super kotlin.Unit> $completion) {
        return null;
    }
    
    private final void refreshFromSupabase() {
    }
    
    private final java.lang.Object loadDriversFromSupabaseSync(kotlin.coroutines.Continuation<? super kotlin.Unit> $completion) {
        return null;
    }
    
    private final void loadDriversFromSupabase() {
    }
    
    private final void refreshDrivers() {
    }
    
    private final void showDriverActionsDialog(com.example.autotasks.database.Driver driver) {
    }
    
    private final void showDriverDetails(com.example.autotasks.database.Driver driver) {
    }
    
    private final void editDriver(com.example.autotasks.database.Driver driver) {
    }
    
    private final void confirmDeleteDriver(com.example.autotasks.database.Driver driver) {
    }
    
    private final void deleteDriver(com.example.autotasks.database.Driver driver) {
    }
    
    @java.lang.Override()
    protected void onActivityResult(int requestCode, int resultCode, @org.jetbrains.annotations.Nullable()
    android.content.Intent data) {
    }
    
    private final java.lang.Object loadDriverInfo(com.example.autotasks.database.Driver driver, kotlin.coroutines.Continuation<? super java.lang.String> $completion) {
        return null;
    }
    
    private final java.lang.Object loadCarInfo(int driverNumber, kotlin.coroutines.Continuation<? super java.lang.String> $completion) {
        return null;
    }
    
    @kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000\u0014\n\u0002\u0018\u0002\n\u0002\u0010\u0000\n\u0002\b\u0002\n\u0002\u0010\b\n\u0002\b\u0002\b\u0086\u0003\u0018\u00002\u00020\u0001B\u0007\b\u0002\u00a2\u0006\u0002\u0010\u0002R\u000e\u0010\u0003\u001a\u00020\u0004X\u0086T\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0005\u001a\u00020\u0004X\u0086T\u00a2\u0006\u0002\n\u0000\u00a8\u0006\u0006"}, d2 = {"Lcom/example/autotasks/MainActivity$Companion;", "", "()V", "REQUEST_ADD_DRIVER", "", "REQUEST_EDIT_DRIVER", "app_debug"})
    public static final class Companion {
        
        private Companion() {
            super();
        }
    }
}