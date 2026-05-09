package com.example.autotasks;

@kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000f\n\u0002\u0018\u0002\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0010 \n\u0002\u0018\u0002\n\u0002\u0010\u000e\n\u0002\u0010\u0006\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010!\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010\u0002\n\u0002\b\b\n\u0002\u0018\u0002\n\u0002\b\u0003\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0010\b\n\u0000\n\u0002\u0010\u0011\n\u0000\n\u0002\u0010\u0015\n\u0002\b\f\u0018\u0000 12\u00020\u0001:\u00011B\u0005\u00a2\u0006\u0002\u0010\u0002J \u0010\u0011\u001a\u00020\u00122\u0006\u0010\u0013\u001a\u00020\u00062\u0006\u0010\u0014\u001a\u00020\u00062\u0006\u0010\u0015\u001a\u00020\tH\u0002J\b\u0010\u0016\u001a\u00020\u0012H\u0002J\b\u0010\u0017\u001a\u00020\u0012H\u0002J\b\u0010\u0018\u001a\u00020\u0012H\u0002J\u0010\u0010\u0019\u001a\u00020\u00122\u0006\u0010\u001a\u001a\u00020\u001bH\u0002J\b\u0010\u001c\u001a\u00020\u0012H\u0002J\u0012\u0010\u001d\u001a\u00020\u00122\b\u0010\u001e\u001a\u0004\u0018\u00010\u001fH\u0014J-\u0010 \u001a\u00020\u00122\u0006\u0010!\u001a\u00020\"2\u000e\u0010#\u001a\n\u0012\u0006\b\u0001\u0012\u00020\u00060$2\u0006\u0010%\u001a\u00020&H\u0016\u00a2\u0006\u0002\u0010\'J\b\u0010(\u001a\u00020\u0012H\u0014J\b\u0010)\u001a\u00020\u0012H\u0014J\b\u0010*\u001a\u00020\u0012H\u0002J\b\u0010+\u001a\u00020\u0012H\u0002J\b\u0010,\u001a\u00020\u0012H\u0002J\u0010\u0010-\u001a\u00020\u00122\u0006\u0010\u0015\u001a\u00020\tH\u0002J\u0018\u0010.\u001a\u00020\u00122\u0006\u0010/\u001a\u00020\u00062\u0006\u0010\u0015\u001a\u00020\tH\u0002J\u0010\u00100\u001a\u00020\u00122\u0006\u0010\u001a\u001a\u00020\u001bH\u0002R&\u0010\u0003\u001a\u001a\u0012\u0016\u0012\u0014\u0012\u0004\u0012\u00020\u0006\u0012\u0004\u0012\u00020\u0007\u0012\u0004\u0012\u00020\u00070\u00050\u0004X\u0082\u0004\u00a2\u0006\u0002\n\u0000R\u000e\u0010\b\u001a\u00020\tX\u0082\u0004\u00a2\u0006\u0002\n\u0000R\u000e\u0010\n\u001a\u00020\u000bX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\f\u001a\u00020\rX\u0082.\u00a2\u0006\u0002\n\u0000R\u0014\u0010\u000e\u001a\b\u0012\u0004\u0012\u00020\u00100\u000fX\u0082\u0004\u00a2\u0006\u0002\n\u0000\u00a8\u00062"}, d2 = {"Lcom/example/autotasks/MapActivity;", "Landroidx/appcompat/app/AppCompatActivity;", "()V", "FAMOUS_PLACES", "", "Lkotlin/Triple;", "", "", "MOSCOW_CENTER", "Lcom/yandex/mapkit/geometry/Point;", "database", "Lcom/example/autotasks/database/AppDatabase;", "mapView", "Lcom/yandex/mapkit/mapview/MapView;", "markers", "", "Lcom/yandex/mapkit/map/PlacemarkMapObject;", "addMarker", "", "title", "description", "point", "checkPermissions", "clearAllMarkers", "confirmClearMarkers", "deleteMarker", "marker", "Lcom/example/autotasks/database/MapMarker;", "loadMarkersFromDatabase", "onCreate", "savedInstanceState", "Landroid/os/Bundle;", "onRequestPermissionsResult", "requestCode", "", "permissions", "", "grantResults", "", "(I[Ljava/lang/String;[I)V", "onStart", "onStop", "setupButtons", "setupMapListener", "showAddMarkerDialog", "showAddMarkerDialogAtPoint", "showMarkerDetailsDialog", "name", "showMarkerInfo", "Companion", "app_debug"})
public final class MapActivity extends androidx.appcompat.app.AppCompatActivity {
    private com.yandex.mapkit.mapview.MapView mapView;
    private com.example.autotasks.database.AppDatabase database;
    @org.jetbrains.annotations.NotNull()
    private final java.util.List<com.yandex.mapkit.map.PlacemarkMapObject> markers = null;
    @org.jetbrains.annotations.NotNull()
    private final com.yandex.mapkit.geometry.Point MOSCOW_CENTER = null;
    @org.jetbrains.annotations.NotNull()
    private final java.util.List<kotlin.Triple<java.lang.String, java.lang.Double, java.lang.Double>> FAMOUS_PLACES = null;
    private static final int PERMISSIONS_REQUEST_CODE = 100;
    @org.jetbrains.annotations.NotNull()
    public static final com.example.autotasks.MapActivity.Companion Companion = null;
    
    public MapActivity() {
        super();
    }
    
    @java.lang.Override()
    protected void onCreate(@org.jetbrains.annotations.Nullable()
    android.os.Bundle savedInstanceState) {
    }
    
    private final void checkPermissions() {
    }
    
    @java.lang.Override()
    public void onRequestPermissionsResult(int requestCode, @org.jetbrains.annotations.NotNull()
    java.lang.String[] permissions, @org.jetbrains.annotations.NotNull()
    int[] grantResults) {
    }
    
    private final void setupButtons() {
    }
    
    private final void setupMapListener() {
    }
    
    private final void showAddMarkerDialog() {
    }
    
    private final void showAddMarkerDialogAtPoint(com.yandex.mapkit.geometry.Point point) {
    }
    
    private final void showMarkerDetailsDialog(java.lang.String name, com.yandex.mapkit.geometry.Point point) {
    }
    
    private final void addMarker(java.lang.String title, java.lang.String description, com.yandex.mapkit.geometry.Point point) {
    }
    
    private final void showMarkerInfo(com.example.autotasks.database.MapMarker marker) {
    }
    
    private final void deleteMarker(com.example.autotasks.database.MapMarker marker) {
    }
    
    private final void confirmClearMarkers() {
    }
    
    private final void clearAllMarkers() {
    }
    
    private final void loadMarkersFromDatabase() {
    }
    
    @java.lang.Override()
    protected void onStart() {
    }
    
    @java.lang.Override()
    protected void onStop() {
    }
    
    @kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000\u0012\n\u0002\u0018\u0002\n\u0002\u0010\u0000\n\u0002\b\u0002\n\u0002\u0010\b\n\u0000\b\u0086\u0003\u0018\u00002\u00020\u0001B\u0007\b\u0002\u00a2\u0006\u0002\u0010\u0002R\u000e\u0010\u0003\u001a\u00020\u0004X\u0082T\u00a2\u0006\u0002\n\u0000\u00a8\u0006\u0005"}, d2 = {"Lcom/example/autotasks/MapActivity$Companion;", "", "()V", "PERMISSIONS_REQUEST_CODE", "", "app_debug"})
    public static final class Companion {
        
        private Companion() {
            super();
        }
    }
}