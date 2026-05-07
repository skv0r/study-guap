package com.example.autotasks;

@kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000^\n\u0002\u0018\u0002\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010!\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010\u0002\n\u0000\n\u0002\u0010\u000e\n\u0002\b\u0007\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010\u000b\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0002\b\u000b\u0018\u0000 ,2\u00020\u0001:\u0001,B\u0005\u00a2\u0006\u0002\u0010\u0002J \u0010\f\u001a\u00020\r2\u0006\u0010\u000e\u001a\u00020\u000f2\u0006\u0010\u0010\u001a\u00020\u000f2\u0006\u0010\u0011\u001a\u00020\u0004H\u0002J\b\u0010\u0012\u001a\u00020\rH\u0002J\b\u0010\u0013\u001a\u00020\rH\u0002J\b\u0010\u0014\u001a\u00020\rH\u0002J\u0010\u0010\u0015\u001a\u00020\r2\u0006\u0010\u0016\u001a\u00020\u0017H\u0002J\b\u0010\u0018\u001a\u00020\u0019H\u0002J\b\u0010\u001a\u001a\u00020\rH\u0002J&\u0010\u001b\u001a\u0004\u0018\u00010\u001c2\u0006\u0010\u001d\u001a\u00020\u001e2\b\u0010\u001f\u001a\u0004\u0018\u00010 2\b\u0010!\u001a\u0004\u0018\u00010\"H\u0016J\b\u0010#\u001a\u00020\rH\u0016J\b\u0010$\u001a\u00020\rH\u0016J\b\u0010%\u001a\u00020\rH\u0016J\u001a\u0010&\u001a\u00020\r2\u0006\u0010\'\u001a\u00020\u001c2\b\u0010!\u001a\u0004\u0018\u00010\"H\u0016J\u0010\u0010(\u001a\u00020\r2\u0006\u0010\'\u001a\u00020\u001cH\u0002J\b\u0010)\u001a\u00020\rH\u0002J\u0010\u0010*\u001a\u00020\r2\u0006\u0010\u0011\u001a\u00020\u0004H\u0002J\u0010\u0010+\u001a\u00020\r2\u0006\u0010\u0016\u001a\u00020\u0017H\u0002R\u000e\u0010\u0003\u001a\u00020\u0004X\u0082\u0004\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0005\u001a\u00020\u0006X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0007\u001a\u00020\bX\u0082.\u00a2\u0006\u0002\n\u0000R\u0014\u0010\t\u001a\b\u0012\u0004\u0012\u00020\u000b0\nX\u0082\u0004\u00a2\u0006\u0002\n\u0000\u00a8\u0006-"}, d2 = {"Lcom/example/autotasks/FragmentMap;", "Landroidx/fragment/app/Fragment;", "()V", "MOSCOW_CENTER", "Lcom/yandex/mapkit/geometry/Point;", "database", "Lcom/example/autotasks/database/AppDatabase;", "mapView", "Lcom/yandex/mapkit/mapview/MapView;", "markers", "", "Lcom/yandex/mapkit/map/PlacemarkMapObject;", "addMarker", "", "title", "", "description", "point", "checkPermissions", "clearAllMarkers", "confirmClearMarkers", "deleteMarker", "marker", "Lcom/example/autotasks/database/MapMarker;", "isInternetAvailable", "", "loadMarkersFromDatabase", "onCreateView", "Landroid/view/View;", "inflater", "Landroid/view/LayoutInflater;", "container", "Landroid/view/ViewGroup;", "savedInstanceState", "Landroid/os/Bundle;", "onDestroyView", "onStart", "onStop", "onViewCreated", "view", "setupButtons", "setupMapListener", "showAddMarkerDialogAtPoint", "showMarkerInfo", "Companion", "app_debug"})
public final class FragmentMap extends androidx.fragment.app.Fragment {
    private com.yandex.mapkit.mapview.MapView mapView;
    private com.example.autotasks.database.AppDatabase database;
    @org.jetbrains.annotations.NotNull()
    private final java.util.List<com.yandex.mapkit.map.PlacemarkMapObject> markers = null;
    @org.jetbrains.annotations.NotNull()
    private final com.yandex.mapkit.geometry.Point MOSCOW_CENTER = null;
    private static final int PERMISSIONS_REQUEST_CODE = 100;
    @org.jetbrains.annotations.NotNull()
    public static final com.example.autotasks.FragmentMap.Companion Companion = null;
    
    public FragmentMap() {
        super();
    }
    
    private final boolean isInternetAvailable() {
        return false;
    }
    
    @java.lang.Override()
    @org.jetbrains.annotations.Nullable()
    public android.view.View onCreateView(@org.jetbrains.annotations.NotNull()
    android.view.LayoutInflater inflater, @org.jetbrains.annotations.Nullable()
    android.view.ViewGroup container, @org.jetbrains.annotations.Nullable()
    android.os.Bundle savedInstanceState) {
        return null;
    }
    
    @java.lang.Override()
    public void onViewCreated(@org.jetbrains.annotations.NotNull()
    android.view.View view, @org.jetbrains.annotations.Nullable()
    android.os.Bundle savedInstanceState) {
    }
    
    private final void checkPermissions() {
    }
    
    private final void setupButtons(android.view.View view) {
    }
    
    private final void setupMapListener() {
    }
    
    private final void showAddMarkerDialogAtPoint(com.yandex.mapkit.geometry.Point point) {
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
    public void onStart() {
    }
    
    @java.lang.Override()
    public void onStop() {
    }
    
    @java.lang.Override()
    public void onDestroyView() {
    }
    
    @kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000\u0018\n\u0002\u0018\u0002\n\u0002\u0010\u0000\n\u0002\b\u0002\n\u0002\u0010\b\n\u0000\n\u0002\u0018\u0002\n\u0000\b\u0086\u0003\u0018\u00002\u00020\u0001B\u0007\b\u0002\u00a2\u0006\u0002\u0010\u0002J\u0006\u0010\u0005\u001a\u00020\u0006R\u000e\u0010\u0003\u001a\u00020\u0004X\u0082T\u00a2\u0006\u0002\n\u0000\u00a8\u0006\u0007"}, d2 = {"Lcom/example/autotasks/FragmentMap$Companion;", "", "()V", "PERMISSIONS_REQUEST_CODE", "", "newInstance", "Lcom/example/autotasks/FragmentMap;", "app_debug"})
    public static final class Companion {
        
        private Companion() {
            super();
        }
        
        @org.jetbrains.annotations.NotNull()
        public final com.example.autotasks.FragmentMap newInstance() {
            return null;
        }
    }
}