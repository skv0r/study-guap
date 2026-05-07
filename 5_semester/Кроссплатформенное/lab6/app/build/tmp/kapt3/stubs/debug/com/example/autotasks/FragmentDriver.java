package com.example.autotasks;

@kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000n\n\u0002\u0018\u0002\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010!\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0002\u0018\u0002\n\u0002\u0010\u000e\n\u0002\u0010\u0000\n\u0000\n\u0002\u0018\u0002\n\u0002\u0010\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010\t\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010\b\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0002\b\u0002\u0018\u0000 \"2\u00020\u0001:\u0001\"Bg\u0012\f\u0010\u0002\u001a\b\u0012\u0004\u0012\u00020\u00040\u0003\u0012\"\u0010\u0005\u001a\u001e\b\u0001\u0012\u0004\u0012\u00020\u0004\u0012\n\u0012\b\u0012\u0004\u0012\u00020\b0\u0007\u0012\u0006\u0012\u0004\u0018\u00010\t0\u0006\u0012\u0012\u0010\n\u001a\u000e\u0012\u0004\u0012\u00020\u0004\u0012\u0004\u0012\u00020\f0\u000b\u0012\f\u0010\r\u001a\b\u0012\u0004\u0012\u00020\f0\u000e\u0012\f\u0010\u000f\u001a\b\u0012\u0004\u0012\u00020\u00100\u000e\u00a2\u0006\u0002\u0010\u0011J&\u0010\u001a\u001a\u0004\u0018\u00010\u001b2\u0006\u0010\u001c\u001a\u00020\u001d2\b\u0010\u001e\u001a\u0004\u0018\u00010\u001f2\b\u0010 \u001a\u0004\u0018\u00010!H\u0016J\u0006\u0010\r\u001a\u00020\fR\u000e\u0010\u0012\u001a\u00020\u0013X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0014\u001a\u00020\u0015X\u0082\u000e\u00a2\u0006\u0002\n\u0000R\u0014\u0010\u0002\u001a\b\u0012\u0004\u0012\u00020\u00040\u0003X\u0082\u0004\u00a2\u0006\u0002\n\u0000R\u0014\u0010\u000f\u001a\b\u0012\u0004\u0012\u00020\u00100\u000eX\u0082\u0004\u00a2\u0006\u0002\n\u0000R,\u0010\u0005\u001a\u001e\b\u0001\u0012\u0004\u0012\u00020\u0004\u0012\n\u0012\b\u0012\u0004\u0012\u00020\b0\u0007\u0012\u0006\u0012\u0004\u0018\u00010\t0\u0006X\u0082\u0004\u00a2\u0006\u0004\n\u0002\u0010\u0016R\u0014\u0010\r\u001a\b\u0012\u0004\u0012\u00020\f0\u000eX\u0082\u0004\u00a2\u0006\u0002\n\u0000R\u001a\u0010\n\u001a\u000e\u0012\u0004\u0012\u00020\u0004\u0012\u0004\u0012\u00020\f0\u000bX\u0082\u0004\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0017\u001a\u00020\u0018X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0019\u001a\u00020\u0018X\u0082.\u00a2\u0006\u0002\n\u0000\u00a8\u0006#"}, d2 = {"Lcom/example/autotasks/FragmentDriver;", "Landroidx/fragment/app/Fragment;", "drivers", "", "Lcom/example/autotasks/database/Driver;", "loadDriverInfo", "Lkotlin/Function2;", "Lkotlin/coroutines/Continuation;", "", "", "showDriverActionsDialog", "Lkotlin/Function1;", "", "refreshDrivers", "Lkotlin/Function0;", "getLastLoadTime", "", "(Ljava/util/List;Lkotlin/jvm/functions/Function2;Lkotlin/jvm/functions/Function1;Lkotlin/jvm/functions/Function0;Lkotlin/jvm/functions/Function0;)V", "adapter", "Lcom/example/autotasks/DriversAdapter;", "currentDriverIndex", "", "Lkotlin/jvm/functions/Function2;", "textLastUpdate", "Landroid/widget/TextView;", "textResult", "onCreateView", "Landroid/view/View;", "inflater", "Landroid/view/LayoutInflater;", "container", "Landroid/view/ViewGroup;", "savedInstanceState", "Landroid/os/Bundle;", "Companion", "app_debug"})
public final class FragmentDriver extends androidx.fragment.app.Fragment {
    @org.jetbrains.annotations.NotNull()
    private final java.util.List<com.example.autotasks.database.Driver> drivers = null;
    @org.jetbrains.annotations.NotNull()
    private final kotlin.jvm.functions.Function2<com.example.autotasks.database.Driver, kotlin.coroutines.Continuation<? super java.lang.String>, java.lang.Object> loadDriverInfo = null;
    @org.jetbrains.annotations.NotNull()
    private final kotlin.jvm.functions.Function1<com.example.autotasks.database.Driver, kotlin.Unit> showDriverActionsDialog = null;
    @org.jetbrains.annotations.NotNull()
    private final kotlin.jvm.functions.Function0<kotlin.Unit> refreshDrivers = null;
    @org.jetbrains.annotations.NotNull()
    private final kotlin.jvm.functions.Function0<java.lang.Long> getLastLoadTime = null;
    private int currentDriverIndex = 0;
    private com.example.autotasks.DriversAdapter adapter;
    private android.widget.TextView textResult;
    private android.widget.TextView textLastUpdate;
    @org.jetbrains.annotations.NotNull()
    public static final com.example.autotasks.FragmentDriver.Companion Companion = null;
    
    public FragmentDriver(@org.jetbrains.annotations.NotNull()
    java.util.List<com.example.autotasks.database.Driver> drivers, @org.jetbrains.annotations.NotNull()
    kotlin.jvm.functions.Function2<? super com.example.autotasks.database.Driver, ? super kotlin.coroutines.Continuation<? super java.lang.String>, ? extends java.lang.Object> loadDriverInfo, @org.jetbrains.annotations.NotNull()
    kotlin.jvm.functions.Function1<? super com.example.autotasks.database.Driver, kotlin.Unit> showDriverActionsDialog, @org.jetbrains.annotations.NotNull()
    kotlin.jvm.functions.Function0<kotlin.Unit> refreshDrivers, @org.jetbrains.annotations.NotNull()
    kotlin.jvm.functions.Function0<java.lang.Long> getLastLoadTime) {
        super();
    }
    
    @java.lang.Override()
    @org.jetbrains.annotations.Nullable()
    public android.view.View onCreateView(@org.jetbrains.annotations.NotNull()
    android.view.LayoutInflater inflater, @org.jetbrains.annotations.Nullable()
    android.view.ViewGroup container, @org.jetbrains.annotations.Nullable()
    android.os.Bundle savedInstanceState) {
        return null;
    }
    
    public final void refreshDrivers() {
    }
    
    @kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000B\n\u0002\u0018\u0002\n\u0002\u0010\u0000\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010!\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0002\u0018\u0002\n\u0002\u0010\u000e\n\u0000\n\u0002\u0018\u0002\n\u0002\u0010\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010\t\n\u0002\b\u0002\b\u0086\u0003\u0018\u00002\u00020\u0001B\u0007\b\u0002\u00a2\u0006\u0002\u0010\u0002Jm\u0010\u0003\u001a\u00020\u00042\f\u0010\u0005\u001a\b\u0012\u0004\u0012\u00020\u00070\u00062\"\u0010\b\u001a\u001e\b\u0001\u0012\u0004\u0012\u00020\u0007\u0012\n\u0012\b\u0012\u0004\u0012\u00020\u000b0\n\u0012\u0006\u0012\u0004\u0018\u00010\u00010\t2\u0012\u0010\f\u001a\u000e\u0012\u0004\u0012\u00020\u0007\u0012\u0004\u0012\u00020\u000e0\r2\f\u0010\u000f\u001a\b\u0012\u0004\u0012\u00020\u000e0\u00102\f\u0010\u0011\u001a\b\u0012\u0004\u0012\u00020\u00120\u0010\u00a2\u0006\u0002\u0010\u0013\u00a8\u0006\u0014"}, d2 = {"Lcom/example/autotasks/FragmentDriver$Companion;", "", "()V", "newInstance", "Lcom/example/autotasks/FragmentDriver;", "drivers", "", "Lcom/example/autotasks/database/Driver;", "loadDriverInfo", "Lkotlin/Function2;", "Lkotlin/coroutines/Continuation;", "", "showDriverActionsDialog", "Lkotlin/Function1;", "", "refreshDrivers", "Lkotlin/Function0;", "getLastLoadTime", "", "(Ljava/util/List;Lkotlin/jvm/functions/Function2;Lkotlin/jvm/functions/Function1;Lkotlin/jvm/functions/Function0;Lkotlin/jvm/functions/Function0;)Lcom/example/autotasks/FragmentDriver;", "app_debug"})
    public static final class Companion {
        
        private Companion() {
            super();
        }
        
        @org.jetbrains.annotations.NotNull()
        public final com.example.autotasks.FragmentDriver newInstance(@org.jetbrains.annotations.NotNull()
        java.util.List<com.example.autotasks.database.Driver> drivers, @org.jetbrains.annotations.NotNull()
        kotlin.jvm.functions.Function2<? super com.example.autotasks.database.Driver, ? super kotlin.coroutines.Continuation<? super java.lang.String>, ? extends java.lang.Object> loadDriverInfo, @org.jetbrains.annotations.NotNull()
        kotlin.jvm.functions.Function1<? super com.example.autotasks.database.Driver, kotlin.Unit> showDriverActionsDialog, @org.jetbrains.annotations.NotNull()
        kotlin.jvm.functions.Function0<kotlin.Unit> refreshDrivers, @org.jetbrains.annotations.NotNull()
        kotlin.jvm.functions.Function0<java.lang.Long> getLastLoadTime) {
            return null;
        }
    }
}