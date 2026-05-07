package com.example.autotasks;

/**
 * Задание 2: Оптимизация перебора массива
 *
 * Демонстрация различных способов оптимизации перебора массивов в Kotlin
 * с тестированием на массивах от 100000 элементов
 */
@kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000@\n\u0002\u0018\u0002\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0010\b\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010\u0002\n\u0000\n\u0002\u0010\u000e\n\u0000\n\u0002\u0010 \n\u0002\b\u0004\n\u0002\u0018\u0002\n\u0002\b\f\u0018\u00002\u00020\u0001B\u0005\u00a2\u0006\u0002\u0010\u0002J\u0010\u0010\u000b\u001a\u00020\f2\u0006\u0010\r\u001a\u00020\u000eH\u0002J\u001c\u0010\u000f\u001a\b\u0012\u0004\u0012\u00020\u00040\u00102\u0006\u0010\u0011\u001a\u00020\u0004H\u0082@\u00a2\u0006\u0002\u0010\u0012J\u0012\u0010\u0013\u001a\u00020\f2\b\u0010\u0014\u001a\u0004\u0018\u00010\u0015H\u0014J\b\u0010\u0016\u001a\u00020\fH\u0002J\u001c\u0010\u0017\u001a\u00020\f2\f\u0010\u0018\u001a\b\u0012\u0004\u0012\u00020\u00040\u0010H\u0082@\u00a2\u0006\u0002\u0010\u0019J\u001c\u0010\u001a\u001a\u00020\f2\f\u0010\u0018\u001a\b\u0012\u0004\u0012\u00020\u00040\u0010H\u0082@\u00a2\u0006\u0002\u0010\u0019J\u001c\u0010\u001b\u001a\u00020\f2\f\u0010\u0018\u001a\b\u0012\u0004\u0012\u00020\u00040\u0010H\u0082@\u00a2\u0006\u0002\u0010\u0019J\u001c\u0010\u001c\u001a\u00020\f2\f\u0010\u0018\u001a\b\u0012\u0004\u0012\u00020\u00040\u0010H\u0082@\u00a2\u0006\u0002\u0010\u0019J\u001c\u0010\u001d\u001a\u00020\f2\f\u0010\u0018\u001a\b\u0012\u0004\u0012\u00020\u00040\u0010H\u0082@\u00a2\u0006\u0002\u0010\u0019J\u001c\u0010\u001e\u001a\u00020\f2\f\u0010\u0018\u001a\b\u0012\u0004\u0012\u00020\u00040\u0010H\u0082@\u00a2\u0006\u0002\u0010\u0019J\u001c\u0010\u001f\u001a\u00020\f2\f\u0010\u0018\u001a\b\u0012\u0004\u0012\u00020\u00040\u0010H\u0082@\u00a2\u0006\u0002\u0010\u0019J\u0016\u0010 \u001a\u00020\f2\u0006\u0010\u0011\u001a\u00020\u0004H\u0082@\u00a2\u0006\u0002\u0010\u0012R\u000e\u0010\u0003\u001a\u00020\u0004X\u0082D\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0005\u001a\u00020\u0006X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0007\u001a\u00020\bX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\t\u001a\u00020\nX\u0082.\u00a2\u0006\u0002\n\u0000\u00a8\u0006!"}, d2 = {"Lcom/example/autotasks/ArrayOptimizationActivity;", "Landroidx/appcompat/app/AppCompatActivity;", "()V", "ARRAY_SIZE", "", "btnTest", "Landroid/widget/Button;", "scrollView", "Landroid/widget/ScrollView;", "textResults", "Landroid/widget/TextView;", "appendText", "", "text", "", "createSourceArray", "", "size", "(ILkotlin/coroutines/Continuation;)Ljava/lang/Object;", "onCreate", "savedInstanceState", "Landroid/os/Bundle;", "runOptimizationTests", "test1_SimpleForLoop", "source", "(Ljava/util/List;Lkotlin/coroutines/Continuation;)Ljava/lang/Object;", "test2_ForEach", "test3_Map", "test4_FilterMap", "test5_Sequence", "test6_SequenceOptimized", "test7_ParallelProcessing", "test8_IntArrayVsList", "app_debug"})
public final class ArrayOptimizationActivity extends androidx.appcompat.app.AppCompatActivity {
    private android.widget.TextView textResults;
    private android.widget.ScrollView scrollView;
    private android.widget.Button btnTest;
    private final int ARRAY_SIZE = 100000;
    
    public ArrayOptimizationActivity() {
        super();
    }
    
    @java.lang.Override()
    protected void onCreate(@org.jetbrains.annotations.Nullable()
    android.os.Bundle savedInstanceState) {
    }
    
    private final void runOptimizationTests() {
    }
    
    private final java.lang.Object createSourceArray(int size, kotlin.coroutines.Continuation<? super java.util.List<java.lang.Integer>> $completion) {
        return null;
    }
    
    private final java.lang.Object test1_SimpleForLoop(java.util.List<java.lang.Integer> source, kotlin.coroutines.Continuation<? super kotlin.Unit> $completion) {
        return null;
    }
    
    private final java.lang.Object test2_ForEach(java.util.List<java.lang.Integer> source, kotlin.coroutines.Continuation<? super kotlin.Unit> $completion) {
        return null;
    }
    
    private final java.lang.Object test3_Map(java.util.List<java.lang.Integer> source, kotlin.coroutines.Continuation<? super kotlin.Unit> $completion) {
        return null;
    }
    
    private final java.lang.Object test4_FilterMap(java.util.List<java.lang.Integer> source, kotlin.coroutines.Continuation<? super kotlin.Unit> $completion) {
        return null;
    }
    
    private final java.lang.Object test5_Sequence(java.util.List<java.lang.Integer> source, kotlin.coroutines.Continuation<? super kotlin.Unit> $completion) {
        return null;
    }
    
    private final java.lang.Object test6_SequenceOptimized(java.util.List<java.lang.Integer> source, kotlin.coroutines.Continuation<? super kotlin.Unit> $completion) {
        return null;
    }
    
    private final java.lang.Object test7_ParallelProcessing(java.util.List<java.lang.Integer> source, kotlin.coroutines.Continuation<? super kotlin.Unit> $completion) {
        return null;
    }
    
    private final java.lang.Object test8_IntArrayVsList(int size, kotlin.coroutines.Continuation<? super kotlin.Unit> $completion) {
        return null;
    }
    
    private final void appendText(java.lang.String text) {
    }
}