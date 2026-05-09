package com.example.autotasks;

@kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u00008\n\u0002\u0018\u0002\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0000\n\u0002\u0010\b\n\u0000\n\u0002\u0018\u0002\n\u0002\b\t\n\u0002\u0010\u0002\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0002\b\u0002\u0018\u00002\u00020\u0001B\u0005\u00a2\u0006\u0002\u0010\u0002J\b\u0010\u0014\u001a\u00020\u0015H\u0002J\u0012\u0010\u0016\u001a\u00020\u00152\b\u0010\u0017\u001a\u0004\u0018\u00010\u0018H\u0014J\b\u0010\u0019\u001a\u00020\u0015H\u0002R\u000e\u0010\u0003\u001a\u00020\u0004X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0005\u001a\u00020\u0004X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0006\u001a\u00020\u0007X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\b\u001a\u00020\tX\u0082\u000e\u00a2\u0006\u0002\n\u0000R\u000e\u0010\n\u001a\u00020\u000bX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\f\u001a\u00020\u000bX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\r\u001a\u00020\u000bX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u000e\u001a\u00020\u000bX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u000f\u001a\u00020\u000bX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0010\u001a\u00020\u000bX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0011\u001a\u00020\u000bX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0012\u001a\u00020\u000bX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0013\u001a\u00020\u000bX\u0082.\u00a2\u0006\u0002\n\u0000\u00a8\u0006\u001a"}, d2 = {"Lcom/example/autotasks/EditDriverActivity;", "Landroidx/appcompat/app/AppCompatActivity;", "()V", "btnCancel", "Landroid/widget/Button;", "btnUpdate", "database", "Lcom/example/autotasks/database/AppDatabase;", "driverId", "", "editBroadcastName", "Landroid/widget/EditText;", "editCountryCode", "editDriverNumber", "editFirstName", "editFullName", "editLastName", "editNameAcronym", "editTeamColour", "editTeamName", "loadDriverData", "", "onCreate", "savedInstanceState", "Landroid/os/Bundle;", "updateDriver", "app_debug"})
public final class EditDriverActivity extends androidx.appcompat.app.AppCompatActivity {
    private android.widget.EditText editFullName;
    private android.widget.EditText editDriverNumber;
    private android.widget.EditText editFirstName;
    private android.widget.EditText editLastName;
    private android.widget.EditText editTeamName;
    private android.widget.EditText editTeamColour;
    private android.widget.EditText editNameAcronym;
    private android.widget.EditText editCountryCode;
    private android.widget.EditText editBroadcastName;
    private android.widget.Button btnUpdate;
    private android.widget.Button btnCancel;
    private com.example.autotasks.database.AppDatabase database;
    private int driverId = -1;
    
    public EditDriverActivity() {
        super();
    }
    
    @java.lang.Override()
    protected void onCreate(@org.jetbrains.annotations.Nullable()
    android.os.Bundle savedInstanceState) {
    }
    
    private final void loadDriverData() {
    }
    
    private final void updateDriver() {
    }
}