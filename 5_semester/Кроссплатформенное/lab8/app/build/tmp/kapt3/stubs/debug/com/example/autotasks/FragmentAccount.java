package com.example.autotasks;

@kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000j\n\u0002\u0018\u0002\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0002\b\u0005\n\u0002\u0018\u0002\n\u0002\b\u0003\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0002\b\u0002\n\u0002\u0010\u0002\n\u0002\b\u0006\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0000\n\u0002\u0018\u0002\n\u0002\b\u0004\n\u0002\u0010\u000e\n\u0000\n\u0002\u0010\u000b\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0002\b\u0005\u0018\u0000 12\u00020\u0001:\u00011B\u0005\u00a2\u0006\u0002\u0010\u0002J\b\u0010\u0015\u001a\u00020\u0016H\u0002J\b\u0010\u0017\u001a\u00020\u0016H\u0002J\b\u0010\u0018\u001a\u00020\u0016H\u0002J\b\u0010\u0019\u001a\u00020\u0016H\u0002J\b\u0010\u001a\u001a\u00020\u0016H\u0002J\b\u0010\u001b\u001a\u00020\u0016H\u0002J&\u0010\u001c\u001a\u0004\u0018\u00010\u001d2\u0006\u0010\u001e\u001a\u00020\u001f2\b\u0010 \u001a\u0004\u0018\u00010!2\b\u0010\"\u001a\u0004\u0018\u00010#H\u0016J\b\u0010$\u001a\u00020\u0016H\u0016J\b\u0010%\u001a\u00020\u0016H\u0002J\u0018\u0010&\u001a\u00020\u00162\u0006\u0010\'\u001a\u00020(2\u0006\u0010)\u001a\u00020*H\u0002J\u0012\u0010+\u001a\u00020\u00162\b\u0010,\u001a\u0004\u0018\u00010-H\u0002J\u0018\u0010.\u001a\u00020*2\u0006\u0010/\u001a\u00020(2\u0006\u00100\u001a\u00020(H\u0002R\u000e\u0010\u0003\u001a\u00020\u0004X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0005\u001a\u00020\u0006X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0007\u001a\u00020\u0006X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\b\u001a\u00020\u0006X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\t\u001a\u00020\u0006X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\n\u001a\u00020\u0006X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u000b\u001a\u00020\fX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\r\u001a\u00020\fX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u000e\u001a\u00020\fX\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u000f\u001a\u00020\u0010X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0011\u001a\u00020\u0010X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0012\u001a\u00020\u0013X\u0082.\u00a2\u0006\u0002\n\u0000R\u000e\u0010\u0014\u001a\u00020\u0013X\u0082.\u00a2\u0006\u0002\n\u0000\u00a8\u00062"}, d2 = {"Lcom/example/autotasks/FragmentAccount;", "Landroidx/fragment/app/Fragment;", "()V", "auth", "Lcom/google/firebase/auth/FirebaseAuth;", "btnChangePassword", "Landroid/widget/Button;", "btnDeleteAccount", "btnLogin", "btnLogout", "btnRegister", "editEmail", "Lcom/google/android/material/textfield/TextInputEditText;", "editNewPassword", "editPassword", "layoutLoggedIn", "Landroid/widget/LinearLayout;", "layoutLoggedOut", "textMessage", "Landroid/widget/TextView;", "textUserEmail", "changePassword", "", "clearFields", "confirmDeleteAccount", "deleteAccount", "loginUser", "logoutUser", "onCreateView", "Landroid/view/View;", "inflater", "Landroid/view/LayoutInflater;", "container", "Landroid/view/ViewGroup;", "savedInstanceState", "Landroid/os/Bundle;", "onStart", "registerUser", "showMessage", "message", "", "isSuccess", "", "updateUI", "user", "Lcom/google/firebase/auth/FirebaseUser;", "validateInput", "email", "password", "Companion", "app_debug"})
public final class FragmentAccount extends androidx.fragment.app.Fragment {
    private com.google.firebase.auth.FirebaseAuth auth;
    private android.widget.LinearLayout layoutLoggedIn;
    private android.widget.LinearLayout layoutLoggedOut;
    private android.widget.TextView textUserEmail;
    private android.widget.TextView textMessage;
    private com.google.android.material.textfield.TextInputEditText editEmail;
    private com.google.android.material.textfield.TextInputEditText editPassword;
    private com.google.android.material.textfield.TextInputEditText editNewPassword;
    private android.widget.Button btnLogin;
    private android.widget.Button btnRegister;
    private android.widget.Button btnLogout;
    private android.widget.Button btnChangePassword;
    private android.widget.Button btnDeleteAccount;
    @org.jetbrains.annotations.NotNull()
    public static final com.example.autotasks.FragmentAccount.Companion Companion = null;
    
    public FragmentAccount() {
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
    
    @java.lang.Override()
    public void onStart() {
    }
    
    private final void registerUser() {
    }
    
    private final void loginUser() {
    }
    
    private final void logoutUser() {
    }
    
    private final boolean validateInput(java.lang.String email, java.lang.String password) {
        return false;
    }
    
    private final void updateUI(com.google.firebase.auth.FirebaseUser user) {
    }
    
    private final void showMessage(java.lang.String message, boolean isSuccess) {
    }
    
    private final void changePassword() {
    }
    
    private final void confirmDeleteAccount() {
    }
    
    private final void deleteAccount() {
    }
    
    private final void clearFields() {
    }
    
    @kotlin.Metadata(mv = {1, 9, 0}, k = 1, xi = 48, d1 = {"\u0000\u0012\n\u0002\u0018\u0002\n\u0002\u0010\u0000\n\u0002\b\u0002\n\u0002\u0018\u0002\n\u0000\b\u0086\u0003\u0018\u00002\u00020\u0001B\u0007\b\u0002\u00a2\u0006\u0002\u0010\u0002J\u0006\u0010\u0003\u001a\u00020\u0004\u00a8\u0006\u0005"}, d2 = {"Lcom/example/autotasks/FragmentAccount$Companion;", "", "()V", "newInstance", "Lcom/example/autotasks/FragmentAccount;", "app_debug"})
    public static final class Companion {
        
        private Companion() {
            super();
        }
        
        @org.jetbrains.annotations.NotNull()
        public final com.example.autotasks.FragmentAccount newInstance() {
            return null;
        }
    }
}