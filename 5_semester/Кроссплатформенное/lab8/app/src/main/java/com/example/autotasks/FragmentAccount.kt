package com.example.autotasks

import android.os.Bundle
import android.text.TextUtils
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.Button
import android.widget.LinearLayout
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AlertDialog
import androidx.fragment.app.Fragment
import com.google.android.material.textfield.TextInputEditText
import com.google.firebase.auth.FirebaseAuth
import com.google.firebase.auth.FirebaseUser

class FragmentAccount : Fragment() {

    private lateinit var auth: FirebaseAuth
    
    private lateinit var layoutLoggedIn: LinearLayout
    private lateinit var layoutLoggedOut: LinearLayout
    private lateinit var textUserEmail: TextView
    private lateinit var textMessage: TextView
    private lateinit var editEmail: TextInputEditText
    private lateinit var editPassword: TextInputEditText
    private lateinit var editNewPassword: TextInputEditText
    private lateinit var btnLogin: Button
    private lateinit var btnRegister: Button
    private lateinit var btnLogout: Button
    private lateinit var btnChangePassword: Button
    private lateinit var btnDeleteAccount: Button

    companion object {
        fun newInstance() = FragmentAccount()
    }

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View? {
        val view = inflater.inflate(R.layout.fragment_account, container, false)
        

        auth = FirebaseAuth.getInstance()
        

        layoutLoggedIn = view.findViewById(R.id.layoutLoggedIn)
        layoutLoggedOut = view.findViewById(R.id.layoutLoggedOut)
        textUserEmail = view.findViewById(R.id.textUserEmail)
        textMessage = view.findViewById(R.id.textMessage)
        editEmail = view.findViewById(R.id.editEmail)
        editPassword = view.findViewById(R.id.editPassword)
        editNewPassword = view.findViewById(R.id.editNewPassword)
        btnLogin = view.findViewById(R.id.btnLogin)
        btnRegister = view.findViewById(R.id.btnRegister)
        btnLogout = view.findViewById(R.id.btnLogout)
        btnChangePassword = view.findViewById(R.id.btnChangePassword)
        btnDeleteAccount = view.findViewById(R.id.btnDeleteAccount)
        

        btnLogin.setOnClickListener {
            loginUser()
        }
        
        btnRegister.setOnClickListener {
            registerUser()
        }
        
        btnLogout.setOnClickListener {
            logoutUser()
        }
        
        btnChangePassword.setOnClickListener {
            changePassword()
        }
        
        btnDeleteAccount.setOnClickListener {
            confirmDeleteAccount()
        }
        

        updateUI(auth.currentUser)
        
        return view
    }
    
    override fun onStart() {
        super.onStart()

        updateUI(auth.currentUser)
    }
    
    private fun registerUser() {
        val email = editEmail.text.toString().trim()
        val password = editPassword.text.toString().trim()
        
        if (validateInput(email, password)) {
            showMessage("Регистрация...", false)
            
            auth.createUserWithEmailAndPassword(email, password)
                .addOnCompleteListener(requireActivity()) { task ->
                    if (task.isSuccessful) {
                        val user = auth.currentUser
                        showMessage("Регистрация успешна!", true)
                        updateUI(user)
                        clearFields()
                    } else {
                        val error = task.exception?.message ?: "Ошибка регистрации"
                        showMessage("Ошибка: $error", false)
                    }
                }
        }
    }
    
    private fun loginUser() {
        val email = editEmail.text.toString().trim()
        val password = editPassword.text.toString().trim()
        
        if (validateInput(email, password)) {
            showMessage("Вход...", false)
            
            auth.signInWithEmailAndPassword(email, password)
                .addOnCompleteListener(requireActivity()) { task ->
                    if (task.isSuccessful) {
                        val user = auth.currentUser
                        showMessage("Вход выполнен!", true)
                        updateUI(user)
                        clearFields()
                    } else {
                        val error = task.exception?.message ?: "Ошибка входа"
                        showMessage("Ошибка: $error", false)
                    }
                }
        }
    }
    
    private fun logoutUser() {
        auth.signOut()
        showMessage("Выход выполнен", true)
        updateUI(null)
    }
    
    private fun validateInput(email: String, password: String): Boolean {
        if (TextUtils.isEmpty(email)) {
            showMessage("Введите email", false)
            return false
        }
        
        if (TextUtils.isEmpty(password)) {
            showMessage("Введите пароль", false)
            return false
        }
        
        if (password.length < 6) {
            showMessage("Пароль должен быть не менее 6 символов", false)
            return false
        }
        
        return true
    }
    
    private fun updateUI(user: FirebaseUser?) {
        if (user != null) {
            // Пользователь залогинен
            textUserEmail.text = "Email: ${user.email}"
            layoutLoggedIn.visibility = View.VISIBLE
            layoutLoggedOut.visibility = View.GONE
        } else {
            // Пользователь не залогинен
            layoutLoggedIn.visibility = View.GONE
            layoutLoggedOut.visibility = View.VISIBLE
        }
    }
    
    private fun showMessage(message: String, isSuccess: Boolean) {
        textMessage.text = message
        textMessage.visibility = View.VISIBLE
        textMessage.setTextColor(
            if (isSuccess) {
                resources.getColor(android.R.color.holo_green_dark, null)
            } else {
                resources.getColor(android.R.color.holo_red_dark, null)
            }
        )
        

        textMessage.postDelayed({
            if (textMessage.text == message) {
                textMessage.visibility = View.GONE
            }
        }, 3000)
    }
    
    private fun changePassword() {
        val newPassword = editNewPassword.text.toString().trim()
        
        if (TextUtils.isEmpty(newPassword)) {
            showMessage("Введите новый пароль", false)
            return
        }
        
        if (newPassword.length < 6) {
            showMessage("Пароль должен быть не менее 6 символов", false)
            return
        }
        
        val user = auth.currentUser
        if (user != null) {
            showMessage("Изменение пароля...", false)
            
            user.updatePassword(newPassword)
                .addOnCompleteListener { task ->
                    if (task.isSuccessful) {
                        showMessage("Пароль успешно изменен!", true)
                        editNewPassword.text?.clear()
                    } else {
                        val error = task.exception?.message ?: "Ошибка изменения пароля"
                        showMessage("Ошибка: $error", false)
                    }
                }
        } else {
            showMessage("Пользователь не авторизован", false)
        }
    }
    
    private fun confirmDeleteAccount() {
        AlertDialog.Builder(requireContext())
            .setTitle("Удаление аккаунта")
            .setMessage("Вы уверены, что хотите удалить аккаунт? Это действие необратимо!")
            .setPositiveButton("Да, удалить") { _, _ ->
                deleteAccount()
            }
            .setNegativeButton("Отмена", null)
            .show()
    }
    
    private fun deleteAccount() {
        val user = auth.currentUser
        if (user != null) {
            showMessage("Удаление аккаунта...", false)
            
            user.delete()
                .addOnCompleteListener { task ->
                    if (task.isSuccessful) {
                        showMessage("Аккаунт успешно удален", true)
                        updateUI(null)
                    } else {
                        val error = task.exception?.message ?: "Ошибка удаления аккаунта"
                        showMessage("Ошибка: $error", false)
                        

                        if (error.contains("recent login", ignoreCase = true)) {
                            showMessage("Выйдите и войдите снова, затем повторите попытку", false)
                        }
                    }
                }
        } else {
            showMessage("Пользователь не авторизован", false)
        }
    }
    
    private fun clearFields() {
        editEmail.text?.clear()
        editPassword.text?.clear()
    }
}
