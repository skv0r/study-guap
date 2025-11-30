package com.example.autotasks

import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import androidx.fragment.app.Fragment

class FragmentAccount : Fragment() {

    companion object {
        fun newInstance() = FragmentAccount()
    }

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View? {
        val view = inflater.inflate(R.layout.fragment_car, container, false)
        val textView = view.findViewById<TextView>(R.id.textCar)

        // Показываем сообщение о регистрации
        textView.text = "скоро будет регистрация"

        return view
    }
}

