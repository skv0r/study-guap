package com.example.autotasks

import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import androidx.fragment.app.Fragment
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.launch

class FragmentCar(
    private val loadCarInfo: suspend (Int) -> String,
    private val driverIndex: Int,
    private val driverName: String
) : Fragment() {

    companion object {
        fun newInstance(loadCarInfo: suspend (Int) -> String, driverIndex: Int, driverName: String) =
            FragmentCar(loadCarInfo, driverIndex, driverName)
    }

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View? {
        val view = inflater.inflate(R.layout.fragment_car, container, false)
        val textCar = view.findViewById<TextView>(R.id.textCar)

        lifecycleScope.launch {
            val carInfo = loadCarInfo((activity as MainActivity).drivers[driverIndex].second)
            textCar.text = "Гонщик: $driverName\n$carInfo"
        }

        return view
    }
}

