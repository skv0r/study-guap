package com.example.autotasks

import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import androidx.fragment.app.Fragment
import androidx.lifecycle.lifecycleScope
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import kotlinx.coroutines.launch

class FragmentDriver(
    private val drivers: List<Pair<String, Int>>,
    private val loadDriverInfo: suspend (Int) -> String
) : Fragment() {

    private var currentDriverIndex = 0

    companion object {
        fun newInstance(
            drivers: List<Pair<String, Int>>,
            loadDriverInfo: suspend (Int) -> String
        ): FragmentDriver {
            return FragmentDriver(drivers, loadDriverInfo)
        }
    }

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View? {
        val view = inflater.inflate(R.layout.fragment_driver, container, false)

        val recyclerDrivers = view.findViewById<RecyclerView>(R.id.recyclerDrivers)
        val textResult = view.findViewById<TextView>(R.id.textResult)

        recyclerDrivers.layoutManager = LinearLayoutManager(requireContext())
        val adapter = DriversAdapter(drivers) { index ->
            currentDriverIndex = index
            lifecycleScope.launch {
                textResult.text = loadDriverInfo(drivers[index].second)
                (activity as MainActivity).currentDriverIndex = index
            }
        }
        recyclerDrivers.adapter = adapter

        if (drivers.isNotEmpty()) {
            lifecycleScope.launch {
                textResult.text = loadDriverInfo(drivers[currentDriverIndex].second)
            }
        }

        return view
    }
}
