package com.example.autotasks

import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import androidx.recyclerview.widget.RecyclerView

class DriversAdapter(
    private val drivers: List<Pair<String, Int>>,
    private val onItemClick: (Int) -> Unit
) : RecyclerView.Adapter<DriversAdapter.DriverViewHolder>() {

    inner class DriverViewHolder(view: View) : RecyclerView.ViewHolder(view) {
        val textDriverName: TextView = view.findViewById(R.id.textDriverName)
    }

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): DriverViewHolder {
        val view = LayoutInflater.from(parent.context)
            .inflate(R.layout.item_driver, parent, false)
        return DriverViewHolder(view)
    }

    override fun onBindViewHolder(holder: DriverViewHolder, position: Int) {
        holder.textDriverName.text = drivers[position].first
        holder.itemView.setOnClickListener { onItemClick(position) }
    }

    override fun getItemCount(): Int = drivers.size
}
