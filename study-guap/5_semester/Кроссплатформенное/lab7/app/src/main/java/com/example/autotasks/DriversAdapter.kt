package com.example.autotasks

import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import androidx.recyclerview.widget.RecyclerView
import com.example.autotasks.database.Driver

class DriversAdapter(
    private var drivers: MutableList<Driver>,
    private val onItemClick: (Driver) -> Unit,
    private val onItemLongClick: (Driver) -> Unit
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
        val driver = drivers[position]
        holder.textDriverName.text = "${driver.fullName} (#${driver.driverNumber})"
        
        holder.itemView.setOnClickListener { 
            onItemClick(driver) 
        }
        
        holder.itemView.setOnLongClickListener {
            onItemLongClick(driver)
            true
        }
    }

    override fun getItemCount(): Int = drivers.size
    
    fun updateDrivers(newDrivers: List<Driver>) {
        android.util.Log.d("DriversAdapter", "updateDrivers: обновление с ${newDrivers.size} гонщиками")
        android.util.Log.d("DriversAdapter", "updateDrivers: ID новых гонщиков: ${newDrivers.map { it.id }}")
        
        // ВАЖНО: создаем копию списка, чтобы избежать проблем с ссылками
        val driversCopy = newDrivers.toList()
        
        val oldSize = drivers.size
        drivers.clear()
        android.util.Log.d("DriversAdapter", "updateDrivers: список очищен, старый размер: $oldSize")
        
        drivers.addAll(driversCopy)
        android.util.Log.d("DriversAdapter", "updateDrivers: список обновлен, новый размер: ${drivers.size}")
        
        if (drivers.size != driversCopy.size) {
            android.util.Log.e("DriversAdapter", "ОШИБКА: размер не совпадает! Ожидалось: ${driversCopy.size}, получено: ${drivers.size}")
        }
        
        notifyDataSetChanged()
        android.util.Log.d("DriversAdapter", "updateDrivers: notifyDataSetChanged вызван")
    }
}
