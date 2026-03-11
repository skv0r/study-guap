package com.example.autotasks

import android.net.Uri
import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.Button
import android.widget.CheckBox
import android.widget.EditText
import android.widget.ImageView
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView

class MainActivity : AppCompatActivity() {

    private lateinit var etTask: EditText
    private lateinit var btnAdd: Button
    private lateinit var recyclerView: RecyclerView
    private val taskList = ArrayList<Task>()
    private lateinit var adapter: TaskAdapter

    // хранит индекс задачи, к которой прикрепляется фото
    private var currentAttachIndex: Int = -1

    private val pickImageLauncher =
        registerForActivityResult(ActivityResultContracts.GetContent()) { uri: Uri? ->
            uri?.let {
                if (currentAttachIndex in taskList.indices) {
                    taskList[currentAttachIndex].imageUri = it
                    adapter.notifyItemChanged(currentAttachIndex)
                }
            }
        }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        etTask = findViewById(R.id.etTask)
        btnAdd = findViewById(R.id.btnAdd)
        recyclerView = findViewById(R.id.recyclerView)

        adapter = TaskAdapter(taskList)
        recyclerView.layoutManager = LinearLayoutManager(this)
        recyclerView.adapter = adapter

        btnAdd.setOnClickListener {
            val text = etTask.text.toString().trim()
            if (text.isNotEmpty()) {
                taskList.add(Task(text))
                adapter.notifyItemInserted(taskList.size - 1)
                etTask.text.clear()
            }
        }
    }

    data class Task(var title: String, var done: Boolean = false, var imageUri: Uri? = null)

    inner class TaskAdapter(private val items: ArrayList<Task>) :
        RecyclerView.Adapter<TaskAdapter.TaskViewHolder>() {

        inner class TaskViewHolder(view: View) : RecyclerView.ViewHolder(view) {
            val imgPhoto: ImageView = view.findViewById(R.id.imgPhoto)
            val checkBox: CheckBox = view.findViewById(R.id.checkBox)
            val btnAttach: Button = view.findViewById(R.id.btnAttach)
            val btnDelete: Button = view.findViewById(R.id.btnDelete)
        }

        override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): TaskViewHolder {
            val view = LayoutInflater.from(parent.context)
                .inflate(R.layout.item_task, parent, false)
            return TaskViewHolder(view)
        }

        override fun onBindViewHolder(holder: TaskViewHolder, position: Int) {
            val task = items[position]
            holder.checkBox.text = task.title
            holder.checkBox.isChecked = task.done

            // если фото прикреплено — отображаем
            if (task.imageUri != null) {
                holder.imgPhoto.setImageURI(task.imageUri)
            } else {
                holder.imgPhoto.setImageResource(android.R.drawable.ic_menu_camera)
            }

            holder.checkBox.setOnCheckedChangeListener { _, isChecked ->
                task.done = isChecked
            }

            holder.btnAttach.setOnClickListener {
                currentAttachIndex = holder.adapterPosition
                pickImageLauncher.launch("image/*")
            }

            holder.btnDelete.setOnClickListener {
                val pos = holder.adapterPosition
                items.removeAt(pos)
                notifyItemRemoved(pos)
            }
        }

        override fun getItemCount(): Int = items.size
    }
}
