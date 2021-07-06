package com.wzl.searchexcel

import android.os.Bundle
import android.view.View
import android.widget.EditText
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.IOException

class MainActivity : AppCompatActivity() {

    private lateinit var sheet: XSSFSheet
    private lateinit var rowEditText: EditText
    private lateinit var colEditText: EditText
    private lateinit var resultTextView: TextView

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        rowEditText = findViewById(R.id.row)
        colEditText = findViewById(R.id.col)
        resultTextView = findViewById(R.id.result)

        try {
            val inputStream = assets.open("test.xlsx")
            val xssfWorkbook = XSSFWorkbook(inputStream)
            sheet = xssfWorkbook.getSheetAt(0)
        } catch (e: IOException) {
            e.printStackTrace()
        }

        findViewById<View>(R.id.query).setOnClickListener {
            try {
                val rowText = rowEditText.editableText.toString()
                val colText = colEditText.editableText.toString()
                if (rowText.isEmpty() || colText.isEmpty()) {
                    Toast.makeText(this, "请设置行列值", Toast.LENGTH_SHORT).show()
                    return@setOnClickListener
                }
                val row = rowText.toInt()
                val col = colText.toInt()
                // 表格实际从0坐标开始
                query(row - 1, col - 1)
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }
    }

    private fun query(row: Int, col: Int) {
        val sheetRow = sheet.getRow(row)
        if (sheetRow == null) {
            Toast.makeText(this, "越界", Toast.LENGTH_SHORT).show()
            return
        }
        val cell = sheetRow.getCell(col)
        if (cell == null) {
            Toast.makeText(this, "越界", Toast.LENGTH_SHORT).show()
            return
        }
        val cellValue = cell.numericCellValue
        resultTextView.text = "$cellValue"
    }
}