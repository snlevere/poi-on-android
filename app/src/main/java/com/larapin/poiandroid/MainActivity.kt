package com.larapin.poiandroid

import android.os.Bundle
import android.os.Environment
import android.support.v7.app.AppCompatActivity
import android.widget.Toast
import com.larapin.poiandroid.databinding.ActivityMainBinding
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {
    private lateinit var binding: ActivityMainBinding

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        binding = ActivityMainBinding.inflate(layoutInflater)
        val view = binding.root
        setContentView(view)
        
        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl")
        System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl")
        System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl")

        val path = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS)

        binding.btnWriteXlsx.setOnClickListener {
            writeXlsx(path, binding.inputText.text.toString().trim())
        }
        binding.btnReadXlsx.setOnClickListener{
            readXlsx(path, binding.inputText.text.toString().trim())
        }
    }

    private fun writeXlsx(path: File, message: String) {
        try {
            val workbook = XSSFWorkbook()

            val outputStream = FileOutputStream(File(path, "/poi.xlsx"))

            val sheet = workbook.createSheet("Sheet 1")
            val row = sheet.createRow(2)
            val fields: List<String> = message.split(",")
            for ((index, field) in fields.withIndex()) {
                val cell = row.createCell(index)
                cell.setCellValue(field)
            }

            workbook.write(outputStream)
            outputStream.close()
            Toast.makeText(this, "poi.xlsx was successfully created", Toast.LENGTH_SHORT).show()
        }catch (e: Exception){
            e.printStackTrace()
        }
    }

    private fun readXlsx(path: File, message: String) {
        try {
            val workbook = XSSFWorkbook(File(path, "/poi.xlsx"))

            val sheet = workbook.getSheetAt(0)
            val rowIterator = sheet.iterator()
            val sb = StringBuilder()
            while (rowIterator.hasNext()) {
                val row = rowIterator.next()
                val cellIterator = row.cellIterator()
                while (cellIterator.hasNext()) {
                    val cell = cellIterator.next()
                    sb.append(cell.stringCellValue)
                }
                sb.append('\n')
            }

            Toast.makeText(this, "Read [$sb]", Toast.LENGTH_SHORT).show()
        }catch (e: Exception){
            e.printStackTrace()
        }
    }
}
