import com.google.gson.Gson
import com.google.gson.GsonBuilder
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import java.io.File
import java.io.FileInputStream

private const val FAKULTET_NAME = "ИМОП"
private const val TIMETABLE_FILE = "C:\\timetable\\$FAKULTET_NAME.xls"
private val myExcelBook = HSSFWorkbook(FileInputStream(TIMETABLE_FILE))
private const val DELTA_BETWEEN_FIRST_LAST_DAY_ROW = 70
val gson: Gson = GsonBuilder().serializeNulls().create()

fun main() {
//    val sheet = myExcelBook.getSheetAt(0)
//    readSheet(sheet)
    readFile(myExcelBook)
}

private fun readFile(hssfWorkbook: HSSFWorkbook) {
    for (i in 0 until hssfWorkbook.numberOfSheets) {
        readSheet(hssfWorkbook.getSheetAt(i))
    }
}

private fun readSheet(sheet: Sheet) {
    val firstGroupRow = findFirstRow(sheet)

    // проходим по столбцам, если попадается пустой, то переходим к следующему листу
    for (columnIndex in 2..500) {
        if (sheet.getRow(firstGroupRow)?.getCell(columnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)?.stringCellValue == null) continue
        readWeek(sheet = sheet, startRow = firstGroupRow + 1, columnIndex = columnIndex)
    }
}

// Возвращаем номер строки, в которой перечислены номера групп
private fun findFirstRow(sheet: Sheet): Int {
    for (rowIndex in 0..30) {
        if (sheet.getRow(rowIndex)?.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)?.stringCellValue == "группа") {
            return rowIndex
        }
    }
    throw Exception("Не найдена ячейка \"группа\"")
}

private fun readWeek(sheet: Sheet, startRow: Int, columnIndex: Int) {
    // получаем название группы
    val numberOfGroup = sheet.getRow(findFirstRow(sheet)).getCell(columnIndex).stringCellValue
    val days = mutableListOf<TimetableDay>()
    for (startRowIndex in startRow..startRow + DELTA_BETWEEN_FIRST_LAST_DAY_ROW step 14) {
        days.add(readDay(sheet = sheet, startRow = startRowIndex, column = columnIndex))
    }
    val strDay = gson.toJson(GroupTimetable(numberOfGroup, days))
    println(numberOfGroup)
    File("C:\\timetable\\$FAKULTET_NAME").mkdirs()
    File("C:\\timetable\\$FAKULTET_NAME\\${numberOfGroup}").writeText(strDay, Charsets.UTF_16)
}

private fun readDay(sheet: Sheet, startRow: Int, column: Int): TimetableDay {
    val day = sheet.getRow(startRow).getCell(0).stringCellValue
    println(day)
    val lessonsList = mutableListOf<PairKlass>()

    for (i in startRow..startRow + 12 step 2) {
        lessonsList.add(getPairKlass(sheet, i, column))
    }

    return TimetableDay(day, lessonsList)
}

private fun getPairKlass(sheet: Sheet, row: Int, column: Int): PairKlass {
    println("Строчка - $row, столбец - $column")
    val number = sheet.getRow(row).getCell(1).stringCellValue.toInt()

    var firstCell = sheet.getRow(row).getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
    var secondCell = sheet.getRow(row + 1).getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
    getMergedRegion(firstCell)?.let {
        firstCell = sheet.getRow(it.firstRow).getCell(it.firstColumn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)
    }
    getMergedRegion(secondCell)?.let {
        secondCell = sheet.getRow(it.firstRow).getCell(it.firstColumn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)
        if (secondCell == firstCell || (secondCell.stringCellValue.isEmpty() && firstCell.stringCellValue.isEmpty())) {
            secondCell = null
        }
    }
    return PairKlass(
        ClassNumber.getNumber(number).time,
        firstCell?.stringCellValue,
        secondCell?.stringCellValue
    ).cleanBlank()
}

private fun PairKlass.cleanBlank(): PairKlass {
    return this.copy(
        time = time,
        odd = if (odd == "" && even == "") null else odd,
        even = if (even == "" && odd == "") null else even
    )
}

private fun getMergedRegion(cell: Cell): CellRangeAddress? {
    val sheet: Sheet = cell.sheet
    for (range in sheet.mergedRegions) {
        if (range.isInRange(cell.rowIndex, cell.columnIndex)) {
            return range
        }
    }
    return null
}