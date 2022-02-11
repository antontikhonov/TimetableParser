import com.google.gson.Gson
import com.google.gson.GsonBuilder
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import java.io.File
import java.io.FileInputStream

private const val TIMETABLE_FILE = "C:\\timetable\\М.xls"
private val myExcelBook = HSSFWorkbook(FileInputStream(TIMETABLE_FILE))
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
    val START_FIRST_GROUP_ROW = findFirstRow(sheet)

    for (i in 2..30) {
        if (sheet.getRow(START_FIRST_GROUP_ROW)?.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)?.stringCellValue == null) continue
        readWeek(sheet, START_FIRST_GROUP_ROW + 1, i)
    }
}

private fun findFirstRow(sheet: Sheet): Int {
    var z = 0
    for (i in 0..30) {
        if (sheet.getRow(i)?.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)?.stringCellValue == "группа") {
            z = i
            return z
        }
    }
    return z
}

private fun readWeek(sheet: Sheet, startRow: Int, column: Int) {
    val numberOfGroup = sheet.getRow(findFirstRow(sheet)).getCell(column).stringCellValue
    val days = mutableListOf<TimetableDay>()
    for (i in startRow..startRow + 70 step 14) {
        days.add(readDay(sheet, i, column))
    }
    val strDay = gson.toJson(GroupTimetable(numberOfGroup, days))
    File("C:\\timetable\\М\\${numberOfGroup}").writeText(strDay, Charsets.UTF_16)
}

private fun readDay(sheet: Sheet, startRow: Int, column: Int): TimetableDay {
    val day = sheet.getRow(startRow).getCell(0).stringCellValue
    val classes = mutableListOf<PairKlass>()
    for (i in startRow..startRow + 12 step 2) {
        classes.add(getPairKlass(sheet, i, column))
    }
    return TimetableDay(day, classes)
}

private fun getPairKlass(sheet: Sheet, row: Int, column: Int): PairKlass {
    val number = sheet.getRow(row).getCell(1).stringCellValue.toInt()

    var firstCell = sheet.getRow(row).getCell(column)
    var secondCell = sheet.getRow(row + 1).getCell(column)
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