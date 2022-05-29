import com.mongodb.client.MongoCollection
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream


class Controller(var workbook: XSSFWorkbook, val teachersMongo: MongoCollection<Teacher>)
{
    fun createTimeTable() {
        val table = workbook.createSheet("Teachers").apply {
            defaultColumnWidth = 15
            defaultRowHeight = 1500
        }
        val teachers = teachersMongo.find().toList();

        val rows = Array(teachers.size+3) { i -> table.createRow(i) }
        rows.forEach { row ->
            for(i in 0..61){
                row.createCell(i)
            }
        }

        val centerCellStyle = workbook.createCellStyle().apply {
            verticalAlignment = VerticalAlignment.CENTER
            alignment = HorizontalAlignment.CENTER
            wrapText = true
            borderLeft = BorderStyle.THIN
            borderTop = BorderStyle.THIN
            borderRight = BorderStyle.THIN
            borderBottom = BorderStyle.THIN
        }
        val centerCellStyleLines = workbook.createCellStyle().apply {
            verticalAlignment = VerticalAlignment.CENTER
            alignment = HorizontalAlignment.CENTER
            wrapText = true
            borderLeft = BorderStyle.MEDIUM
            borderTop = BorderStyle.MEDIUM
            borderRight = BorderStyle.MEDIUM
            borderBottom = BorderStyle.MEDIUM
        }

        table.addMergedRegion(CellRangeAddress(0, 2, 0, 1))

        table.addMergedRegion(CellRangeAddress(0 ,0, 2, 31))
        table.addMergedRegion(CellRangeAddress(0, 0, 32, 61))
        rows[0].getCell(2).apply { cellStyle = centerCellStyle }.setCellValue("ЧЕТНАЯ НЕДЕЛЯ")
        rows[0].getCell(32).apply { cellStyle = centerCellStyle }.setCellValue("НЕЧЕТНАЯ НЕДЕЛЯ")

        val daysNames = arrayOf("Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота")
        val lessonNames = arrayOf(
            "1-я пара\n08:00-09:30",
            "2-я пара\n09:45-11:15",
            "3-я пара\n11:30-13:00",
            "4-я пара\n13:55-15:25",
            "5-я пара\n15:40-17:15")

        for (i in daysNames.indices){
            table.addMergedRegion(CellRangeAddress(1, 1, 2+5*i, 1+5*(i+1)))
            table.addMergedRegion(CellRangeAddress(1, 1, 2+5*i+30, 1+5*(i+1)+30))

            rows[1].getCell(2+5*i).apply { cellStyle = centerCellStyle }.setCellValue(daysNames[i])
            rows[1].getCell(30+(2+5*i)).apply { cellStyle = centerCellStyle }.setCellValue(daysNames[i])

            for (j in lessonNames.indices){
                rows[2].getCell(2+j+5*i).apply { cellStyle = centerCellStyle }.setCellValue(lessonNames[j])
                rows[2].getCell(30+(2+j)+5*i).apply { cellStyle = centerCellStyle }.setCellValue(lessonNames[j])
            }
        }

        teachers.forEachIndexed { i, teacher ->
            table.addMergedRegion(CellRangeAddress(i+3 ,i+3, 0, 1))
            rows[i+3].getCell(0).apply { cellStyle = centerCellStyle }.setCellValue(teacher.name)

            teacher.timeTable.oddWeek.forEachIndexed{ j, day ->
                day.lessons.forEachIndexed{ k, lesson ->
                    var lessonType = lesson?.type ?: ""
                    if(!lessonType.isEmpty()) { lessonType += ". " }
                    var lessonName = lesson?.name ?: ""
                    if(!lessonName.isEmpty()) { lessonName += "\n" }
                    val lessonHousing = lesson?.housing ?: ""
                    val lessonAudience = lesson?.audience ?: ""
                    var lessonAddress = ""
                    if(!lessonAudience.isEmpty()) { lessonAddress = "$lessonHousing-$lessonAudience\n" }

                    val lessonGroup = lesson?.group?.joinToString(separator = ",", postfix = "", prefix = "") ?: ""

                    val lessonInfo = "$lessonType$lessonName$lessonAddress$lessonGroup"
                    rows[i+3].getCell(30+(2+j*5+k)).apply { cellStyle = centerCellStyle }.setCellValue(lessonInfo)
                }
            }
            teacher.timeTable.evenWeek.forEachIndexed{ j, day ->
                day.lessons.forEachIndexed{ k, lesson ->
                    var lessonType = lesson?.type ?: ""
                    if(!lessonType.isEmpty()) { lessonType += ". " }
                    var lessonName = lesson?.name ?: ""
                    if(!lessonName.isEmpty()) { lessonName += "\n" }
                    val lessonHousing = lesson?.housing ?: ""
                    val lessonAudience = lesson?.audience ?: ""
                    var lessonAddress = ""
                    if(!lessonAudience.isEmpty()) { lessonAddress = "$lessonHousing-$lessonAudience\n" }

                    val lessonGroup = lesson?.group?.joinToString(separator = ",", postfix = "", prefix = "") ?: ""

                    val lessonInfo = "$lessonType$lessonName$lessonAddress$lessonGroup"

                    rows[i+3].getCell(2+j*5+k).apply { cellStyle = centerCellStyle }.setCellValue(lessonInfo)
                }
            }
        }

        for (i in rows.indices){
            for (j in 0..1){
                rows[i].getCell(j).apply { cellStyle = centerCellStyleLines }
            }
        }

        for (i in 0..2){
            for (j in 0 until rows[0].lastCellNum){
                rows[i].getCell(j).apply { cellStyle = centerCellStyleLines }
            }
        }

        workbook.write(FileOutputStream("./src/main/resources/Table.xlsx"))
    }
}
