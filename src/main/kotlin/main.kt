import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.litote.kmongo.*

fun main() {
    val dbTeachers = mongoDatabase.getCollection<Teacher>()

    val workbook = XSSFWorkbook()

    val controller = Controller(workbook, dbTeachers)
    controller.createTimeTable()
    dbTeachers.find().toList().forEach { println(it) }
}
