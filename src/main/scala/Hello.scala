import java.io._

import org.apache.poi.ss.util._
import org.apache.poi.xssf.usermodel._

object Hello extends App {
 // println("Hello, World!")

  val columnHeaders = List(
    "Дата перерахунку коштів",
    "Сума принятих коштів",
    "Сума утриманої винагороди",
    "Сума перерахованих коштів",
    "Сервіс",
    "ПІБ Покупця",
    "№ ЕН НП",
    "Номер замовлення"
  )

  val data = Map(
    (10,"1")   -> List("07.07.2021", "1739.00", "",  "1739.00", "Безготівковий 6",	"Крупко Зоя Олеговна",	"20400236713873",	"РОЗ121293402"),
    (11,"2")   -> List("07.07.2021", "1239.00", "",  "1239.00", "Безготівковий 6",	"Перфілова Наталя Сергіївна",	"20400236713873",	"РОЗ121293402"),
    (12,"3")   -> List("07.07.2021", "1439.00", "",  "1539.00", "Безготівковий 6",	"Крупко Зоя Олеговна",	"20400236713873",	"РОЗ121293402"),
    (13,"4")   -> List("07.07.2021", "1439.00", "",  "1539.00", "Безготівковий 6",	"Кокош Генріетта Іванівна",	"20400236713873",	"РОЗ121293402")
  )


  // Open template
  val wb = new XSSFWorkbook(
    getClass.getResourceAsStream("template6.xlsx")
  )
  val sheet = wb.getSheetAt(0)

  // Fill headers
  val headerRow = sheet.getRow(8)
    for(idx <- 1 to columnHeaders.length) {
      headerRow.getCell(idx).setCellValue(columnHeaders(idx - 1))
    }


  // Fill body
  for(key @ (rowNumber, rowName) <- data.keys) {
    val row = sheet.getRow(rowNumber)
    row.getCell(0).setCellValue(rowName)

    for(idx <- 1 to data(key).length) {
      row.getCell(idx).setCellValue(data(key)(idx - 1))
    }
  }

  // Save report
  val resultFile = new FileOutputStream("report.xlsx")
  wb.write(resultFile)
  resultFile.close()

}