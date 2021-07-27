import com.norbitltd.spoiwo.model.{Row, Sheet}
import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.model.enums._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions.XlsxSheet
import java.time.LocalDate

object Hello extends App {

  val titleRow = Row(index = 0).withCellValues("РЕЄСТР ПЕРЕКАЗІВ")
  val secondRow = Row(index = 2).withCellValues("Одержувач:", "Товариство з обмеженою відповідальністю '???'")
  val thirdRow = Row(index = 4).withCellValues("П//р в банку:", "UA08798723920jjhsd00909809")
  val fourthRow = Row(index = 6).withCellValues("Дата:", LocalDate.of(1867, 11, 7))


  val reportSheet: Sheet = Sheet(name = "Test sheet")
    .withRows(
      titleRow,
      secondRow,
      thirdRow,
      fourthRow,
      Row(index = 8).withCellValues("№", "Дата перерахунку коштів", "Сума принятих коштів", "Сума утриманої винагороди", "Сума перерахованих коштів", "Сервіс", "ПІБ Покупця", "№ ЕН НП", "Номер замовлення"),
      Row(index = 9).withCellValues(1, LocalDate.of(1867, 11, 7), BigDecimal(1739), BigDecimal(0), BigDecimal(1739), "Безготівковий 6", "ПІБ", "12345678987654", "РОЗ012345678"),
      Row(index = 10).withCellValues(2, LocalDate.of(2020, 7, 7), BigDecimal(1339), BigDecimal(0), BigDecimal(1339), "Безготівковий 6", "Ініціали", "12345678987654", "РОЗ012345678"),
    )
    .withColumns(
      Column(index = 0, autoSized = true)
    )

  reportSheet.saveAsXlsx("report.xlsx")

}