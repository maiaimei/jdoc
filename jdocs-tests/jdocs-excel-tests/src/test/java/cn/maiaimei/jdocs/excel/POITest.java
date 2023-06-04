package cn.maiaimei.jdocs.excel;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.bazaarvoice.jolt.Chainr;
import com.bazaarvoice.jolt.JsonUtils;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

@Slf4j
public class POITest {

  private static final String ROOT_PATH = "E:\\app\\tmp\\";
  private static final String XLS = "xls";
  private static final String XLSX = "xlsx";

  @Test
  public void testNewWorkbook() throws IOException {
    try (Workbook wb = new HSSFWorkbook();
        OutputStream os = Files.newOutputStream(Paths.get(generateFileName(XLS)))) {
      wb.write(os);
    }

    try (Workbook wb = new XSSFWorkbook();
        OutputStream os = Files.newOutputStream(Paths.get(generateFileName(XLSX)))) {
      wb.write(os);
    }
  }

  @Test
  public void testNewSheet() throws IOException {
    try (Workbook wb = new HSSFWorkbook();
        OutputStream os = Files.newOutputStream(Paths.get(generateFileName(XLS)))) {
      String safeSheetName = WorkbookUtil.createSafeSheetName(
          "[O'Brien's sales*?]"); // returns " O'Brien's sales   "
      wb.createSheet(safeSheetName);
      wb.createSheet("MySheet1");
      wb.write(os);
    }

    try (Workbook wb = new XSSFWorkbook();
        OutputStream os = Files.newOutputStream(Paths.get(generateFileName(XLSX)))) {
      String safeSheetName = WorkbookUtil.createSafeSheetName(
          "[Purchase Order*?]"); // returns "Purchase Order"
      wb.createSheet(safeSheetName);
      wb.createSheet("MySheet1");
      wb.write(os);
    }
  }

  @Test
  public void testWrite() throws IOException {
    try (Workbook wb = new XSSFWorkbook();
        OutputStream os = Files.newOutputStream(Paths.get(generateFileName(XLSX)))) {
      // create cell styles
      final CellStyle cellStyle1 = POIExcelUtils.createCellStyle(wb, HorizontalAlignment.CENTER,
          VerticalAlignment.CENTER,
          BorderStyle.THIN, IndexedColors.BLACK.getIndex());
      final CellStyle cellStyle2 = POIExcelUtils.createCellStyle(wb, HorizontalAlignment.LEFT,
          VerticalAlignment.CENTER,
          BorderStyle.THIN, IndexedColors.BLACK.getIndex());
      final CellStyle cellStyle3 = POIExcelUtils.createCellStyle(wb, HorizontalAlignment.CENTER,
          VerticalAlignment.CENTER,
          BorderStyle.THIN, IndexedColors.BLACK.getIndex(), "yyyy-MM-dd");
      // create sheet
      final Sheet sheet = wb.createSheet();
      // create table header row
      final Row headerRow = sheet.createRow(0);
      POIExcelUtils.createCell(headerRow, 0, "序号", cellStyle1);
      POIExcelUtils.createCell(headerRow, 1, "采购时间", cellStyle3);
      POIExcelUtils.createCell(headerRow, 2, "商品编号", cellStyle1);
      POIExcelUtils.createCell(headerRow, 3, "商品名称", cellStyle2);
      POIExcelUtils.createCell(headerRow, 4, "商品单价", cellStyle1);
      POIExcelUtils.createCell(headerRow, 5, "采购数量", cellStyle1);
      POIExcelUtils.createCell(headerRow, 6, "金额", cellStyle1);
      POIExcelUtils.createCell(headerRow, 7, "供应商", cellStyle2);
      // mock data
      List<PurchaseOrderItem> list = new ArrayList<>();
      final Date date = new Date();
      list.add(
          new PurchaseOrderItem(date, "01", "01商品", 5.00, 100000, 500000.00, "广州市A有限公司"));
      list.add(new PurchaseOrderItem(date, "02", "02商品", 2.00, 20, 40.00, "B公司"));
      list.add(new PurchaseOrderItem(date, "03", "03商品", 6.00, 5, 30.00, "B公司"));
      list.add(new PurchaseOrderItem(date, "04", "04商品", 4.00, 15, 60.00, "A公司"));
      list.add(new PurchaseOrderItem(date, "05", "05商品", 8.00, 10, 80.00, "A公司"));
      // create table body row and set cell value
      for (int i = 0; i < list.size(); i++) {
        PurchaseOrderItem item = list.get(i);
        final Row row = sheet.createRow(i + 1);
        POIExcelUtils.createCell(row, 0, i + 1, cellStyle1);
        POIExcelUtils.createCell(row, 1, item.getPurchaseTime(), cellStyle3);
        POIExcelUtils.createCell(row, 2, item.getItemNum(), cellStyle1);
        POIExcelUtils.createCell(row, 3, item.getItemName(), cellStyle2);
        POIExcelUtils.createCell(row, 4, item.getUnitPrice(), cellStyle1);
        POIExcelUtils.createCell(row, 5, item.getCount(), cellStyle1);
        POIExcelUtils.createCell(row, 6, item.getAmount(), cellStyle1);
        POIExcelUtils.createCell(row, 7, item.getSupplier(), cellStyle2);
      }
      // Adjust column width to fit the contents
      sheet.autoSizeColumn(1);
      wb.write(os);
    }
  }

  @Test
  public void testReadNoHeaders() throws JsonProcessingException {
    String path = ROOT_PATH.concat("20230604143626248.xlsx");
    final List<Map<String, Object>> maps = POIExcelUtils.readFile(path, 0, null, null, 1, -1);
    assertTrue(maps.size() > 0);

    List<Object> spec = JsonUtils.classpathToList("/template/import-purchase-order.json");
    Chainr chainr = Chainr.fromSpec(spec);
    Object output = chainr.transform(maps);
    assertNotNull(output);

    final ObjectMapper objectMapper = new ObjectMapper();
    final List<PurchaseOrderItem> purchaseOrderItems = objectMapper.readValue(
        objectMapper.writeValueAsString(output),
        new TypeReference<>() {
        });
    assertTrue(purchaseOrderItems.size() > 0);

    final List<PurchaseOrderItem> transform = cn.maiaimei.jdocs.core.JsonUtils.transform(
        "/template/import-purchase-order.json", maps, new TypeReference<>() {
        });
    assertTrue(transform.size() > 0);
  }

  @Test
  public void testReadWithHeaders() {
    String path = ROOT_PATH.concat("20230604143626248.xlsx");
    List<String> headers = List.of("serialNum", "purchaseTime", "itemNum", "itemName",
        "unitPrice", "count",
        "amount", "supplier");
    final List<Map<String, Object>> maps = POIExcelUtils.readFile(path, 0, null, headers, 1, -1);
    assertTrue(maps.size() > 0);
  }

  private String generateFileName(String ext) {
    final SimpleDateFormat format = new SimpleDateFormat("yyyyMMddHHmmssSSS");
    return ROOT_PATH.concat(format.format(new Date())).concat(".").concat(ext);
  }
}
