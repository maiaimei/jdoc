package cn.maiaimei.jdocs.excel;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

public class POIExcelUtils {

  /**
   * Creates a cell and set the style for the cell.
   *
   * @param row    the row to create the cell in
   * @param column the column number to create the cell in
   */
  public static void createCell(Row row, int column, double value, CellStyle cellStyle) {
    Cell cell = row.createCell(column);
    cell.setCellValue(value);
    cell.setCellStyle(cellStyle);
  }

  /**
   * Creates a cell and set the style for the cell.
   *
   * @param row    the row to create the cell in
   * @param column the column number to create the cell in
   */
  public static void createCell(Row row, int column, Date value, CellStyle cellStyle) {
    Cell cell = row.createCell(column);
    cell.setCellValue(value);
    cell.setCellStyle(cellStyle);
  }

  /**
   * Creates a cell and set the style for the cell.
   *
   * @param row    the row to create the cell in
   * @param column the column number to create the cell in
   */
  public static void createCell(Row row, int column, LocalDateTime value, CellStyle cellStyle) {
    Cell cell = row.createCell(column);
    cell.setCellValue(value);
    cell.setCellStyle(cellStyle);
  }

  /**
   * Creates a cell and set the style for the cell.
   *
   * @param row    the row to create the cell in
   * @param column the column number to create the cell in
   */
  public static void createCell(Row row, int column, LocalDate value, CellStyle cellStyle) {
    Cell cell = row.createCell(column);
    cell.setCellValue(value);
    cell.setCellStyle(cellStyle);
  }

  /**
   * Creates a cell and set the style for the cell.
   *
   * @param row    the row to create the cell in
   * @param column the column number to create the cell in
   */
  public static void createCell(Row row, int column, Calendar value, CellStyle cellStyle) {
    Cell cell = row.createCell(column);
    cell.setCellValue(value);
    cell.setCellStyle(cellStyle);
  }

  /**
   * Creates a cell and set the style for the cell.
   *
   * @param row    the row to create the cell in
   * @param column the column number to create the cell in
   */
  public static void createCell(Row row, int column, RichTextString value, CellStyle cellStyle) {
    Cell cell = row.createCell(column);
    cell.setCellValue(value);
    cell.setCellStyle(cellStyle);
  }

  /**
   * Creates a cell and set the style for the cell.
   *
   * @param row    the row to create the cell in
   * @param column the column number to create the cell in
   */
  public static void createCell(Row row, int column, String value, CellStyle cellStyle) {
    Cell cell = row.createCell(column);
    cell.setCellValue(value);
    cell.setCellStyle(cellStyle);
  }

  /**
   * Creates a cellStyle
   *
   * @param wb     the workbook
   * @param halign the horizontal alignment for the cell.
   * @param valign the vertical alignment for the cell.
   * @return the new Cell Style object
   */
  public static CellStyle createCellStyle(Workbook wb,
      HorizontalAlignment halign, VerticalAlignment valign) {
    CellStyle cellStyle = wb.createCellStyle();
    cellStyle.setAlignment(halign);
    cellStyle.setVerticalAlignment(valign);
    return cellStyle;
  }


  /**
   * Creates a cellStyle
   *
   * @param wb         the workbook
   * @param halign     the horizontal alignment for the cell.
   * @param valign     the vertical alignment for the cell.
   * @param dataFormat the data format (must be a valid format)
   * @return the new Cell Style object
   */
  public static CellStyle createCellStyle(Workbook wb,
      HorizontalAlignment halign, VerticalAlignment valign, String dataFormat) {
    CreationHelper createHelper = wb.getCreationHelper();
    final CellStyle cellStyle = createCellStyle(wb, halign, valign);
    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(dataFormat));
    return cellStyle;
  }

  /**
   * Creates a cellStyle
   *
   * @param wb          the workbook
   * @param halign      the horizontal alignment for the cell.
   * @param valign      the vertical alignment for the cell.
   * @param borderStyle the border type
   * @param borderColor the border color
   * @return the new Cell Style object
   */
  public static CellStyle createCellStyle(Workbook wb,
      HorizontalAlignment halign, VerticalAlignment valign, BorderStyle borderStyle,
      short borderColor) {
    CellStyle cellStyle = wb.createCellStyle();
    cellStyle.setAlignment(halign);
    cellStyle.setVerticalAlignment(valign);
    cellStyle.setBorderTop(borderStyle);
    cellStyle.setBorderRight(borderStyle);
    cellStyle.setBorderBottom(borderStyle);
    cellStyle.setBorderLeft(borderStyle);
    cellStyle.setTopBorderColor(borderColor);
    cellStyle.setRightBorderColor(borderColor);
    cellStyle.setBottomBorderColor(borderColor);
    cellStyle.setLeftBorderColor(borderColor);
    return cellStyle;
  }

  /**
   * Creates a cellStyle
   *
   * @param wb          the workbook
   * @param halign      the horizontal alignment for the cell.
   * @param valign      the vertical alignment for the cell.
   * @param borderStyle the border type
   * @param borderColor the border color
   * @param dataFormat  the data format (must be a valid format)
   * @return the new Cell Style object
   */
  public static CellStyle createCellStyle(Workbook wb,
      HorizontalAlignment halign, VerticalAlignment valign,
      BorderStyle borderStyle, short borderColor,
      String dataFormat) {
    CreationHelper createHelper = wb.getCreationHelper();
    CellStyle cellStyle = createCellStyle(wb, halign, valign, borderStyle, borderColor);
    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(dataFormat));
    return cellStyle;
  }

  public static List<Map<String, Object>> readFile(String path,
      Integer sheetIndex, String sheetName, List<String> headers, int rowStart,
      int columnStart) {
    List<Map<String, Object>> list = new ArrayList<>();
    try (Workbook wb = WorkbookFactory.create(Files.newInputStream(Paths.get(path)))) {
      Sheet sheet = getSheet(wb, sheetName, sheetIndex);
      rowStart = rowStart == -1 ? sheet.getFirstRowNum() : rowStart;
      int rowEnd = sheet.getLastRowNum();
      for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++) {
        Row row = sheet.getRow(rowNum);
        if (row == null) {
          continue;
        }
        int i = 0;
        Map<String, Object> item = new LinkedHashMap<>();
        columnStart = columnStart == -1 ? row.getFirstCellNum() : columnStart;
        int columnEnd = row.getLastCellNum();
        for (int cellNum = columnStart; cellNum < columnEnd; cellNum++) {
          Cell cell = row.getCell(cellNum);
          if (cell != null) {
            String key;
            if (headers == null || StringUtils.isBlank(headers.get(i))) {
              CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
              key = cellRef.formatAsString();
            } else {
              key = headers.get(i);
            }
            item.put(key, getCellValue(cell));
            i++;
          }
        }
        list.add(item);
      }
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
    return list;
  }

  private static Object getCellValue(Cell cell) {
    final CellType cellType = cell.getCellType();
    if (cellType == CellType.STRING) {
      return cell.getRichStringCellValue().getString();
    }
    if (cellType == CellType.NUMERIC) {
      if (DateUtil.isCellDateFormatted(cell)) {
        return cell.getDateCellValue();
      } else {
        return cell.getNumericCellValue();
      }
    }
    if (cellType == CellType.BOOLEAN) {
      return cell.getBooleanCellValue();
    }
    if (cellType == CellType.FORMULA) {
      return cell.getCellFormula();
    }
    return null;
  }

  private static Sheet getSheet(Workbook wb, String sheetName, Integer sheetIndex) {
    if (StringUtils.isNotBlank(sheetName)) {
      return wb.getSheet(sheetName);
    }
    if (sheetIndex != null && sheetIndex >= 0) {
      return wb.getSheetAt(sheetIndex);
    }
    return wb.getSheetAt(0);
  }

  private POIExcelUtils() {
    throw new UnsupportedOperationException();
  }
}
