package cn.maiaimei.jdocs.excel;

import cn.maiaimei.jdocs.core.BeetlUtils;
import java.util.List;
import java.util.Map;

public final class ExcelUtils {

  public static List<Map<String, Object>> readFile(String path) {
    return POIExcelUtils.readFile(path, null, null, null, -1, -1);
  }

  public static List<Map<String, Object>> readFile(String path, Integer sheetIndex) {
    return POIExcelUtils.readFile(path, sheetIndex, null, null, -1, -1);
  }

  public static List<Map<String, Object>> readFile(String path, Integer sheetIndex,
      List<String> headers) {
    return POIExcelUtils.readFile(path, sheetIndex, null, headers, -1, -1);
  }

  public static List<Map<String, Object>> readFile(String path, Integer sheetIndex,
      List<String> headers, int rowStart) {
    return POIExcelUtils.readFile(path, sheetIndex, null, headers, rowStart, -1);
  }

  public static List<Map<String, Object>> readFile(String path, Integer sheetIndex,
      List<String> headers, int rowStart, int columnStart) {
    return POIExcelUtils.readFile(path, sheetIndex, null, headers, rowStart, columnStart);
  }

  public static List<Map<String, Object>> readFile(String path, String sheetName) {
    return POIExcelUtils.readFile(path, null, sheetName, null, -1, -1);
  }

  public static List<Map<String, Object>> readFile(String path, String sheetName,
      List<String> headers) {
    return POIExcelUtils.readFile(path, null, sheetName, headers, -1, -1);
  }

  public static List<Map<String, Object>> readFile(String path, String sheetName,
      List<String> headers, int rowStart) {
    return POIExcelUtils.readFile(path, null, sheetName, headers, rowStart, -1);
  }

  public static List<Map<String, Object>> readFile(String path, String sheetName,
      List<String> headers, int rowStart, int columnStart) {
    return POIExcelUtils.readFile(path, null, sheetName, headers, rowStart, columnStart);
  }

  public static void writeFileByClasspathTemplate(String path, String templatePath,
      Map<String, Object> data) {
    BeetlUtils.renderToFileByClasspathTemplate(path, templatePath, data);
  }

  public static void writeFileByStringTemplate(String path, String template,
      Map<String, Object> data) {
    BeetlUtils.renderToFileByStringTemplate(path, template, data);
  }

  private ExcelUtils() {
    throw new UnsupportedOperationException();
  }
}
