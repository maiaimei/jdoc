package cn.maiaimei.jdoc.excel.model;

import java.util.List;
import lombok.Data;

@Data
public class ExcelReadProperties {

  private String sheetName;
  private Integer sheetIndex;
  private Integer rowStart;
  private Integer rowEnd;
  private Short columnStart;
  private Short columnEnd;
  private List<String> headers;
}
