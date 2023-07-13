package cn.maiaimei.jdoc;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import cn.maiaimei.framework.utils.JoltUtil;
import cn.maiaimei.jdoc.excel.model.ExcelReadProperties;
import cn.maiaimei.jdoc.excel.utils.ExcelUtils;
import cn.maiaimei.jdoc.model.PurchaseOrderItem;
import com.bazaarvoice.jolt.Chainr;
import com.bazaarvoice.jolt.JsonUtils;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

public class ExcelUtilsTest {

  private static final String ROOT_PATH = "E:\\app\\tmp\\";

  @Test
  public void testReadNoHeaders() throws JsonProcessingException {
    final String pathname = ROOT_PATH.concat("20230604143626248.xlsx");
    final ExcelReadProperties properties = new ExcelReadProperties();
    properties.setSheetIndex(0);
    properties.setRowStart(1);
    final List<Map<String, Object>> maps = ExcelUtils.read(pathname, properties);
    assertTrue(maps.size() > 0);

    List<Object> spec = JsonUtils.classpathToList("/template/import-purchase-order.json");
    Chainr chainr = Chainr.fromSpec(spec);
    Object output = chainr.transform(maps);
    assertNotNull(output);

    final ObjectMapper objectMapper = new ObjectMapper();
    final List<PurchaseOrderItem> purchaseOrderItems = objectMapper.readValue(
        objectMapper.writeValueAsString(output),
        new TypeReference<List<PurchaseOrderItem>>() {
        });
    assertTrue(purchaseOrderItems.size() > 0);
    
    final List<PurchaseOrderItem> transform = JoltUtil.transform(spec, maps,
        new TypeReference<List<PurchaseOrderItem>>() {
        });
    assertTrue(transform.size() > 0);
  }

  @Test
  public void testReadWithHeaders() {
    final String pathname = ROOT_PATH.concat("20230604143626248.xlsx");
    final ExcelReadProperties properties = new ExcelReadProperties();
    properties.setSheetIndex(0);
    properties.setRowStart(1);
    properties.setHeaders(List.of("serialNum", "purchaseTime", "itemNum", "itemName",
        "unitPrice", "count",
        "amount", "supplier"));
    final List<Map<String, Object>> maps = ExcelUtils.read(pathname, properties);
    assertTrue(maps.size() > 0);
  }

}
