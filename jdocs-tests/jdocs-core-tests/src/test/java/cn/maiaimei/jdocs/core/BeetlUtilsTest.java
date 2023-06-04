package cn.maiaimei.jdocs.core;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

public class BeetlUtilsTest {

  @Test
  void testRender() {
    Map<String, Object> data = new HashMap<>();
    data.put("to", "World");
    final String result = BeetlUtils.renderByStringTemplate("Hello, ${to}", data);
    assertEquals("Hello, World", result);
  }

  @Test
  void testRenderToTxt() {
    Map<String, Object> data = new HashMap<>();
    data.put("subject", "Beetl入门");
    data.put("age", 17);
    BeetlUtils.renderToFileByClasspathTemplate("E:\\app\\tmp\\t1.txt", "template/t1.btl", data);
    assertTrue(Boolean.TRUE);
  }

  @Test
  void testRenderToXls() {
    Map<String, Object> p1 = new HashMap<>();
    p1.put("name", "A产品");
    p1.put("unitPrice", 8);
    p1.put("count", 20);
    Map<String, Object> p2 = new HashMap<>();
    p2.put("name", "B产品");
    p2.put("unitPrice", 11);
    p2.put("count", 5);
    List<Map<String, Object>> poList = new ArrayList<>();
    poList.add(p1);
    poList.add(p2);
    Map<String, Object> data = new HashMap<>();
    data.put("poList", poList);
    BeetlUtils.renderToFileByClasspathTemplate("E:\\app\\tmp\\t2.xls", "template/t2.btl", data);
    assertTrue(Boolean.TRUE);
  }

  @Test
  void testRenderToXlsx() {
    Map<String, Object> p1 = new HashMap<>();
    p1.put("name", "A产品");
    p1.put("unitPrice", 8);
    p1.put("count", 20);
    Map<String, Object> p2 = new HashMap<>();
    p2.put("name", "B产品");
    p2.put("unitPrice", 11);
    p2.put("count", 5);
    List<Map<String, Object>> poList = new ArrayList<>();
    poList.add(p1);
    poList.add(p2);
    Map<String, Object> data = new HashMap<>();
    data.put("poList", poList);
    BeetlUtils.renderToFileByClasspathTemplate("E:\\app\\tmp\\t3.xlsx", "template/t3.btl", data);
    assertTrue(Boolean.TRUE);
  }

  @Test
  void testRenderToDoc() {
    Map<String, Object> p1 = new HashMap<>();
    p1.put("name", "A产品");
    p1.put("unitPrice", 8);
    p1.put("count", 20);
    Map<String, Object> p2 = new HashMap<>();
    p2.put("name", "B产品");
    p2.put("unitPrice", 11);
    p2.put("count", 5);
    List<Map<String, Object>> poList = new ArrayList<>();
    poList.add(p1);
    poList.add(p2);
    Map<String, Object> data = new HashMap<>();
    data.put("poList", poList);
    BeetlUtils.renderToFileByClasspathTemplate("E:\\app\\tmp\\t4.doc", "template/t4.btl", data);
    assertTrue(Boolean.TRUE);
  }

}
