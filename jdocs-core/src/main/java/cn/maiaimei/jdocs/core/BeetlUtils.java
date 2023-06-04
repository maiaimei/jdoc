package cn.maiaimei.jdocs.core;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Map;
import org.beetl.core.Configuration;
import org.beetl.core.GroupTemplate;
import org.beetl.core.Template;
import org.beetl.core.resource.ClasspathResourceLoader;
import org.beetl.core.resource.StringTemplateResourceLoader;

/**
 * @author maiaimei
 */
public final class BeetlUtils {

  private static final GroupTemplate GROUP_TEMPLATE_FOR_CLASSPATH_RESOURCE_LOADER;
  private static final GroupTemplate GROUP_TEMPLATE_FOR_STRING_TEMPLATE_RESOURCE_LOADER;

  static {
    try {
      ClasspathResourceLoader classpathResourceLoader = new ClasspathResourceLoader();
      StringTemplateResourceLoader stringTemplateResourceLoader = new StringTemplateResourceLoader();
      Configuration configuration = Configuration.defaultConfiguration();
      GROUP_TEMPLATE_FOR_CLASSPATH_RESOURCE_LOADER = new GroupTemplate(classpathResourceLoader,
          configuration);
      GROUP_TEMPLATE_FOR_STRING_TEMPLATE_RESOURCE_LOADER = new GroupTemplate(
          stringTemplateResourceLoader,
          configuration);
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  public static String renderByStringTemplate(String template, Map<String, Object> data) {
    final Template t = GROUP_TEMPLATE_FOR_STRING_TEMPLATE_RESOURCE_LOADER.getTemplate(template);
    t.binding(data);
    return t.render();
  }

  public static void renderToFileByStringTemplate(String path, String template,
      Map<String, Object> data) {
    final Template t = GROUP_TEMPLATE_FOR_STRING_TEMPLATE_RESOURCE_LOADER.getTemplate(template);
    t.binding(data);
    try (OutputStream os = Files.newOutputStream(Paths.get(path))) {
      t.renderTo(os);
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  public static void renderToFileByClasspathTemplate(String path, String templatePath,
      Map<String, Object> data) {
    final Template t = GROUP_TEMPLATE_FOR_CLASSPATH_RESOURCE_LOADER.getTemplate(templatePath);
    t.binding(data);
    try (OutputStream os = Files.newOutputStream(Paths.get(path))) {
      t.renderTo(os);
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  private BeetlUtils() {
    throw new UnsupportedOperationException();
  }
}
