package cn.maiaimei.jdocs.core;

import com.bazaarvoice.jolt.Chainr;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import java.util.List;

public class JsonUtils {

  private static final ObjectMapper OBJECT_MAPPER = new ObjectMapper();

  static {
    // ignore unknown properties
    OBJECT_MAPPER.configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false);

    // java.lang.IllegalArgumentException: Java 8 date/time type `java.time.LocalDateTime` not supported by default
    OBJECT_MAPPER.disable(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS);
    OBJECT_MAPPER.registerModule(new JavaTimeModule());
  }

  public static String stringify(Object value) throws JsonProcessingException {
    return OBJECT_MAPPER.writeValueAsString(value);
  }

  public static <T> T parse(String value, Class<T> clazz) throws JsonProcessingException {
    return OBJECT_MAPPER.readValue(value, clazz);
  }

  public static <T> T parse(String value, TypeReference<T> valueTypeRef)
      throws JsonProcessingException {
    return OBJECT_MAPPER.readValue(value, valueTypeRef);
  }

  public static <T> T transform(String specClassPath, Object input, Class<T> clazz)
      throws JsonProcessingException {
    List<Object> spec = com.bazaarvoice.jolt.JsonUtils.classpathToList(specClassPath);
    Chainr chainr = Chainr.fromSpec(spec);
    Object output = chainr.transform(input);
    return parse(stringify(output), clazz);
  }

  public static <T> T transform(String specClassPath, Object input, TypeReference<T> valueTypeRef)
      throws JsonProcessingException {
    List<Object> spec = com.bazaarvoice.jolt.JsonUtils.classpathToList(specClassPath);
    Chainr chainr = Chainr.fromSpec(spec);
    Object output = chainr.transform(input);
    return parse(stringify(output), valueTypeRef);
  }

  private JsonUtils() {
    throw new UnsupportedOperationException();
  }
}
