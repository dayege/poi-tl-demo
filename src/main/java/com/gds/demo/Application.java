package com.gds.demo;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.policy.HackLoopTableRenderPolicy;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.IntStream;
import org.apache.commons.lang3.RandomStringUtils;

/**
 * @Author gedingshan
 * @Date 2024/3/10 21:38
 */
public class Application {

  public static void main(String[] args) throws IOException {
    Map<String, Object> map = new HashMap<>();
    List<Map<String, Object>> maps = new ArrayList<>();
    IntStream.rangeClosed(1,5).forEach(e->{
      Map<String,Object> lineData=new HashMap<>();
      lineData.put("no",e);
      lineData.put("text1", RandomStringUtils.randomAlphabetic(5));
      lineData.put("text2", RandomStringUtils.randomAlphabetic(5));
      maps.add(lineData);
    });
    map.put("table1", maps);
    HackLoopTableRenderPolicy policy = new HackLoopTableRenderPolicy();
    Configure config = Configure.builder()
        .bind("table1", policy)
        .build();
    XWPFTemplate template = XWPFTemplate.compile(Objects.requireNonNull(
        Application.class.getClassLoader().getResourceAsStream("demo.weintdocx")), config);
    template.render(map,new FileOutputStream("E:\\template\\data\\demo.docx"));
  }

}
