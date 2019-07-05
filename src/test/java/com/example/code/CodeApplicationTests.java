package com.example.code;

import cn.hutool.json.JSONUtil;
import com.example.code.excel.ImportExcelUtil;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

@RunWith(SpringRunner.class)
@SpringBootTest
public class CodeApplicationTests {

    @Test
    public void contextLoads() throws Exception {
        String xlsx="F:/withhead.xlsx";
        InputStream  in=new FileInputStream(xlsx);
        List<Map<String, Object>> mapList = ImportExcelUtil.parseExcel(xlsx);
        System.out.println(JSONUtil.toJsonStr(mapList));
    }

}
