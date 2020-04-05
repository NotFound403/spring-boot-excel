package cn.felord.excel.controller;

import cn.felord.excel.config.DefaultXlsView;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.util.Assert;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author dax
 * @since 2020/4/5 14:27
 */
@Controller
@RequestMapping("/excel")
public class ExcelController {

    @Autowired
    private DefaultXlsView defaultXlsView;

    @GetMapping("/{export}")
    public ModelAndView export(@PathVariable String export) {
        Map<String, String> map = new HashMap<>();

        map.put("1", "zhangshan");
        map.put("2", "lisi");
        map.put("3", "wangwu");
        map.put("4", "zhaoliu");
        map.put("5", "liuqi");

        // 数据写入excel工作表逻辑
        defaultXlsView.dataInjector((m, w) -> {
            Object data = m.get("data");
            Assert.notNull(data, "data must not be null");
            Map<String, String> d = (Map<String, String>) data;

            Sheet userData = w.createSheet("userData");

            Row head = userData.createRow(0);

            head.createCell(0).setCellValue("Id");
            head.createCell(1).setCellValue("Name");

            AtomicInteger rowNum = new AtomicInteger(1);
            d.forEach((k, v) -> {
                Row row = userData.createRow(rowNum.getAndIncrement());

                row.createCell(0).setCellValue(k);
                row.createCell(1).setCellValue(v);
            });
        });

        // modelName 一定要对应 上面注入逻辑获取数据的key 即41 行
        return new ModelAndView(defaultXlsView, "data", map);
    }
}
