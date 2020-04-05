package cn.felord.excel.config;


import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;
import org.springframework.util.Assert;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;
import java.util.function.BiConsumer;

/**
 * 2003 excel 视图
 *
 * @author dax
 * @since 2020/4/5 14:17
 */
@Component
public class DefaultXlsView extends AbstractXlsView {

    private BiConsumer<Map<String, Object>, Workbook> dataInjector;


    public void dataInjector(BiConsumer<Map<String, Object>, Workbook> dataInjector) {
        this.dataInjector = dataInjector;
    }

    @Override
    protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
        Assert.notNull(dataInjector, "dataInjector must not be null");
        dataInjector.accept(model, workbook);
    }
}
