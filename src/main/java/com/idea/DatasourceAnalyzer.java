package com.idea;

import com.alibaba.fastjson.JSONObject;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;
import org.springframework.util.ObjectUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: DatasourceAnalyzer
 * @Description: TODO <br>
 * @Author: yuanbao
 * @Date: 2025/3/10
 **/
@Component
public class DatasourceAnalyzer
{
    private static final Logger log = LoggerFactory.getLogger(DatasourceAnalyzer.class);

    @Autowired
    private JdbcTemplate jdbcTemplate;

    private static final String OUTPUT_DIR = "input/";

    static {
        log.info("--------程序启动了");
    }

    public static void main(String[] args) throws Exception
    {
        //        MysqlDataSourceCfg.getConnection();
        new DatasourceAnalyzer().getDataFromMysql();
    }

    // 获取mysql中的数据
    public void getDataFromMysql() throws IOException
    {
        // 使用springboot的方式获取数据
        String sql = "select t.table_title,t.data_source from eh_dynamic_grid_config t where t.is_delete = 0";
        List<Map<String, Object>> list = jdbcTemplate.queryForList(sql);
        log.info("查询结果数量：" + list.size());

        // 数据处理
        List<Map<String, Object>> resultList = transformData(list);

        // 生成excel文件
        generateExcel(resultList);
    }

    /**
     * @MethodName: transformData
     * @Description: 转换其中的数据，如拼接出url等
     * @param list
     * @Return List<Map<String,Object>>
     * @Author: yuanbao
     * @Date: 2025/3/10
     **/
    private List<Map<String, Object>> transformData(List<Map<String, Object>> list)
    {
//        List<Map<String, Object>> resultList = Lists.newArrayList();

        for (Map<String, Object> map : list)
        {
            String tableTitle = (String) map.get("table_title");
            String dataSource = (String) map.get("data_source");
            if (!ObjectUtils.isEmpty(dataSource))
            {
                JSONObject jsonObject = JSONObject.parseObject(dataSource);
                if (jsonObject.containsKey("bean") && jsonObject.containsKey("method"))
                {
                    String bean = jsonObject.getString("bean");
                    String method = jsonObject.getString("method");
                    String url = bean + "!" + method + ".m";
                    map.put("url", url);
                }
            }
        }
        return list;
    }

    /**
     * @MethodName: generateExcel
     * @Description: 生成excel文件
     * @param list
     * @Return void
     * @Author: yuanbao
     * @Date: 2025/3/10
     **/
    private static void generateExcel(List<Map<String, Object>> list) throws IOException
    {
        // 创建Excel工作簿和工作表
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("V55动态表格请求");

        // 创建表头行
        createHeaderRow(sheet);

        // 填充数据
        createDataRow(sheet, list);


        // 写入文件
        // 按日模式和月模式分别生成不同的文件名
        String fileName ="V55_EH_GRID_OUTPUT.xlsx";
        Path excelPath = Paths.get(OUTPUT_DIR, fileName);
        try (FileOutputStream fileOut = new FileOutputStream(excelPath.toFile()))
        {
            workbook.write(fileOut);
        } catch (IOException e)
        {
            e.printStackTrace();
            System.err.println("写入Excel文件出错：" + e.toString());
            throw new RuntimeException("写入Excel文件出错：" + e.toString());
        }

        // 关闭工作簿
        workbook.close();
        log.info("Excel文件生成成功：" + excelPath.toAbsolutePath());
    }

    /**
     * @MethodName: createHeaderRow
     * @Description: 创建表头行
     * @param sheet
     * @Return void
     */
    private static void createHeaderRow(Sheet sheet)
    {
        // 表头
        Row headerRow = sheet.createRow(0);
        String[] headers = {"序号", "功能名", "请求URL", "请求Form-Data"};
        for (int i = 0; i < headers.length; i++)
        {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);

            // 设置列宽
            sheet.setColumnWidth(i, 256 * 30); // 每列宽度为20个字符

            // 如果是最后一列，设置列宽宽一点
            if (i == headers.length - 1)
            {
                sheet.setColumnWidth(i, 256 * 120); // 最后一列宽度为120个字符
            } else if (i == 0)
            {
                sheet.setColumnWidth(i, 256 * 5); // 第1列宽度为5个字符
            } else if (i == 2)
            {
                sheet.setColumnWidth(i, 256 * 40); // 第3列宽度为40个字符
            }

            // 设置自动换行
            CellStyle wrapStyle = sheet.getWorkbook().createCellStyle();
            wrapStyle.setWrapText(true);
            wrapStyle.setVerticalAlignment(VerticalAlignment.CENTER);            // 设置垂直居中
            cell.setCellStyle(wrapStyle);
        }
    }


    /**
     * @MethodName: createDataRow
     * @Description: 填充数据
     * @param sheet
     * @param list
     * @Return void
     */
    private static void createDataRow(Sheet sheet, List<Map<String, Object>> list)
    {
        for (int i = 0; i < list.size(); i++)
        {
            Map<String, Object> map = list.get(i);
            Row row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(i + 1);
            row.createCell(1).setCellValue((String) map.get("table_title"));
            row.createCell(2).setCellValue((String) map.get("url"));
//            row.createCell(3).setCellValue((String) map.get("data_source"));

            // 设置Form-data列的样式（自动换行）
            CellStyle wrapStyle = sheet.getWorkbook().createCellStyle();
            wrapStyle.setWrapText(true);
            wrapStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 设置垂直居中


            // Form-data列
            Cell formdataCell = row.createCell(3);
            formdataCell.setCellStyle(wrapStyle);
            formdataCell.setCellValue(formatJSON((String) map.get("data_source")));
        }
    }

    /**
     * @MethodName: formatJSON
     * @Description: 格式化JSON字符串，使其自动换行与缩进
     * @param jsonStr
     * @Return String
     */
    private static String formatJSON(String jsonStr)
    {
        if (jsonStr == null || jsonStr.isEmpty())
        {
            return "";
        }
        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        JsonElement jsonElement = JsonParser.parseString(jsonStr);
        String formattedJson = gson.toJson(jsonElement);
        return formattedJson;
    }
}
