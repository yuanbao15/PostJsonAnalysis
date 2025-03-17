package com.idea;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.google.gson.*;
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
 * @Description: 提取接口数据输出Excel：连接MOM V55动态表格配置获取数据后提取接口信息，需要拼接url和关键字段信息，之后输出excel <br>
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

    /**
     * @MethodName: getDataFromMysql
     * @Description: 获取mysql中的数据，从V55动态表格配置中提取请求信息生成excel
     * @param
     * @Return void
     * @Author: yuanbao
     * @Date: 2025/3/10
     **/
    public void getDataFromMysql() throws IOException
    {
        // 使用springboot的方式获取数据
        // 动态表格配置的数据源分两部分，一部分是写接口url方式，一部分是写sql方式

        // 接口url方式
        String sql = "select t.table_title,t.data_source,t.grid from eh_dynamic_grid_config t where t.is_delete = 0" +
                " and t.table_title is not null and t.search_type='common'";
        List<Map<String, Object>> list = jdbcTemplate.queryForList(sql);
        log.info("查询接口结果数量：" + list.size());
        // sql方式
        String sql2 = "select t.table_title,t.data_source,t.grid,t.options from eh_dynamic_grid_config t where t.is_delete = 0" +
                " and t.search_type='sql'";
        List<Map<String, Object>> list2 = jdbcTemplate.queryForList(sql2);
        log.info("查询SQL结果数量：" + list2.size());

        // 数据处理
        List<Map<String, Object>> resultList = transformData(list);
        List<Map<String, Object>> resultList2 = transformData2(list2);

        // 生成excel文件
        generateExcel(resultList, resultList2);
    }

    /**
     * @MethodName: transformData
     * @Description: 转换其中的数据，如拼接出url等。动态表格配置中主要取用data_source和grid字段
     * @param list
     * @Return List<Map<String,Object>>
     * @Author: yuanbao
     * @Date: 2025/3/10
     **/
    private List<Map<String, Object>> transformData(List<Map<String, Object>> list)
    {
        // 遍历处理list
        for (Map<String, Object> map : list)
        {
            String tableTitle = (String) map.get("table_title");
            String dataSource = (String) map.get("data_source");
            String grid = (String) map.get("grid");
            // dataSource中包含请求信息，进行处理为url
            if (!ObjectUtils.isEmpty(dataSource))
            {
                JSONObject objDS = JSONObject.parseObject(dataSource);
                if (objDS.containsKey("bean") && objDS.containsKey("method"))
                {
                    String bean = objDS.getString("bean");
                    String method = objDS.getString("method");
                    String url = bean + "!" + method + ".m";

                    // 如果包含了param传参，则拼接到url后面
                    if (objDS.containsKey("params"))
                    {
                        JSONObject objParams = objDS.getJSONObject("params");
                        if (objParams != null)
                        {
                            String paramStr = "";
                            for (String key : objParams.keySet())
                            {
                                String value = objParams.getString(key);
                                paramStr += key + "=" + value + "&";
                            }
                            if (!ObjectUtils.isEmpty(paramStr) && paramStr.endsWith("&"))
                            {
                                paramStr = paramStr.substring(0, paramStr.length() - 1);
                                url += "?" + paramStr;
                            }
                        }
                    }
                    map.put("url", url);
                }

                // 处理grid信息，从中提取columns和defaultSort信息，并合并到data_source中
                JSONObject objGridNew = extractGridInfo(grid, objDS);
                map.put("grid", objGridNew.toJSONString()); // 覆盖旧grid

                // 更新data_source
                map.put("data_source", objDS.toJSONString());
            }

        }
        return list;
    }

    /***
     * @MethodName: transformData2
     * @Description: 转换其中的数据，如提取出sql。动态表格配置中主要取用options和grid字段。
     * @param list
     * @Return List<Map<String,Object>>
     * @Author: yuanbao
     * @Date: 2025/3/12
     **/
    private List<Map<String, Object>> transformData2(List<Map<String, Object>> list)
    {
        // 遍历处理list
        for (Map<String, Object> map : list)
        {
            String dataSource = (String) map.get("data_source");
            String options = (String) map.get("options");
            String grid = (String) map.get("grid");

            // dataSource中包含请求信息，进行处理为url
            if (!ObjectUtils.isEmpty(dataSource))
            {
                JSONObject objDS = JSONObject.parseObject(dataSource);
                if (objDS.containsKey("bean") && objDS.containsKey("method"))
                {
                    String bean = objDS.getString("bean");
                    String method = objDS.getString("method");
                    String url = bean + "!" + method + ".m";
                    map.put("url", url);
                }

                // options中包含sql，进行提取
                if (!ObjectUtils.isEmpty(options))
                {
                    JSONObject objOtions = JSONObject.parseObject(options);
                    if (objOtions.containsKey("extraParamFields"))
                    {
                        String extraParamFields = objOtions.getString("extraParamFields");
                        JSONObject objExtraParamFields = JSONObject.parseObject(extraParamFields);
                        if (objExtraParamFields.containsKey("queryString"))
                        {
                            String sql = objExtraParamFields.getString("queryString");

                            // 将sql语句放入到data_source里
                            objDS.put("queryString", sql);
                        }
                    }
                }

                // 处理grid信息，从中提取columns和defaultSort信息，并合并到data_source中
                JSONObject objGridNew = extractGridInfo(grid, objDS);
                map.put("grid", objGridNew.toJSONString()); // 覆盖旧grid

                // 更新data_source
                map.put("data_source", objDS.toJSONString());
            }


        }
        return list;
    }


    /**
     * @MethodName: extractGridInfo
     * @Description: 提取并处理grid中的columns和defaultSort信息，并合并到data_source中
     * @param grid 待处理的grid配置信息
     * @param objDS 待完善的data_source信息
     * @Return JSONObject
     * @Author: yuanbao
     * @Date: 2025/3/12
     **/
    private JSONObject extractGridInfo(String grid, JSONObject objDS)
    {
        JSONObject objGridNew = new JSONObject(); // 提取并处理后的column信息
        if (!ObjectUtils.isEmpty(grid))
        {
            JSONObject jsonObject = JSONObject.parseObject(grid);
            // 提取columns信息
            if (jsonObject.containsKey("columns"))
            {
                // columns为数组，遍历提取
                String columns = jsonObject.getString("columns");
                if (!ObjectUtils.isEmpty(columns))
                {
                    JSONArray jsonArray = JSONArray.parseArray(columns);
                    JSONArray jsonArrayNew = new JSONArray();
                    for (int i = 0; i < jsonArray.size(); i++)
                    {
                        JSONObject column = jsonArray.getJSONObject(i);
                        String label = column.getString("label");
                        String prop = column.getString("prop");

                        JSONObject objNew = new JSONObject();
                        // 增加判定如果字段设置为忽略了，则不添加进columns中. 注意不是isIgnore，而是pass
                        if (column.containsKey("pass") && column.getBoolean("pass"))
                        {
                            continue;
                        }

//                        objNew.put("label", label);
                        objNew.put("name", prop); // prop 改为name

                        // refEntity和refName 不一定有，有则加入
                        String refEntity = column.getString("refEntity");
                        String refName = column.getString("refName");
                        if (!ObjectUtils.isEmpty(refEntity))
                        {
                            objNew.put("refEntity", refEntity);
                        }
                        if (!ObjectUtils.isEmpty(refName))
                        {
                            objNew.put("refName", refName);
                        }
                        jsonArrayNew.add(objNew);
                    }
                    objGridNew.put("columns", jsonArrayNew);
                    // 将columns信息合并到data_source中
                    objDS.put("columns", jsonArrayNew);
                }
            }

            // 提取defaultSort信息
            if (jsonObject.containsKey("defaultSort"))
            {
                String defaultSort = jsonObject.getString("defaultSort");
                if (!ObjectUtils.isEmpty(defaultSort))
                {
                    JSONObject objDefaultSort = JSONObject.parseObject(defaultSort);
                    objGridNew.put("defaultSort", objDefaultSort);

                    // 将defaultSort信息合并到data_source中
                    if (objDefaultSort.containsKey("prop"))
                    {
                        // 排序字段
                        objDS.put("sidx", objDefaultSort.getString("prop"));
                    }
                    if (objDefaultSort.containsKey("order"))
                    {
                        // 排序方式
                        objDS.put("sord", objDefaultSort.getString("order"));
                    }
                }
            }
        }
        return objGridNew;
    }

    /**
     * @MethodName: generateExcel
     * @Description: 生成excel文件
     * @param list 接口url的结果集
     * @param list2 sql的结果集
     * @Return void
     * @Author: yuanbao
     * @Date: 2025/3/10
     **/
    private static void generateExcel(List<Map<String, Object>> list, List<Map<String, Object>> list2) throws IOException
    {
        // 创建Excel工作簿和工作表
        Workbook workbook = new XSSFWorkbook();

        // 1、处理接口url的结果集
        Sheet sheet = workbook.createSheet("V55动态表格请求");
        // 创建表头行
        createHeaderRow(sheet);
        // 填充数据
        createDataRow(sheet, list);

        // 2、处理sql的结果集
        Sheet sheet2 = workbook.createSheet("V55动态表格SQL");
        createHeaderRow(sheet2);
        createDataRow2(sheet2, list2);

        // 3、写入文件
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
        String[] headers = {"序号", "功能名", "请求URL", "请求Form-Data", "Grid配置"};
        for (int i = 0; i < headers.length; i++)
        {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);

            // 设置列宽
            sheet.setColumnWidth(i, 256 * 30); // 每列宽度为20个字符

            // 特殊列宽设置
            if (i == 0)
            {
                sheet.setColumnWidth(i, 256 * 5); // 第1列宽度为5个字符
            } else if (i == 2)
            {
                sheet.setColumnWidth(i, 256 * 40); // 第3列宽度为40个字符
            } else if (i == 3)
            {
                sheet.setColumnWidth(i, 256 * 80); // 第4列宽度为120个字符
            } else if (i == 4)
            {
                sheet.setColumnWidth(i, 256 * 60); // 最后一列宽度
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

            // 创建自动换行的样式
            CellStyle wrapStyle = sheet.getWorkbook().createCellStyle();
            wrapStyle.setWrapText(true);
            wrapStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 设置垂直居中

            // url列
            Cell urlCell = row.createCell(2);
            urlCell.setCellStyle(wrapStyle);
            urlCell.setCellValue((String) map.get("url"));

            // Form-data列
            Cell formdataCell = row.createCell(3);
            formdataCell.setCellStyle(wrapStyle);
            formdataCell.setCellValue((String) map.get("data_source")); // 不用formatJSON格式化了

            // Grid配置列
            Cell gridCell = row.createCell(4);
            gridCell.setCellStyle(wrapStyle);
            gridCell.setCellValue((String) map.get("grid"));
        }
    }

    /**
     * @MethodName: createDataRow2
     * @Description: 填充数据-针对SQL的结果集
     * @param sheet2
     * @param list2
     * @Return void
     * @Author: yuanbao
     * @Date: 2025/3/12
     **/
    private static void createDataRow2(Sheet sheet2, List<Map<String, Object>> list2)
    {
        for (int i = 0; i < list2.size(); i++)
        {
            Map<String, Object> map = list2.get(i);
            Row row = sheet2.createRow(i + 1);
            row.createCell(0).setCellValue(i + 1);
            row.createCell(1).setCellValue((String) map.get("table_title"));

            // 创建自动换行的样式
            CellStyle wrapStyle = sheet2.getWorkbook().createCellStyle();
            wrapStyle.setWrapText(true);
            wrapStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 设置垂直居中

            // url列
            Cell urlCell = row.createCell(2);
            urlCell.setCellStyle(wrapStyle);
            urlCell.setCellValue((String) map.get("url"));

            // Form-data列
            Cell formdataCell = row.createCell(3);
            formdataCell.setCellStyle(wrapStyle);
            formdataCell.setCellValue((String) map.get("data_source"));

            // Grid配置列
            Cell gridCell = row.createCell(4);
            gridCell.setCellStyle(wrapStyle);
            gridCell.setCellValue((String) map.get("grid"));
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
