package com.idea;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: PostJsonAnalyzer
 * @Description: Postman JSON数据的解析工具类
 *  功能：
 *      解析Postman导出的JSON集合文件，生成Excel报表
 *      每个JSON文件对应一个Sheet，提取特定请求的URL和Form-data
 * @Author: yuanbao
 * @Date: 2025/3/5
 **/
public class PostJsonAnalyzer
{
    // JSON解析工具
    private static final ObjectMapper mapper = new ObjectMapper();

    // URL匹配模式，仅处理包含!select*.m的请求
    private static final String URL_PATTERN = "!select.*\\.m$";

    // 静态初始化块
    static
    {
        try
        {
            System.out.println("Initializing PostJsonAnalyzer...");
            // 其他静态初始化代码
        } catch (Exception e)
        {
            e.printStackTrace();
            throw new RuntimeException("Failed to initialize PostJsonAnalyzer", e);
        }
    }

    /**
     * 主方法，程序入口
     *
     * @param args
     *   命令行参数：
     *         1. 输入目录（可选，默认为jar所在目录，可传参 "input/"）
     *         2. 输出文件（可选，默认为 "output.xlsx"）
     *         3. URL匹配模式（可选，默认为 "!select.*\\.m$"）
     * @throws Exception
     *         异常处理
     */
    public static void main(String[] args) throws Exception
    {

        // 获取当前JAR包所在目录
        String jarDir = getJarDirectory();
        if (jarDir == null)
        {
            System.err.println("无法获取JAR包所在目录，程序退出。");
            throw new RuntimeException("无法获取JAR包所在目录，程序退出。");
        }
        System.out.println("当前JAR包所在目录：" + jarDir);

        // 参数初始化
        String inputDirStr = jarDir; // 默认输入目录为jar所在目录

        // 解析命令行参数
        if (args.length >= 1)
        {
            inputDirStr = args[0];
        }
        
        // 校验输入目录是否存在
        Path inputDir = Paths.get(inputDirStr);
        if (!Files.exists(inputDir) || !Files.isDirectory(inputDir))
        {
            throw new IllegalArgumentException("输入目录不存在或不是一个有效的目录: " + inputDirStr);
        }

        analysisAndSolve(inputDir);
    }

    /**
     * @MethodName: analysisAndSolve
     * @Description: 获取文件进行分析，之后生成excel
     * @param inputDir
     * @Return void
     * @Author: yuanbao
     * @Date: 2025/3/5
     **/
    private static void analysisAndSolve(Path inputDir) throws IOException
    {
        // 创建输出文件
        String outputFileStr = inputDir.toString() + "/output.xlsx";
        Path outputFile = Paths.get(outputFileStr);
        try
        {
            Files.deleteIfExists(outputFile); // 如果文件已存在，先删除
        } catch (IOException e)
        {
            throw new RuntimeException("删除输出文件失败: " + outputFileStr, e);
        }


        // 创建Excel工作簿
        try (Workbook workbook = new XSSFWorkbook())
        {
            // 遍历输入目录下的所有JSON文件
            Files.list(inputDir).filter(path -> path.toString().endsWith(".json")) // 仅处理JSON文件
                    .forEach(path -> processFile(path, workbook)); // 处理每个文件

            // 将工作簿写入输出文件
            try (OutputStream os = Files.newOutputStream(outputFile))
            {
                workbook.write(os);
            } catch (Exception e)
            {
                throw new RuntimeException("写入Excel文件失败: " + outputFileStr, e);
            }
        } catch (Exception e)
        {
            throw new RuntimeException("创建Excel工作簿失败", e);
        }

    }

    /**
     * 获取当前JAR包所在目录
     *
     * @return JAR包所在目录的路径，如果无法获取则返回null
     */
    private static String getJarDirectory() {
        try {
            // 获取JAR包的路径
            String jarPath = PostJsonAnalyzer.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath();
            File jarFile = new File(jarPath);
            return jarFile.getParent(); // 返回JAR包所在目录
        } catch (URISyntaxException e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 处理单个JSON文件
     *
     * @param jsonFile
     *         JSON文件路径
     * @param workbook
     *         Excel工作簿
     */
    private static void processFile(Path jsonFile, Workbook workbook)
    {
        try
        {
            // 解析JSON文件
            JsonNode root = mapper.readTree(jsonFile.toFile());
            // 提取Sheet名称（从文件名中提取）
            String sheetName = extractSheetName(jsonFile.getFileName().toString());

            // 创建Sheet
            Sheet sheet = workbook.createSheet(sheetName);
            // 创建表头行
            createHeaderRow(sheet);

            // 存储解析后的数据
            List<Map<String, Object>> dataList = new ArrayList<>();
            // 递归处理JSON中的item节点
            processItems(root.path("item"), dataList);

            // 将数据写入Sheet
            int rowNum = 1;
            for (Map<String, Object> data : dataList)
            {
                createDataRow(sheet, rowNum++, data);
            }
        } catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    /**
     * 从文件名中提取Sheet名称
     *
     * @param fileName
     *         文件名（如"V5.2-生产执行.postman_collection.json"）
     * @return Sheet名称（如"生产执行"）
     */
    private static String extractSheetName(String fileName)
    {
        return fileName.split("-")[1].split("\\.")[0];
    }

    /**
     * 创建表头行，并设置自动调整列宽
     *
     * @param sheet
     *         Excel Sheet
     */
    private static void createHeaderRow(Sheet sheet)
    {
        String[] headers = { "序号", "功能名", "请求URL", "请求Form-Data" };
        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < headers.length; i++)
        {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);

            // 设置列宽
            sheet.setColumnWidth(i, 256 * 30); // 每列宽度为30个字符

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
     * 递归处理JSON中的item节点
     *
     * @param items
     *         JSON中的item节点
     * @param dataList
     *         存储解析后的数据
     */
    private static void processItems(JsonNode items, List<Map<String, Object>> dataList)
    {
        items.forEach(item -> {
            if (item.has("item"))
            {
                // 如果item包含子item，递归处理
                processItems(item.path("item"), dataList);
            } else if (isValidRequest(item, URL_PATTERN))
            {
                // 如果是有效的请求，提取数据
                Map<String, Object> data = new LinkedHashMap<>();
                data.put("name", item.path("name").asText()); // 功能名
                data.put("url", item.path("request").path("url").asText()); // 请求URL
                data.put("formdata", processFormData(item)); // 请求Form-data
                dataList.add(data);
            }
        });
    }

    /**
     * 判断请求是否有效（URL匹配指定模式）
     *
     * @param item
     *         JSON节点
     * @param urlPattern
     *         URL匹配模式
     * @return 是否有效
     */
    private static boolean isValidRequest(JsonNode item, String urlPattern)
    {
        String url = item.path("request").path("url").asText("");
        return url.matches(".*" + urlPattern);
    }

    /**
     * 处理Form-data，转换为键值对
     *
     * @param item
     *         JSON节点
     * @return Form-data键值对
     */
    private static Map<String, String> processFormData(JsonNode item)
    {
        Map<String, String> formDataMap = new LinkedHashMap<>();
        JsonNode formdata = item.path("request").path("body").path("formdata");
        formdata.forEach(field -> {
            String key = field.path("key").asText();
            String value = field.path("value").asText();
            if (!value.isEmpty())
            {
                formDataMap.put(key, value);
            }
        });
        return formDataMap;
    }

    /**
     * 创建数据行
     *
     * @param sheet
     *         Excel Sheet
     * @param rowNum
     *         行号
     * @param data
     *         数据
     */
    private static void createDataRow(Sheet sheet, int rowNum, Map<String, Object> data)
    {
        Row row = sheet.createRow(rowNum);
        // 序号列
        row.createCell(0).setCellValue(rowNum);
        // 功能名列
        row.createCell(1).setCellValue(data.get("name").toString());
        // 请求URL列
        row.createCell(2).setCellValue(data.get("url").toString().replace("{{url}}/", "")); // 将url中的{{url}}/剔除掉

        // 设置Form-data列的样式（自动换行）
        CellStyle wrapStyle = sheet.getWorkbook().createCellStyle();
        wrapStyle.setWrapText(true);
        wrapStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 设置垂直居中


        // Form-data列
        Cell formdataCell = row.createCell(3);
        formdataCell.setCellStyle(wrapStyle);
        formdataCell.setCellValue(formatFormData((Map<String, String>) data.get("formdata")));
    }

    /**
     * 格式化Form-data为JSON字符串
     *
     * @param formdata
     *         Form-data键值对
     * @return 格式化后的JSON字符串
     */
    public static String formatFormData(Map<String, String> formdata)
    {
        try
        {
            return mapper.writerWithDefaultPrettyPrinter().writeValueAsString(formdata).replace("\\\"", "\""); // 去除转义字符
        } catch (Exception e)
        {
            return "{}"; // 异常时返回空JSON
        }
    }
}