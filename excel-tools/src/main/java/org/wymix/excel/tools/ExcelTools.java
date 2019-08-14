package org.wymix.excel.tools;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.serializer.SerializerFeature;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.joda.time.format.DateTimeFormat;
import org.json.XML;
import org.wymix.excel.tools.exception.TemplateNotMatchException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.IntStream;

public final class ExcelTools {

    public static String import2Json(File xml) throws IOException {

        String xmlString = FileUtils.readFileToString(xml, Charset.forName("UTF-8"));
        String s = XML.toJSONObject(xmlString).getJSONObject("root").get("file").toString();
        ExcelInfo excelInfo = JSON.parseObject(s, ExcelInfo.class);
        return importExcel(excelInfo);
    }

    public static String importExcel(ExcelInfo excelInfo) {
        try (Workbook workbook = generateWorkbook(new File(excelInfo.getPath()))) {
            Sheet sheet = workbook.getSheetAt(excelInfo.getSheet().getIndex());
            Row titleRow = sheet.getRow(excelInfo.getSheet().getTitle().getIndex());

            AtomicBoolean flag = new AtomicBoolean(false);
            IntStream.range(0, titleRow.getPhysicalNumberOfCells()).boxed().forEach(i -> {
                String stringCellValue = titleRow.getCell(i).getStringCellValue();
                excelInfo.getSheet().getTitle().getColumn().forEach(column -> {
                    if (stringCellValue.equals(column.getValue())) {
                        column.setIndex(i);
                        flag.set(true);
                    }
                });
            });
            if (!flag.get()) {
                throw new TemplateNotMatchException();
            }

            List<Map<String, Object>> results = new LinkedList<>();
            for (int i = excelInfo.getSheet().getTitle().getIndex() + 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                final Row row = sheet.getRow(i);
                Map<String, Object> result = new HashMap<>();
                int finalI = i;
                excelInfo.getSheet().getTitle().getColumn().forEach(column -> {
                    Cell cell = row.getCell(column.getIndex());
                    Object value = getValue(cell, column);
                    column.validate(finalI, value);
                    result.put(column.getProperty(), value);
                });
                results.add(result);
            }
            AtomicBoolean excptionFlag = new AtomicBoolean(false);
            excelInfo.getSheet().getTitle().getColumn().forEach(column -> {
                if (!column.getErrorMsg().isEmpty()) {
                    excptionFlag.set(true);
                    column.getErrorMsg().forEach((k, v) -> {
                        Cell cell = sheet.getRow(k).getCell(column.getIndex());
                        if(cell==null){
                            cell = sheet.getRow(k).createCell(column.getIndex());
                        }
                        Comment comment = generateComment(excelInfo, sheet);
                        comment.setString(richString(excelInfo, v));
                        cell.setCellComment(comment);
                    });
                }
            });
            if (excptionFlag.get()) {
                String error_path = excelInfo.getPath().substring(0, excelInfo.getPath().lastIndexOf('.')) + "_error." + excelInfo.getPath().substring(excelInfo.getPath().lastIndexOf('.'));
                try (FileOutputStream fileOut = new FileOutputStream(error_path)) {
                    workbook.write(fileOut);
                }
                throw new RuntimeException("导入异常");
            }
            return JSON.toJSONString(results, SerializerFeature.DisableCircularReferenceDetect);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private static RichTextString richString(ExcelInfo excelInfo, String v) {
        return excelInfo.getPath().endsWith(".xls") ? new HSSFRichTextString(v) : new XSSFRichTextString(v);
    }

    public static String import2Json(String xmlPath) throws IOException {
        return import2Json(new File(xmlPath));
    }

    public static <T> List<T> convert2Object(String json, Class<T> clazz) throws IOException {
        return JSONObject.parseArray(json, clazz);
    }

    private ExcelTools() {
    }

    private static Object getValue(Cell cell, ExcelInfo.Column column) {
        if (cell == null) {
            return null;
        }
        Object cellValue;
        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    cellValue = new DateTime(cell.getDateCellValue()).toString(DateTimeFormat.forPattern((null != column.getDataFormat() && column.getDataFormat().length() > 0) ? column.getDataFormat() : "yyyy-MM-dd HH:mm:ss"));
                } else {
                    cellValue = cell.getNumericCellValue();
                }
                break;
            case FORMULA:
                try {
                    cellValue = cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    cellValue = cell.getNumericCellValue();
                }
                break;
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case BLANK:
                cellValue = null;
                break;
            case ERROR:
                cellValue = cell.getErrorCellValue();
                break;
            case STRING:
            default:
                cellValue = cell.getStringCellValue();
                break;
        }
        return cellValue;
    }

    private static Comment generateComment(ExcelInfo excelInfo, Sheet sheet) {

        Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();
        ClientAnchor clientAnchor = excelInfo.getPath().endsWith(".xls") ? new HSSFClientAnchor() : new XSSFClientAnchor();
        Comment cellComment = drawingPatriarch.createCellComment(clientAnchor);
        return cellComment;

    }

    private static Workbook generateWorkbook(File excel) throws IOException {
        FileInputStream s = new FileInputStream(excel);
        return (excel.getName().endsWith(".xls") ? new HSSFWorkbook(s) : new XSSFWorkbook(s));
    }

    public static void main(String[] args) throws IOException {
        String json = import2Json("/Users/WangSir/Project/workSpace_java/my_project/tools/excel-tools/src/main/resources/test.xml");
    }

}
