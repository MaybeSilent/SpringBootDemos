

package com.maybesilent.filesoperator.util;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
    private static final String REGEX = "[a-zA-Z]";
    private Workbook workbook;
    private OutputStream os;
    private String pattern = "yyyy-MM-dd HH:mm:ss";
    private Set<String> titles = new LinkedHashSet();

    public void setPattern(String pattern) {
        this.pattern = pattern;
    }

    public ExcelUtil(Workbook workboook) {
        this.workbook = workboook;
    }

    public ExcelUtil(InputStream is, String version) throws FileNotFoundException, IOException {
        if ("2003".equals(version)) {
            this.workbook = new HSSFWorkbook(is);
        } else {
            this.workbook = new XSSFWorkbook(is);
        }

    }

    public ExcelUtil(InputStream is) throws FileNotFoundException, IOException {
        this.workbook = new XSSFWorkbook(is);
    }

    public String toString() {
        return "共有 " + this.getSheetCount() + "个sheet 页！";
    }

    public String toString(int sheetIx) throws IOException {
        return "第 " + (sheetIx + 1) + "个sheet 页，名称： " + this.getSheetName(sheetIx) + "，共 " + this.getRowCount(sheetIx) + "行！";
    }

    public static boolean isExcel(String pathname) {
        if (pathname == null) {
            return false;
        } else {
            return pathname.endsWith(".xls") || pathname.endsWith(".xlsx");
        }
    }

    public List<List<String>> read() throws Exception {
        return this.read(0, 0, this.getRowCount(0) - 1);
    }

    public List<List<String>> read(int sheetIx) throws Exception {
        return this.read(sheetIx, 0, this.getRowCount(sheetIx) - 1);
    }

    public List<List<String>> read(int sheetIx, int start, int end) throws Exception {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        List<List<String>> list = new ArrayList();
        if (end > this.getRowCount(sheetIx)) {
            end = this.getRowCount(sheetIx);
        }

        int cols = sheet.getRow(0).getLastCellNum();

        for(int i = start; i <= end; ++i) {
            List<String> rowList = new ArrayList();
            Row row = sheet.getRow(i);

            for(int j = 0; j < cols; ++j) {
                if (row == null) {
                    rowList.add((String) null);
                } else {
                    rowList.add(this.getCellValueToString(row.getCell(j)));
                }
            }

            list.add(rowList);
        }

        return list;
    }

    public boolean write(List<List<String>> rowData) throws IOException {
        return this.write(0, rowData, 0);
    }

    public boolean write(List<List<String>> rowData, String sheetName, boolean isNewSheet) throws IOException {
        Sheet sheet = null;
        if (isNewSheet) {
            sheet = this.workbook.createSheet(sheetName);
        } else {
            sheet = this.workbook.createSheet();
        }

        int sheetIx = this.workbook.getSheetIndex(sheet);
        return this.write(sheetIx, rowData, 0);
    }

    public boolean write(int sheetIx, List<List<String>> rowData, boolean isAppend) throws IOException {
        if (isAppend) {
            return this.write(sheetIx, rowData, this.getRowCount(sheetIx));
        } else {
            this.clearSheet(sheetIx);
            return this.write(sheetIx, rowData, 0);
        }
    }

    public boolean write(int sheetIx, List<List<String>> rowData, int startRow) throws IOException {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        int dataSize = rowData.size();
        if (this.getRowCount(sheetIx) > 0) {
            sheet.shiftRows(startRow, this.getRowCount(sheetIx), dataSize);
        }

        for(int i = 0; i < dataSize; ++i) {
            Row row = sheet.createRow(i + startRow);

            for(int j = 0; j < ((List)rowData.get(i)).size(); ++j) {
                Cell cell = row.createCell(j);
                cell.setCellValue((String)((List)rowData.get(i)).get(j) + "");
            }
        }

        return true;
    }

    public boolean setStyle(int sheetIx, int rowIndex, int colIndex, CellStyle style) throws IOException {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        sheet.setColumnWidth(colIndex, 4000);
        Cell cell = sheet.getRow(rowIndex).getCell(colIndex);
        cell.setCellStyle(style);
        return true;
    }

    public CellStyle makeStyle(int type) {
        CellStyle style = this.workbook.createCellStyle();
        DataFormat format = this.workbook.createDataFormat();
        style.setDataFormat(format.getFormat("@"));
        style.setAlignment(HorizontalAlignment.CENTER);
        Font font = this.workbook.createFont();
        if (type == 1) {
            font.setBold(true);
            font.setFontHeight((short)500);
        }

        if (type == 2) {
            font.setBold(true);
            font.setFontHeight((short)300);
        }

        style.setFont(font);
        return style;
    }

    public void region(int sheetIx, int firstRow, int lastRow, int firstCol, int lastCol) {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    public boolean isRowNull(int sheetIx, int rowIndex) throws IOException {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        return sheet.getRow(rowIndex) == null;
    }

    public boolean createRow(int sheetIx, int rownum) throws IOException {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        sheet.createRow(rownum);
        return true;
    }

    public boolean isCellNull(int sheetIx, int rowIndex, int colIndex) throws IOException {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        if (!this.isRowNull(sheetIx, rowIndex)) {
            return false;
        } else {
            Row row = sheet.getRow(rowIndex);
            return row.getCell(colIndex) == null;
        }
    }

    public boolean createCell(int sheetIx, int rowIndex, int colIndex) throws IOException {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        Row row = sheet.getRow(rowIndex);
        row.createCell(colIndex);
        return true;
    }

    public int getRowCount(int sheetIx) {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        return sheet.getPhysicalNumberOfRows() == 0 ? 0 : sheet.getLastRowNum() + 1;
    }

    public int getColumnCount(int sheetIx, int rowIndex) {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        Row row = sheet.getRow(rowIndex);
        return row == null ? -1 : row.getLastCellNum();
    }

    public boolean setValueAt(int sheetIx, int rowIndex, int colIndex, String value) throws IOException {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        sheet.getRow(rowIndex).getCell(colIndex).setCellValue(value);
        return true;
    }

    public String getValueAt(int sheetIx, int rowIndex, int colIndex) {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        return this.getCellValueToString(sheet.getRow(rowIndex).getCell(colIndex));
    }

    public boolean setRowValue(int sheetIx, List<String> rowData, int rowIndex) throws IOException {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        Row row = sheet.getRow(rowIndex);

        for(int i = 0; i < rowData.size(); ++i) {
            row.getCell(i).setCellValue((String)rowData.get(i));
        }

        return true;
    }

    public List<String> getRowValue(int sheetIx, int rowIndex) {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        Row row = sheet.getRow(rowIndex);
        List<String> list = new ArrayList();
        if (row == null) {
            list.add((String) null);
        } else {
            for(int i = 0; i < row.getLastCellNum(); ++i) {
                list.add(this.getCellValueToString(row.getCell(i)));
            }
        }

        return list;
    }

    public List<String> getColumnValue(int sheetIx, int rowIndex, int colIndex) {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        List<String> list = new ArrayList();

        for(int i = rowIndex; i < this.getRowCount(sheetIx); ++i) {
            Row row = sheet.getRow(i);
            if (row == null) {
                list.add((String) null);
            } else {
                list.add(this.getCellValueToString(sheet.getRow(i).getCell(colIndex)));
            }
        }

        return list;
    }

    public int getSheetCount() {
        return this.workbook.getNumberOfSheets();
    }

    public void createSheet() {
        this.workbook.createSheet();
    }

    public boolean setSheetName(int sheetIx, String name) throws IOException {
        this.workbook.setSheetName(sheetIx, name);
        return true;
    }

    public String getSheetName(int sheetIx) throws IOException {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        return sheet.getSheetName();
    }

    public int getSheetIndex(String name) {
        return this.workbook.getSheetIndex(name);
    }

    public boolean removeSheetAt(int sheetIx) throws IOException {
        this.workbook.removeSheetAt(sheetIx);
        return true;
    }

    public boolean removeRow(int sheetIx, int rowIndex) throws IOException {
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        sheet.shiftRows(rowIndex + 1, this.getRowCount(sheetIx), -1);
        Row row = sheet.getRow(this.getRowCount(sheetIx) - 1);
        sheet.removeRow(row);
        return true;
    }

    public void setSheetOrder(String sheetname, int sheetIx) {
        this.workbook.setSheetOrder(sheetname, sheetIx);
    }

    public boolean clearSheet(int sheetIx) throws IOException {
        String sheetname = this.getSheetName(sheetIx);
        this.removeSheetAt(sheetIx);
        this.workbook.createSheet(sheetname);
        this.setSheetOrder(sheetname, sheetIx);
        return true;
    }

    public Workbook getWorkbook() {
        return this.workbook;
    }

    public void close() throws IOException {
        if (this.os != null) {
            this.os.close();
        }

        this.workbook.close();
    }

    public <T> List<T> getBeans(Class<T> clazz) {
        int titleRowIndex = 0;
        int startRow = 1;
        int sheetIx = 0;
        return this.getBeans(this.pattern, titleRowIndex, startRow, sheetIx, clazz);
    }

    public <T> List<T> getBeans(String datePattern, int titleRowIndex, int startRow, int sheetIx, Class<T> clazz) {
        int endRow = this.getRowCount(sheetIx);
        Sheet sheet = this.workbook.getSheetAt(sheetIx);
        Row filedsRow = null;
        List<Row> rowList = new ArrayList();
        int lastRowNum = sheet.getLastRowNum();
        int rowLength = lastRowNum;
        if (endRow > 0) {
            rowLength = endRow;
        } else if (endRow < 0) {
            rowLength = lastRowNum + endRow;
        }

        filedsRow = sheet.getRow(titleRowIndex);

        for(int i = startRow; i < rowLength; ++i) {
            Row row = sheet.getRow(i);
            rowList.add(row);
        }

        return this.returnObjectList(datePattern, filedsRow, rowList, clazz);
    }

    public <T> List<T> returnObjectList(String datePattern, Row filedsRow, List<Row> rowList, Class<T> clazz) {
        ArrayList objectList = new ArrayList();

        try {
            Iterator var9 = rowList.iterator();

            while(var9.hasNext()) {
                Row row = (Row)var9.next();
                T obj = clazz.newInstance();

                for(int j = 0; j < filedsRow.getLastCellNum(); ++j) {
                    String attribute = getCellValue(filedsRow.getCell(j));
                    if (!attribute.equals("")) {
                        this.titles.add(attribute);
                        String value = getCellValue(row.getCell(j));
                        setAttrributeValue(obj, attribute, value, datePattern);
                    }
                }

                objectList.add(obj);
            }
        } catch (Exception var12) {
            var12.printStackTrace();
        }

        return objectList;
    }

    private static void setAttrributeValue(Object obj, String attribute, String value, String datePattern) {
        String method_name = convertToMethodName(attribute, obj.getClass(), true);
        Method[] methods = obj.getClass().getMethods();
        Method[] var6 = methods;
        int var7 = methods.length;

        for(int var8 = 0; var8 < var7; ++var8) {
            Method method = var6[var8];
            if (method.getName().equals(method_name)) {
                Class[] parameterC = method.getParameterTypes();

                try {
                    if (parameterC[0] != Integer.TYPE && parameterC[0] != Integer.class) {
                        if (parameterC[0] != Long.TYPE && parameterC[0] != Long.class) {
                            if (parameterC[0] != Float.TYPE && parameterC[0] != Float.class) {
                                if (parameterC[0] != Double.TYPE && parameterC[0] != Double.class) {
                                    if (parameterC[0] != Byte.TYPE && parameterC[0] != Byte.class) {
                                        if (parameterC[0] != Boolean.TYPE && parameterC[0] != Boolean.class) {
                                            if (parameterC[0] == Date.class) {
                                                if (value != null && value.length() > 0) {
                                                    SimpleDateFormat sdf = new SimpleDateFormat(datePattern);
                                                    Date date = null;

                                                    try {
                                                        date = sdf.parse(value);
                                                    } catch (Exception var14) {
                                                        var14.printStackTrace();
                                                    }

                                                    method.invoke(obj, date);
                                                }
                                            } else if (parameterC[0] == BigDecimal.class) {
                                                if (value != null && value.length() > 0) {
                                                    method.invoke(obj, new BigDecimal(value));
                                                }
                                            } else if (value != null && value.length() > 0) {
                                                method.invoke(obj, parameterC[0].cast(value));
                                            }
                                            break;
                                        }

                                        if (value != null && value.length() > 0) {
                                            method.invoke(obj, Boolean.valueOf(value));
                                        }
                                        break;
                                    }

                                    if (value != null && value.length() > 0) {
                                        method.invoke(obj, Byte.valueOf(value));
                                    }
                                    break;
                                }

                                if (value != null && value.length() > 0) {
                                    method.invoke(obj, Double.valueOf(value));
                                }
                                break;
                            }

                            if (value != null && value.length() > 0) {
                                method.invoke(obj, Float.valueOf(value));
                            }
                            break;
                        }

                        if (value != null && value.length() > 0) {
                            value = value.substring(0, value.lastIndexOf("."));
                            method.invoke(obj, Long.valueOf(value));
                        }
                        break;
                    }

                    if (value != null && value.length() > 0) {
                        value = value.substring(0, value.lastIndexOf("."));
                        method.invoke(obj, Integer.valueOf(value));
                    }
                    break;
                } catch (IllegalArgumentException var15) {
                    var15.printStackTrace();
                } catch (IllegalAccessException var16) {
                    var16.printStackTrace();
                } catch (InvocationTargetException var17) {
                    var17.printStackTrace();
                } catch (SecurityException var18) {
                    var18.printStackTrace();
                }
            }
        }

    }

    private static String convertToMethodName(String attribute, Class<?> objClass, boolean isSet) {
        Pattern p = Pattern.compile("[a-zA-Z]");
        Matcher m = p.matcher(attribute);
        StringBuilder sb = new StringBuilder();
        if (isSet) {
            sb.append("set");
        } else {
            try {
                Field attributeField = objClass.getDeclaredField(attribute);
                if (attributeField.getType() != Boolean.TYPE && attributeField.getType() != Boolean.class) {
                    sb.append("get");
                } else {
                    sb.append("is");
                }
            } catch (SecurityException var7) {
                var7.printStackTrace();
            } catch (NoSuchFieldException var8) {
                var8.printStackTrace();
            }
        }

        if (attribute.charAt(0) != '_' && m.find()) {
            sb.append(m.replaceFirst(m.group().toUpperCase()));
        } else {
            sb.append(attribute);
        }

        return sb.toString();
    }

    private static String getCellValue(Cell cell) {
        Object result = "";
        if (cell != null) {
            switch(cell.getCellType()) {
                case STRING:
                    result = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    result = cell.getNumericCellValue();
                    break;
                case BOOLEAN:
                    result = cell.getBooleanCellValue();
                    break;
                case FORMULA:
                    result = cell.getCellFormula();
                    break;
                case ERROR:
                    result = cell.getErrorCellValue();
                case BLANK:
            }
        }

        return result.toString();
    }

    private String getCellValueToString(Cell cell) {
        String strCell = "";
        if (cell == null) {
            return null;
        } else {
            switch(cell.getCellType()) {
                case STRING:
                    strCell = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        if (this.pattern != null) {
                            SimpleDateFormat sdf = new SimpleDateFormat(this.pattern);
                            strCell = sdf.format(date);
                        } else {
                            strCell = date.toString();
                        }
                    } else {
                        cell.setCellType(CellType.STRING);
                        strCell = cell.toString();
                    }
                    break;
                case BOOLEAN:
                    strCell = String.valueOf(cell.getBooleanCellValue());
            }

            return strCell;
        }
    }

    public Set<String> getTitles() {
        return this.titles;
    }

    public void setTitles(Set<String> titles) {
        this.titles = titles;
    }
}
