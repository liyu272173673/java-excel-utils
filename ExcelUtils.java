package com.demo.utils.excel;

import com.alibaba.fastjson.JSONObject;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class ExcelUtils {
    private static Logger log = LoggerFactory.getLogger(ExcelUtils.class);

    private static <T> void transferToArrAndSetValue(HSSFWorkbook hssf, List<T> list, Class<T> clzz) throws IllegalArgumentException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        ExportExcelPar classNotpar = clzz.getAnnotation(ExportExcelPar.class);
        if (classNotpar == null) {
            log.error("����excel��entity����Ҫ���ע��ExportExcelPar");
            return;
        }
        boolean classIfExport = classNotpar.ifExport();

        if (list == null || list.size() == 0) {
            return;
        }

        HSSFSheet sheet = hssf.getSheetAt(0);
        Field[] fields = clzz.getDeclaredFields();

        for (Field field : fields) {
            if (!"serialVersionUID".equalsIgnoreCase(field.getName())) {
                // �ж��Ƿ� ����Ҫ����
                ExportExcelPar notpar = field.getAnnotation(ExportExcelPar.class);
                if (notpar == null || (notpar != null && notpar.ifExport())) {
                    // ����
                    field.setAccessible(true);
                }
            }
        }

        for (int i = 0; i < list.size(); i++) {
            List<String> strings = new ArrayList<>();
            for (Field field : fields) {
                String fName = field.getName();
                if (!"serialVersionUID".equalsIgnoreCase(fName)) {
                    // �ж��Ƿ� ����Ҫ����

                    boolean filedIfExport = classIfExport;//���classIfExportΪtrue����Ĭ������field���ɵ�������Ϊfalse��Ĭ������field�����ܵ�����
                    ExportExcelPar notpar = field.getAnnotation(ExportExcelPar.class);
                    if (notpar != null) {
                        filedIfExport = notpar.ifExport();//���fieldע�Ͳ�Ϊ�գ��򸲸�classע�͵�ֵ
                    }
                    if (filedIfExport) {
                        // ����
                        String str;
                        //ƴ��get����
                        String methodName = convertMethod("get", fName);
                        //����get����
                        Method get = clzz.getDeclaredMethod(methodName);
                        str = String.valueOf(get.invoke(list.get(i)));
                        if (notpar != null) {
//							str = (str == null || str.equalsIgnoreCase("null") ? "" : str);
//						}else{
                            if (StringUtils.isNotBlank(str)) {
                                //ǰ׺/ ��׺
                                if (!"".equals(notpar.prefix())) {
                                    str = notpar.prefix() + str;
                                }
                                if (!"".equals(notpar.postfix())) {
                                    str = str + notpar.postfix();
                                }
                            } else {
                                str += "";
                            }
                        }
                        str = (str == null || str.equalsIgnoreCase("null") ? "" : str);
                        strings.add(str);
                    }
                }
            }
            for (int j = 0; j < strings.size(); j++) {
                setValue(sheet, i, j, strings);
            }
        }

    }

    private static <T> void transferToArrAndSetValue(HSSFWorkbook hssf, List<T> list, Class<T> clzz, List<String> selectColumns) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        if (list == null || list.size() == 0) {
            return;
        }
        HSSFSheet sheet = hssf.getSheetAt(0);
        for (int i = 0; i < list.size(); i++) {
            List<String> strings = new ArrayList<>();
            for (String filed : selectColumns) {
                // ����
                String str = null;
                //ƴ��get����
                String methodName = convertMethod("get", filed);
                //����get����
                Method get = clzz.getDeclaredMethod(methodName);
                Object filedValue = get.invoke(list.get(i));
                if (filedValue != null) {
                    switch (filed) {
                        case "country":
                            str = ContextTools.getCountry(filedValue.toString());
                            break;
                        case "occupation":
                            str = ContextTools.getJob(filedValue.toString());
                            break;
                        case "gender":
                            str = "1".equals(filedValue.toString()) ? "��" : "Ů";
                            break;
                        case "bankId":
                            str = ContextTools.getBank(filedValue.toString());
                            break;
                        default:
                            str = String.valueOf(filedValue);
                            break;
                    }
                }
                str = (str == null || str.equalsIgnoreCase("null") ? "" : str);
                strings.add(str);
            }
            for (int j = 0; j < strings.size(); j++) {
                setValue(sheet, i, j, strings);
            }
        }
    }

    private static String convertMethod(String prefix, String name) {
        return prefix + Character.toUpperCase(name.charAt(0)) + name.substring(1);
    }

    private static void setValue(HSSFSheet sheet, int i, int j, List<String> strings) {
        HSSFCell cellTemp = null;
        if (j == 0) {
            cellTemp = sheet.createRow(i + 1).createCell(j);
            cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
            cellTemp.setCellValue(strings.get(j));
        } else {
            cellTemp = sheet.getRow(i + 1).createCell(j);
            cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
            cellTemp.setCellValue(strings.get(j));
        }
    }

    /**
     * ����excel
     *
     * @param response
     * @param arr      ��ͷ�б�
     * @param list     ��������
     * @param clzz     ���ݶ�������
     * @throws IllegalArgumentException
     * @throws IllegalAccessException
     * @author howard
     */
    public static <T> void export(HttpServletResponse response, String[] arr, List<T> list, Class<T> clzz) throws IllegalArgumentException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        //String arr[] = {"����","1��","2-10��","11-50��","51-100��","101��"};
        HSSFWorkbook hssf = writeExcelFirst(arr, true);
        transferToArrAndSetValue(hssf, list, clzz);
        responseXLS("" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()), response, hssf);
    }

    /**
     * ����excel
     *
     * @param response
     * @param xlsName  xcel����ǰ׺
     * @param arr      ��ͷ�б�
     * @param list     ��������
     * @param clzz     ���ݶ�������
     * @throws IllegalArgumentException
     * @throws IllegalAccessException
     * @author howard
     */
    public static <T> void export(HttpServletResponse response, String xlsName, String[] arr, List<T> list, Class<T> clzz) throws IllegalArgumentException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        //String arr[] = {"����","1��","2-10��","11-50��","51-100��","101��"};
        HSSFWorkbook hssf = writeExcelFirst(arr, true);
        transferToArrAndSetValue(hssf, list, clzz);
        if (StringUtils.isBlank(xlsName)) {
            xlsName = "";
        }
        responseXLS(xlsName + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()), response, hssf);
    }

    public static <T> void export(HttpServletResponse response, String xlsName, String[] arr, List<T> list, Class<T> clzz, List<String> selectColumns) throws IllegalArgumentException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        HSSFWorkbook hssf = writeExcelFirst(arr, true);
        transferToArrAndSetValue(hssf, list, clzz, selectColumns);
        if (StringUtils.isBlank(xlsName)) {
            xlsName = "";
        }
        responseXLS(xlsName + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()), response, hssf);
    }


    public static void exportProfit(HttpServletResponse response, String xlsName, List<String> titleList, List<String> keyList, List<Map<String, String>> dataList) throws IllegalArgumentException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        if (dataList == null || dataList.size() == 0) {
            return;
        }

        HSSFWorkbook hssf = new HSSFWorkbook(); // ��������������
        HSSFSheet sheet = hssf.createSheet(); // �������������
        hssf.setSheetName(0, "default");
        sheet.setColumnWidth(0, 20 * 256);
        int size = titleList.size();

        HSSFRow hssfRow = sheet.createRow(0);
        HSSFCell cellTemp = hssfRow.createCell(0);
        cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
        cellTemp.setCellValue("С��");

        cellTemp = hssfRow.createCell(1);
        cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
        cellTemp.setCellValue("��ַ");

        for (int i = 0; i < size; i++) {
            cellTemp = hssfRow.createCell(i + 2);
            cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
            cellTemp.setCellValue(titleList.get(i));
        }

        Map<String, String> map = null;
        String key = "";
        JSONObject jsonObject = null;
        String estateName = "";
        for (int i = 0; i < dataList.size(); i++) {
            map = dataList.get(i);
            jsonObject = JSONObject.parseObject(YuxiaorUtils.objectToJson(map));

            hssfRow = sheet.createRow(i + 1);

            cellTemp = hssfRow.createCell(0);
            cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
            cellTemp.setCellValue(map.get("estateName"));


            estateName = map.get("estateName");
            if (jsonObject.get("buildingId") == null || "0".equals(jsonObject.getString("buildingId"))) {
                if (StringUtils.isNotBlank(jsonObject.getString("building"))) {
                    estateName = estateName + jsonObject.getString("building") + "��";
                }

                if (StringUtils.isNotBlank(jsonObject.getString("cell"))) {
                    estateName = estateName + jsonObject.getString("cell") + "��Ԫ";
                }
            }

            if (StringUtils.isNotBlank(jsonObject.getString("room"))) {
                estateName = estateName + jsonObject.getString("room") + "��";
            }

            cellTemp = hssfRow.createCell(1);
            cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
            cellTemp.setCellValue(estateName);

            for (int j = 0; j < keyList.size(); j++) {
                key = keyList.get(j);
                cellTemp = hssfRow.createCell(j + 2);
                cellTemp.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                cellTemp.setCellValue(jsonObject.getDoubleValue(key));
            }
        }

        if (StringUtils.isBlank(xlsName)) {
            xlsName = "";
        }
        responseXLS(xlsName + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()), response, hssf);
    }

    public static void exportDevice(HttpServletResponse response, String xlsName, List<String> titleList, List<String> keyList, List<Map<String, String>> dataList) throws IllegalArgumentException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        if (dataList == null || dataList.size() == 0) {
            return;
        }

        HSSFWorkbook hssf = new HSSFWorkbook(); // ��������������
        HSSFSheet sheet = hssf.createSheet(); // �������������
        hssf.setSheetName(0, "default");
        sheet.setColumnWidth(0, 20 * 256);
        int size = titleList.size();

        HSSFRow hssfRow = sheet.createRow(0);
        HSSFCell cellTemp = null;
        for (int i = 0; i < size; i++) {
            cellTemp = hssfRow.createCell(i);
            cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
            cellTemp.setCellValue(titleList.get(i));
        }

        Map<String, String> map = null;
        String key = "";
        for (int i = 0; i < dataList.size(); i++) {
            map = dataList.get(i);
            hssfRow = sheet.createRow(i + 1);
            for (int j = 0; j < titleList.size(); j++) {
                key = keyList.get(j);
                cellTemp = hssfRow.createCell(j);
                cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
                cellTemp.setCellValue(map.get(key));
            }
        }

        if (StringUtils.isBlank(xlsName)) {
            xlsName = "";
        }
        responseXLS(xlsName + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()), response, hssf);
    }

    /**
     * trueΪдһ�б���
     * falseΪдһ�б���
     *
     * @param arr
     * @param ifRow
     * @return
     */
    private static HSSFWorkbook writeExcelFirst(String[] arr, boolean ifRow) {
        HSSFWorkbook workbook = new HSSFWorkbook(); // ��������������
        HSSFSheet sheet = workbook.createSheet(); // �������������
        workbook.setSheetName(0, "default");
        sheet.setColumnWidth(0, 20 * 256);
        int size = arr.length;
        for (int i = 0; i < size; i++) {
            if (ifRow) {
                HSSFCell cellTemp = null;
                if (i == 0) {
                    cellTemp = sheet.createRow(0).createCell(i);
                } else {
                    cellTemp = sheet.getRow(0).createCell(i);
                }
                cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
                cellTemp.setCellValue(arr[i]);
            } else {
                HSSFCell cellTemp = sheet.createRow(i).createCell(0);
                cellTemp.setCellType(HSSFCell.CELL_TYPE_STRING);
                cellTemp.setCellValue(arr[i]);
            }

        }
        return workbook;
    }

    public static void responseXLS(String xlsName, HttpServletResponse response, HSSFWorkbook workbook) {
        OutputStream os = null;
        try {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("content-disposition", "attachment;filename=" + xlsName + ".xls");
            // д�뵽 �ͻ���response
            os = response.getOutputStream();
            workbook.write(os);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (os != null) {
                    os.flush();
                    os.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
    }

}
