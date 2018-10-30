package com.excel.sheet;

import com.entity.Constants;
import com.file.GetExcelValue;
import com.util.Util;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;
import java.util.List;

/**
 * 材料数据表里的数据获取
 * <p>
 * 时间 : 2018/10/6.
 */
public class ExcelCaculateParams {

    /**
     * 子结构框架梁 参数获取
     *
     * @param excel
     * @return
     */
    public static Map<String, Object> getParamsOfGirder(XSSFWorkbook excel) {
        Map<String, Object> map = new HashMap<>();
        System.out.println("================================ 获取梁的计算参数 =================================");
        XSSFSheet sheet = excel.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        //b
        XSSFCell cell = row.getCell(1);
        String value = Util.getValueFromXssfcell(cell);
        map.put(Constants.SECTION_B, value);
        //h
        cell = row.getCell(4);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.SECTION_H, value);

        //v
        row = sheet.getRow(2);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.STRESS_CONDITION_V, value);
        //M
        cell = row.getCell(4);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.STRESS_CONDITION_M, value);

        //混泥土等级
        row = sheet.getRow(4);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.CONCRETE_GRADE, value);

        //钢筋等级
        row = sheet.getRow(6);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.STEEL_GRADE, value);

        //箍筋
        row = sheet.getRow(8);
        cell = row.getCell(2);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.HOOP_D, value);

        cell = row.getCell(5);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.HOOP_L, value);

        //阿尔法
        row = sheet.getRow(10);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.PARAM_af1, value);

        cell = row.getCell(5);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.PARAM_afS, value);

        //获取材料属性
        sheet = excel.getSheetAt(1);
        getMaterialParams(map.get(Constants.CONCRETE_GRADE).toString(), map.get(Constants.STEEL_GRADE).toString(), sheet, map);
        return map;
    }

    /**
     * 子结构框架柱 参数获取
     *
     * @param excel
     * @return
     */
    public static Map<String, Object> getParamsOfPillar(XSSFWorkbook excel) {
        Map<String, Object> map = new HashMap<>();
        System.out.println("================================ 获取柱的计算参数 =================================");
        XSSFSheet sheet = excel.getSheetAt(2);
        XSSFRow row = sheet.getRow(0);
        //b
        XSSFCell cell = row.getCell(1);
        String value = Util.getValueFromXssfcell(cell);
        map.put(Constants.SECTION_B, value);
        //h
        cell = row.getCell(4);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.SECTION_H, value);

        //v
        row = sheet.getRow(2);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.STRESS_CONDITION_V, value);
        //M
        cell = row.getCell(4);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.STRESS_CONDITION_M, value);
        //P
        cell = row.getCell(7);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.STRESS_CONDITION_P, value);

        //楼层高度
        row = sheet.getRow(4);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.FLOOR_H, value);

        //混泥土等级
        row = sheet.getRow(6);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.CONCRETE_GRADE, value);

        //钢筋等级
        row = sheet.getRow(8);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.STEEL_GRADE, value);

        //箍筋
        row = sheet.getRow(10);
        cell = row.getCell(2);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.HOOP_D, value);

        cell = row.getCell(5);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.HOOP_L, value);

        row = sheet.getRow(10);
        cell = row.getCell(5);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.PARAM_afS, value);

        //获取材料属性
        sheet = excel.getSheetAt(1);
        getMaterialParams(map.get(Constants.CONCRETE_GRADE).toString(), map.get(Constants.STEEL_GRADE).toString(), sheet, map);
        return map;
    }

    /**
     * 悬臂墙配筋 参数获取
     *
     * @param excel
     * @return
     */
    public static Map<String, Object> getParamsOfCantilever(XSSFWorkbook excel) {
        Map<String, Object> map = new HashMap<>();
        System.out.println("================================ 获取悬臂的计算参数 =================================");
        XSSFSheet sheet = excel.getSheetAt(3);
        XSSFRow row = sheet.getRow(0);
        //b
        XSSFCell cell = row.getCell(1);
        String value = Util.getValueFromXssfcell(cell);
        map.put(Constants.SECTION_B, value);
        //h
        cell = row.getCell(4);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.SECTION_H, value);

        //v
        row = sheet.getRow(2);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.STRESS_CONDITION_V, value);
//        //原v
//        cell = row.getCell(4);
//        value = Util.getValueFromXssfcell(cell);
//        map.put(Constants.STRESS_CONDITION_OLDV,value);
//        //M
//        cell = row.getCell(7);
//        value = Util.getValueFromXssfcell(cell);
//        map.put(Constants.STRESS_CONDITION_M,value);

        //楼层高度
        row = sheet.getRow(4);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.FLOOR_H, value);

        //阻尼器高度
        cell = row.getCell(4);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.DAMPER_H, value);

        //混泥土等级
        row = sheet.getRow(6);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.CONCRETE_GRADE, value);

        //钢筋等级
        row = sheet.getRow(8);
        cell = row.getCell(1);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.STEEL_GRADE, value);

        //箍筋
        row = sheet.getRow(10);
        cell = row.getCell(2);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.HOOP_D, value);

        cell = row.getCell(5);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.HOOP_L, value);


        //纵筋
        row = sheet.getRow(12);
        cell = row.getCell(2);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.HOOP_VERTICl_D, value);

        cell = row.getCell(5);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.HOOP_VERTICl_L, value);

        row = sheet.getRow(14);
        cell = row.getCell(5);
        value = Util.getValueFromXssfcell(cell);
        map.put(Constants.PARAM_afS, value);

        //获取材料属性
        sheet = excel.getSheetAt(1);
        getMaterialParams(map.get(Constants.CONCRETE_GRADE).toString(), map.get(Constants.STEEL_GRADE).toString(), sheet, map);
        return map;
    }

    /**
     * 获取混泥土等级参数和钢筋等级参数
     *
     * @param concreteGrade
     * @param steelGrade
     * @return
     */
    public static void getMaterialParams(String concreteGrade, String steelGrade, XSSFSheet sheet, Map<String, Object> map) {
        XSSFRow row;
        XSSFCell cell;
        String value;
        Iterator it = sheet.iterator();
        it.next();
        it.next();
        while (it.hasNext()) {
            row = (XSSFRow) it.next();

            cell = row.getCell(0);
            value = Util.getValueFromXssfcell(cell);
            //混泥土参数添加
            if (map.get(Constants.CONCRETE_GRADE).equals(value)) {
                cell = row.getCell(1);
                value = Util.getValueFromXssfcell(cell);
                map.put(Constants.CONCRETE_GRADE_FCK, value);

                cell = row.getCell(2);
                value = Util.getValueFromXssfcell(cell);
                map.put(Constants.CONCRETE_GRADE_FC, value);

                cell = row.getCell(3);
                value = Util.getValueFromXssfcell(cell);
                map.put(Constants.CONCRETE_GRADE_FTK, value);

                cell = row.getCell(4);
                value = Util.getValueFromXssfcell(cell);
                map.put(Constants.CONCRETE_GRADE_FT, value);

                //已经有钢筋的
                if (map.containsKey(Constants.STEEL_GRADE_FYK)) {
                    break;
                }
            }

            cell = row.getCell(7);
            value = Util.getValueFromXssfcell(cell);
            //钢筋参数添加
            if (map.get(Constants.STEEL_GRADE).equals(value)) {
                cell = row.getCell(8);
                value = Util.getValueFromXssfcell(cell);
                map.put(Constants.STEEL_GRADE_FYK, value);

                cell = row.getCell(9);
                value = Util.getValueFromXssfcell(cell);
                map.put(Constants.STEEL_GRADE_FSTK, value);

                cell = row.getCell(10);
                value = Util.getValueFromXssfcell(cell);
                map.put(Constants.STEEL_GRADE_FYVK, value);
                //已经有混泥土的
                if (map.containsKey(Constants.CONCRETE_GRADE_FCK)) {
                    break;
                }
            }
        }
    }


    /**
     * 周期获取
     *
     * @param sheet
     */
    public static String[] getCycle(XSSFSheet sheet) {
        String[] data = new String[3];
        data[0] = Util.getValueFromXssfcell(sheet.getRow(2).getCell(1));
        data[1] = Util.getValueFromXssfcell(sheet.getRow(3).getCell(1));
        data[2] = Util.getValueFromXssfcell(sheet.getRow(4).getCell(1));
        System.out.println("周期对比：");
        GetExcelValue.arrayToString(data);
        return data;
    }

    /**
     * 质量获取
     *
     * @param sheet
     */
    public static String getQuality(XSSFSheet sheet) {
        String quality = Util.getValueFromXssfcell(sheet.getRow(0).getCell(0));
        System.out.println("==========================================");
        System.out.println("质量： " + quality );
        return quality;
    }


    /**
     * 非减震剪力获取
     * 减震剪力获取
     * @param sheet
     */
    public static List<String>[] getEarthquakeAndShear(XSSFSheet sheet) {
        List<String> fx = new ArrayList<>();
        List<String> fy = new ArrayList<>();
        XSSFRow row;
        Iterator it = sheet.iterator();
        it.next();
        it.next();
        while (it.hasNext()) {
            row = (XSSFRow) it.next();
            fx.add(Util.getValueFromXssfcell(row.getCell(3)));
            fy.add(Util.getValueFromXssfcell(row.getCell(10)));
        }

        System.out.println("X方向 ：");
        System.out.println(fx);
        System.out.println("Y方向 ：");
        System.out.println(fy);
        return new List[]{fx, fy};
    }


    /**
     * 获取层高，累计层高
     * @param sheet
     * @param type
     * @return
     */
    public static Double[] getFloorH(XSSFSheet sheet,String type){
        Iterator<Row> it = sheet.iterator();
        if (Constants.FLOOR_H.equals(type)){

        }else if (Constants.ACCOUNT_FLOOR_H.equals(type)){

        }
        return null;
    }


    //CAD模型编号  第一维表示楼层，第二维表示编号
    public static String[][] getCADModel(XSSFSheet sheet,String direction){
        if (Constants.X.equals(direction)){

        }else if (Constants.Y.equals(direction)){

        }
        return null;
    }


}
