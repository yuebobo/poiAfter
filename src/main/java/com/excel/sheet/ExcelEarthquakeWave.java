package com.excel.sheet;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 *
 * @author : zyb
 * 时间 : 2018/8/4.
 */
public class ExcelEarthquakeWave {

    /**
     * 根据编号获取地震波信息
     * @param sheet
     * @param number
     * @return
     */
    public static  Map<String,String[]> getEarthquakeWaveInfo(XSSFSheet sheet,String[] number){
        String[] value ;
        Map<String,String[]> map = new HashMap<>();
        Iterator it = sheet.iterator();
        XSSFRow row ;
        String str = null;
        while (it.hasNext()){
            row = (XSSFRow) it.next();
            try {
                str = row.getCell(0).getStringCellValue();
            } catch (Exception e) {
               str = row.getCell(0).getRawValue();
            }
            if (str == null ||"".equals(str)){
                return map;
            }
            for (String s : number){
                if (str.equals(s)){
                    value = new String[3];
                    try {
                        value[0] = row.getCell(1).getStringCellValue();
                    } catch (Exception e) {
                        value[0] = row.getCell(1).getRawValue();
                    }
                    try {
                        value[1] = row.getCell(2).getStringCellValue();
                    } catch (Exception e) {
                        value[1] = row.getCell(2).getRawValue();
                    }
                    try {
                        value[2] = row.getCell(3).getRawValue();
                    } catch (Exception e) {
                        value[2] = row.getCell(3).getStringCellValue();
                    }
                    map.put(s,value);
                }
            }
        }
        return  map;
    }
}
