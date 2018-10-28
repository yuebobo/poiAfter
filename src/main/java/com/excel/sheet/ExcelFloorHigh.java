package com.excel.sheet;

import com.util.Util;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Created by lizhongxiang on
 *
 * @author : lzx
 * 时间 : 2018/9/2.
 */
public class ExcelFloorHigh {

    public static List<String> getFloorH(XSSFSheet sheet, int col){
        Iterator it = sheet.iterator();
        it.next();
        List<String> list = new ArrayList<>();
        String h;
        XSSFRow row;
        XSSFCell cell;
        while(it.hasNext()) {
            row = (XSSFRow) it.next();
            cell = row.getCell(col);
            try {
                h = Util.getPrecisionString(cell.getNumericCellValue(),0);
            }catch (Exception e){
                try {
                    h = cell.getStringCellValue();
                }catch (Exception e1){
                    h = cell.getRawValue();
                }
            }
          list.add(h);
        }
        return list;
    }

}
