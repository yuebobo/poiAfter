package com.excel.sheet;

import com.util.Util;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * 2.基底剪力对比
 * 
 *
 */
public class ExcelBaseShear {
	
	
	/**
	 * 获取Base Reactions里的数据
	 * EX 和EY 对应的取D列和E列的值
	 * 
	 * T1X ~ T5X ，R1X ,R2X
	 * T1Y ~ T5Y ,R1Y ,R2Y
	 * 这14个值 表里每个都有两个值，需要取较大值
	 * X取 D列
	 * Y取 E列
	 * 
	 * @param sheet
	 * @return
	 */
	public static String[][] getE2_T5_R2(XSSFSheet sheet){

		Iterator it = sheet.iterator();
		//跨过表头（三行）
		int count = 0;
		if(it.hasNext()) {it.next();count++;}
		if(it.hasNext()) {it.next();count++;}
		if(it.hasNext()) {it.next();count++;}
		if(count < 3)   return null;
		
		Map<String, Double> map = new HashMap<>();
		XSSFRow  row;
		//保存ex 和 ey
		Double[] ex = new Double[2];
		Double[] ey = new Double[2];
		String first;
		while(it.hasNext()){
			row = (XSSFRow) it.next();
			try {
				first = row.getCell(0).getStringCellValue();
			} catch (Exception e) {
				first = row.getCell(0).getRawValue();
			}
			if("".equals(first) || null == first) break;
			if(!map.containsKey(first)){
				if(!"EX".equals(first) && !"EY".equals(first)){
					if(first.contains("X")){
						map.put(first, Math.abs(row.getCell(3).getNumericCellValue()));
					}else if(first.contains("Y")){
						map.put(first, Math.abs(row.getCell(4).getNumericCellValue()));
					}
				}else if("EX".equals(first)){
					ex[0] = Math.abs(row.getCell(3).getNumericCellValue());
					ex[1] = Math.abs(row.getCell(4).getNumericCellValue());
				}else if("EY".equals(first)){
					ey[0] = Math.abs(row.getCell(3).getNumericCellValue());
					ey[1] = Math.abs(row.getCell(4).getNumericCellValue());
				}
			} else {
				if(first.contains("X")){
					map.put(first,Math.max(map.get(first), Math.abs(row.getCell(3).getNumericCellValue())));
				}else if(first.contains("Y")){
					map.put(first,Math.max(map.get(first), Math.abs(row.getCell(4).getNumericCellValue())));
				}
			}
		}
		
		ex[0] = Math.max(ex[0], ey[0]);
		ey[1] = Math.max(ex[1], ey[1]);
		String[][] data = new String[2][8];
		
		//x方向
		data[0][0] = Util.getPrecisionString(ex[0],2); 
		data[0][1] = Util.getPrecisionString(map.get("T1X"),2);
		data[0][2] = Util.getPrecisionString(map.get("T2X"),2);
		data[0][3] = Util.getPrecisionString(map.get("T3X"),2);
		data[0][4] = Util.getPrecisionString(map.get("T4X"),2);
		data[0][5] = Util.getPrecisionString(map.get("T5X"),2);
		data[0][6] = Util.getPrecisionString(map.get("R1X"),2);
		data[0][7] = Util.getPrecisionString(map.get("R2X"),2);
		
		//y方向
		data[1][0] = Util.getPrecisionString(ey[1],2); 
		data[1][1] = Util.getPrecisionString(map.get("T1Y"),2);
		data[1][2] = Util.getPrecisionString(map.get("T2Y"),2);
		data[1][3] = Util.getPrecisionString(map.get("T3Y"),2);
		data[1][4] = Util.getPrecisionString(map.get("T4Y"),2);
		data[1][5] = Util.getPrecisionString(map.get("T5Y"),2);
		data[1][6] = Util.getPrecisionString(map.get("R1Y"),2);
		data[1][7] = Util.getPrecisionString(map.get("R2Y"),2);
		
		return data;
	}
}
