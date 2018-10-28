package com.excel.sheet;

import com.util.Util;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.math.BigDecimal;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * 1.模型对比
 *
 */
public class ExcleModel {

	/**
	 * 结构质量对比
	 *获取 在sheet为Base Reactions里获取数据
	 * 首行为DEAD和LIVE的G列里的数据
	 * 并按照 1*DEAD+0.5*LIVE的方式进行计算 获得返回值
	 * @return
	 * @throws IOException 
	 */
	public static String get_DEAD_LIVE(XSSFSheet sheet){
		Iterator it = sheet.iterator();
		XSSFRow  row;
		String DEAD = "";
		String LIVE = "";
		String firstValue ;
		while(it.hasNext()){
			row = (XSSFRow) it.next();
			try {
				firstValue = row.getCell(0).getStringCellValue();
			} catch (Exception e) {
				firstValue = row.getCell(0).getRawValue();
			}
			if("DEAD".equals(firstValue))                    DEAD = row.getCell(6).getRawValue();
			else if("LIVE".equals(firstValue))               LIVE = row.getCell(6).getRawValue();
			else if (!"".equals(DEAD) && !"".equals(LIVE))   break;
			else if("".equals(firstValue))	                 break;

		}
		BigDecimal dead = new BigDecimal(DEAD);
		BigDecimal live = new BigDecimal(LIVE);
		BigDecimal divisor = new BigDecimal("2");
		String value = live.divide(divisor).add(dead).toString();
		return Util.getPrecisionString(value, 2);
	}

	/**
	 * 周期对比
	 * 获取在sheet为Modal Participating Mass Ratios里获取数据
	 * 获取D列里的前三行数据
	 * @param sheet
	 * @return
	 */
	public static String[] getMODAL(XSSFSheet sheet){
		Iterator it = sheet.iterator();
		//跨过表头（三行）
		int count = 0;
		if(it.hasNext()) {it.next();count++;}
		if(it.hasNext()) {it.next();count++;}
		if(it.hasNext()) {it.next();count++;}
		if(count < 3)   return null;
		int i = 0;
		XSSFRow  row;
		String[] value = new String[3];
		Double vc;
		while(it.hasNext()){
			row = (XSSFRow) it.next();
			vc = row.getCell(3).getNumericCellValue();
			value[i] = Util.getPrecisionString(vc.toString(), 3);
			if(i++ == 2) break;
		}
		return value;
	}
	
	/**
	 * 地震剪力对比
	 * 获取在sheet为Section Cut Forces - Analysis里获取数据
	 * 获取B列为Ex和Ey时的F列和G列里的数据
	 * F列里的为X
	 * G列里的为Y
	 * 每层有两个值，取大值
	 * @param sheet
	 * @return
	 */
	public static String[][] getExEy(XSSFSheet sheet){
		Iterator it = sheet.iterator();
		//跨过表头（三行）
		int count = 0;
		if(it.hasNext()) {it.next();count++;}
		if(it.hasNext()) {it.next();count++;}
		if(it.hasNext()) {it.next();count++;}
		if(count < 3)   return null;
		
		Map<Integer, Double[]> map = new HashMap<>();
		XSSFRow  row;
		//保存ex 和 ey
		Double[] e ;
		Double ex;
		Double ey;
		String first;
		//标记 只取ex或ey的行
		String outputCase;
		//记录当前楼层
		int floor;
		//标记最高层
		int maxFloor = 0;
		while(it.hasNext()){
			row = (XSSFRow) it.next();
			try {
				first = row.getCell(0).getStringCellValue();
			} catch (Exception e1) {
				first = row.getCell(0).getRawValue();
			}
			if("".equals(first) || null == first) break;
			try {
				outputCase = row.getCell(1).getStringCellValue();
			} catch (Exception e1) {
				outputCase = row.getCell(1).getRawValue();
			}
			if(!"EX".equals(outputCase) && !"EY".equals(outputCase)) continue;
			floor = Integer.valueOf(first.substring(2));
			maxFloor = Math.max(floor, maxFloor);
			
			if(!map.containsKey(Integer.valueOf(floor))){
				e = new Double[2];
				e[0] = Math.abs(row.getCell(5).getNumericCellValue());
				e[1] = Math.abs(row.getCell(6).getNumericCellValue());
				map.put(floor, e);
			} else {
				e = map.get(floor);
				ex = Math.abs(row.getCell(5).getNumericCellValue());
				ey = Math.abs(row.getCell(6).getNumericCellValue());
				e[0] = Math.max(ex, e[0]);
				e[1] = Math.max(ey, e[1]);
			}
		}
		
		String[][] data = new String[2][maxFloor];
		for (int i = 1; i <= maxFloor; i++) {
			e = map.get(i);
			data[0][i-1] = Util.getPrecisionString(e[0].toString(), 2);
			data[1][i-1] = Util.getPrecisionString(e[1].toString(), 2);
		}
		return data;
	}
}
