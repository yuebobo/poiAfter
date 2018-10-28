package com.excel.sheet;

import com.entity.ValueNote;
import com.util.Util;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * 地震波下   阻尼器出力位移
 * 
 * @author 
 *
 */
public class ExcelDamper {

	/**
	 * 内力
	 * @param sheet
	 * @param precision
	 * @return
	 */
//	public static String[][][] getInnerForce(XSSFSheet sheet,int precision,int flagPosition, int valuePosition){
////		return getDamperValue(sheet, precision, 3, 8);
//		return getDamperValue(sheet, precision, flagPosition, valuePosition);
//	}
	
	/**
	 * 变形
	 * @param sheet
	 * @param precision
	 * @return
	 */
//	public static String[][][] getShapeChange(XSSFSheet sheet,int precision, int flagPosition, int valuePosition){
////		return getDamperValue(sheet, precision, 2, 6);
//	}


	/**
	 * 地震波下阻尼器处理位移/阻尼器形变
	 * @param sheet
	 * @return
	 */
	public static Double[][][] getDamperValue(XSSFSheet sheet,int flagPosition , int valuePosition){
		Map<Integer, ValueNote> map = new HashMap<>();
		//三行表头
		Iterator it = sheet.iterator();
		it.next();
		it.next();
		it.next();
		
		XSSFRow  row;
		String firstCell;
		String flagCell;
		Double vaileCell = null;
		ValueNote valueNote;
//		当前楼层
		Integer floor;
//		最大楼层数
		Integer maxFloor = 0; 
		while (it.hasNext()) {
			row = (XSSFRow) it.next();
			try {
				firstCell  = row.getCell(1).getStringCellValue();
			} catch (Exception e) {
				firstCell  = row.getCell(1).getRawValue();
			}

			try {
				flagCell = row.getCell(flagPosition).getStringCellValue();
			} catch (Exception e) {
				flagCell = row.getCell(flagPosition).getRawValue();
			}
			if(firstCell == null) break;
			floor = Integer.valueOf(firstCell);
			maxFloor = Math.max(maxFloor, floor);
			if(flagCell.contains("X")) vaileCell = Math.abs(row.getCell(valuePosition).getNumericCellValue());
			if(flagCell.contains("Y")) vaileCell = Math.abs(row.getCell(valuePosition).getNumericCellValue());
			
			//该楼层还没有在map里
			if(!map.containsKey(floor)){
				map.put(floor, new ValueNote(floor.toString()));
			}
				valueNote = map.get(floor);
				Util.insertValue(valueNote, flagCell, vaileCell);
		}
		return Util.mapToArray1(map, maxFloor);
	}

}
