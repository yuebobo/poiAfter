package com.excel.sheet;

import com.entity.ValueNote;
import com.util.Util;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;


/**
 * 减震/非减震   层间结构位移/剪力
 * @author 
 *
 */
public class ExcelFloorDisplaceShear {

	/**
	 * 获取非减震结构层间位移
	 * @param sheet
	 * @return
	 */
	public static String[][][] getDisplace(XSSFSheet sheet,int precision){
		Map<Integer, ValueNote> map = new HashMap<>();
		//三行表头
		Iterator it = sheet.iterator();
		it.next();
		it.next();
		it.next();
		
		XSSFRow  row;
		String firstCell;
		String flagCell;
		Double vaileCell;
		ValueNote valueNote;
//		当前楼层
		Integer floor;
//		最大楼层数
		Integer maxFloor = 0; 
		while (it.hasNext()) {
			row = (XSSFRow) it.next();
			try {
				firstCell  = row.getCell(0).getStringCellValue();
			} catch (Exception e) {
				firstCell  = row.getCell(0).getRawValue();
			}
			try {
				flagCell = row.getCell(2).getStringCellValue();
			} catch (Exception e) {
				flagCell = row.getCell(2).getRawValue();
			}
			if(firstCell == null) break;
			floor = Integer.valueOf(firstCell.substring(2));
			maxFloor = Math.max(maxFloor, floor);
			vaileCell = Math.abs(row.getCell(5).getNumericCellValue());
			
			//该楼层还没有在map里
			if(!map.containsKey(floor)){
				map.put(floor, new ValueNote(floor.toString()));
			}
			
			if((firstCell.contains("X") && flagCell.contains("X")) || (firstCell.contains("Y") && flagCell.contains("Y"))){
				valueNote = map.get(floor);
				Util.insertValue(valueNote, flagCell, vaileCell);
			}
		}
		return Util.mapToArray(map, maxFloor,precision);
	}

	/**
	 * 获取减震结构层间剪力
	 * @param sheet
	 * @return
	 */
	public static String[][][] getShear(XSSFSheet sheet,int precision){
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
				firstCell  = row.getCell(0).getStringCellValue();
			} catch (Exception e) {
				firstCell  = row.getCell(0).getRawValue();
			}
			try {
				flagCell = row.getCell(1).getStringCellValue();
			} catch (Exception e) {
				flagCell = row.getCell(1).getRawValue();
			}
			if(firstCell == null) break;
			floor = Integer.valueOf(firstCell.substring(2));
			maxFloor = Math.max(maxFloor, floor);
			if(flagCell.contains("X")) vaileCell = Math.abs(row.getCell(4).getNumericCellValue());
			if(flagCell.contains("Y")) vaileCell = Math.abs(row.getCell(5).getNumericCellValue());
			
			//该楼层还没有在map里
			if(!map.containsKey(floor)){
				map.put(floor, new ValueNote(floor.toString()));
			}
				valueNote = map.get(floor);
				Util.insertValue(valueNote, flagCell, vaileCell);
		}
		return Util.mapToArray(map, maxFloor,precision);
	}
}
