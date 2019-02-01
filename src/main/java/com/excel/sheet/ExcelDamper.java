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


	/**
	 * 地震波下阻尼器处理位移/阻尼器形变   x和y不做区分
	 * @param sheet
	 * @return
	 */
	public static Double[][] getDamperValueXAndY(XSSFSheet sheet,int flagPosition , int valuePosition){
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
//		当前编号
		Integer floor;
//		最大编号
		Integer maxFloor = 0;
		while (it.hasNext()) {
			row = (XSSFRow) it.next();
			try {
				firstCell  = row.getCell(0).getStringCellValue();
			} catch (Exception e) {
				firstCell  = row.getCell(0).getRawValue();
			}
			try {
				flagCell = row.getCell(flagPosition).getStringCellValue();
			} catch (Exception e) {
				flagCell = row.getCell(flagPosition).getRawValue();
			}
			if(firstCell == null) break;
			floor = Integer.valueOf(firstCell);
			maxFloor = Math.max(maxFloor, floor);
			if (flagCell.indexOf("TXB") >= 0 || flagCell.indexOf("TYB") >= 0) continue;
			flagCell = flagCell.replace("Y","X");
			vaileCell = row.getCell(valuePosition).getNumericCellValue();
			//该楼层还没有在map里
			if(!map.containsKey(floor)){
				map.put(floor, new ValueNote(floor.toString()));
			}
			valueNote = map.get(floor);
			Util.insertValue(valueNote, flagCell, vaileCell);
		}
		return Util.mapToArray1(map, maxFloor)[0];
	}


	/**
	 *  TXB TYB
	 * @param sheet
	 * @return
	 */
	public static Double[][] getDamperValueTB(XSSFSheet sheet,int flagPosition , int valuePosition){
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
//		当前编号
		Integer floor;
//		最大编号
		Integer maxFloor = 0;
		while (it.hasNext()) {
			row = (XSSFRow) it.next();
			try {
				firstCell  = row.getCell(0).getStringCellValue();
			} catch (Exception e) {
				firstCell  = row.getCell(0).getRawValue();
			}
			try {
				flagCell = row.getCell(flagPosition).getStringCellValue();
			} catch (Exception e) {
				flagCell = row.getCell(flagPosition).getRawValue();
			}
			if(firstCell == null) break;
			floor = Integer.valueOf(firstCell);
			maxFloor = Math.max(maxFloor, floor);
			if (flagCell.indexOf("TXB") < 0 && flagCell.indexOf("TYB") < 0) continue;
			//只有一列，存在T1X里
			flagCell = "T1X";
			vaileCell = row.getCell(valuePosition).getNumericCellValue();
			//该楼层还没有在map里
			if(!map.containsKey(floor)){
				map.put(floor, new ValueNote(floor.toString()));
			}
			valueNote = map.get(floor);
			Util.insertValue(valueNote, flagCell, vaileCell);
		}
		return Util.mapToArray1(map, maxFloor)[0];
	}



	/**
	 *  X1 Y2
	 * @param sheet
	 * @return
	 */
	public static Map<Integer,ValueNote>[] getDamperValueX_Y(XSSFSheet sheet,int flagPosition , int valuePositionX,int valuePositionY ){
		Map<Integer, ValueNote> mapX = new HashMap<>();
		Map<Integer, ValueNote> mapY = new HashMap<>();
		Map<Integer,ValueNote> map ;
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
//		当前编号
		Integer floor;
//		最大编号
		Integer maxFloor = 0;
		while (it.hasNext()) {
			row = (XSSFRow) it.next();
			try {
				firstCell  = row.getCell(0).getStringCellValue();
			} catch (Exception e) {
				firstCell  = row.getCell(0).getRawValue();
			}
			try {
				flagCell = row.getCell(flagPosition).getStringCellValue();
			} catch (Exception e) {
				flagCell = row.getCell(flagPosition).getRawValue();
			}
			if(firstCell == null || "".equals(firstCell)) break;
			firstCell = firstCell.trim();
			//只取X1 X2 Y1 Y2 类型的数据
			if (firstCell.indexOf("F") >= 0) continue;
			if (flagCell.indexOf("TXB") >= 0 && flagCell.indexOf("TYB") >= 0) continue;
			if (firstCell.indexOf("X") >= 0){
				if (flagCell.indexOf("Y") >= 0) continue;
				map = mapX;
				vaileCell = row.getCell(valuePositionX).getNumericCellValue();
			}else if (firstCell.indexOf("Y") >= 0){
				if (flagCell.indexOf("X") >= 0) continue;
				map = mapY;
				vaileCell = row.getCell(valuePositionY).getNumericCellValue();
			}else{
				continue;
			}
			floor = Integer.valueOf(firstCell.substring(1,firstCell.length()));
			maxFloor = Math.max(maxFloor, floor);
			//该楼层还没有在map里
			if(!map.containsKey(floor)){
				map.put(floor, new ValueNote(floor.toString()));
			}
			valueNote = map.get(floor);
			Util.insertValue(valueNote, flagCell, vaileCell);
		}
		return new Map[]{mapX,mapY};
	}


}
