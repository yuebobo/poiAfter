package com.util;

import com.entity.Constants;
import com.entity.ValueNote;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

import java.math.BigDecimal;
import java.util.Collection;
import java.util.List;
import java.util.Map;


public class Util {

	public static void printData(Collection<ValueNote> valueNotes){
		if(valueNotes == null || valueNotes.size() < 1){
			return ;
		}
		for (ValueNote valueNote : valueNotes) {
			System.out.println("===================================================");
			System.out.println(valueNote.getFloor());

		}
	}

	public static void printArray(Object[] objects){
		int i = 0;
		for (Object object : objects){
			System.out.println( "   "+(++i)+"      "+object.toString());
		}
	}

	/**
	 * 四舍五入 返回String类型
	 * @param value
	 * @param precision 保留几位小数
	 * @return
	 */
	public static String getPrecisionString(String value,int precision){
		if(value == null || value == "") return "0";
		double v =Math.abs(Double.valueOf(value));
		if(precision < 0) return String.format("%.0f",v); 
		return String.format("%." + precision + "f",v); 
	}

	/**
	 * 四舍五入
	 * @param value
	 * @param precision
	 * @return
	 */
	public static String getPrecisionString(Double value,int precision){
		if(value == null ) return "0";
		if(precision < 0) return String.format("%.0f",value); 
		return String.format("%." + precision + "f",value); 
	}

	/**
	 * 四舍五入 返回Double类型
	 * @param value
	 * @param precision 保留几位小数
	 * @return
	 */
	public static Double getPrecisionDouble(String value,int precision){
		if(value == null || value == "") return 0D;
		double v =Math.abs(Double.valueOf(value));
		if(precision < 1)	return Double.valueOf(String.format("%.0f",v)); 
		return Double.valueOf((String.format("%." + precision + "f",v))); 
	}

	/**
	 * 对给定的值进行插入
	 * @param note
	 * @param flagCell
	 * @param vaileCell
	 */
	public static  void insertValue(ValueNote note ,String flagCell,double vaileCell ){
		if(note == null) return;

		switch (flagCell) {
		//x方向
		case Constants.T1X:
			note.setT1x(Math.max(note.getT1x(), vaileCell));
			break;
		case Constants.T2X:
			note.setT2x(Math.max(note.getT2x(), vaileCell));
			break;
		case Constants.T3X:
			note.setT3x(Math.max(note.getT3x(), vaileCell));
			break;
		case Constants.T4X:
			note.setT4x(Math.max(note.getT4x(), vaileCell));
			break;
		case Constants.T5X:
			note.setT5x(Math.max(note.getT5x(), vaileCell));
			break;
		case Constants.R1X:
			note.setR1x(Math.max(note.getR1x(), vaileCell));
			break;
		case Constants.R2X:
			note.setR2x(Math.max(note.getR2x(), vaileCell));
			break;

			//y方向
		case Constants.T1Y:
			note.setT1y(Math.max(note.getT1y(), vaileCell));
			break;
		case Constants.T2Y:
			note.setT2y(Math.max(note.getT2y(), vaileCell));
			break;
		case Constants.T3Y:
			note.setT3y(Math.max(note.getT3y(), vaileCell));
			break;
		case Constants.T4Y:
			note.setT4y(Math.max(note.getT4y(), vaileCell));
			break;
		case Constants.T5Y:
			note.setT5y(Math.max(note.getT5y(), vaileCell));
			break;
		case Constants.R1Y:
			note.setR1y(Math.max(note.getR1y(), vaileCell));
			break;
		case Constants.R2Y:
			note.setR2y(Math.max(note.getR2y(), vaileCell));
			break;			
		default:
			break;
		}
	}

	/**
	 * map转array
	 * @param map
	 * @param maxFloor
	 * @param precision
	 * @return
	 */
	public static String[][][] mapToArray(Map<Integer, ValueNote> map, int maxFloor,int precision){
		ValueNote valueNote;
		String[][][] data  = new String[2][maxFloor][7];
		for (int i = 1; i <= maxFloor; i++) {
			valueNote = map.get(i);
			data[0][i-1][0] = Util.getPrecisionString(valueNote.getT1x(),precision);
			data[0][i-1][1] = Util.getPrecisionString(valueNote.getT2x(),precision);
			data[0][i-1][2] = Util.getPrecisionString(valueNote.getT3x(),precision);
			data[0][i-1][3] = Util.getPrecisionString(valueNote.getT4x(),precision);
			data[0][i-1][4] = Util.getPrecisionString(valueNote.getT5x(),precision);
			data[0][i-1][5] = Util.getPrecisionString(valueNote.getR1x(),precision);
			data[0][i-1][6] = Util.getPrecisionString(valueNote.getR2x(),precision);

			data[1][i-1][0] = Util.getPrecisionString(valueNote.getT1y(),precision);
			data[1][i-1][1] = Util.getPrecisionString(valueNote.getT2y(),precision);
			data[1][i-1][2] = Util.getPrecisionString(valueNote.getT3y(),precision);
			data[1][i-1][3] = Util.getPrecisionString(valueNote.getT4y(),precision);
			data[1][i-1][4] = Util.getPrecisionString(valueNote.getT5y(),precision);
			data[1][i-1][5] = Util.getPrecisionString(valueNote.getR1y(),precision);
			data[1][i-1][6] = Util.getPrecisionString(valueNote.getR2y(),precision);
		}
		return data;
	}

	/**
	 * map转array
	 * 针对阻尼器 楼层数不连续
	 * @param map
	 * @param maxFloor
	 * @return
	 */
	public static Double[][][] mapToArray1(Map<Integer, ValueNote> map, int maxFloor){
		ValueNote valueNote;
		Double[][][] data  = new Double[2][map.size()][8];
		//楼层连续
		if(map.size() == maxFloor) {
			for (int i = 1; i <= maxFloor; i++) {
				valueNote = map.get(i);
				data[0][i-1][0] = Double.valueOf(i);
				data[0][i-1][1] = valueNote.getT1x();
				data[0][i-1][2] = valueNote.getT2x();
				data[0][i-1][3] = valueNote.getT3x();
				data[0][i-1][4] = valueNote.getT4x();
				data[0][i-1][5] = valueNote.getT5x();
				data[0][i-1][6] = valueNote.getR1x();
				data[0][i-1][7] = valueNote.getR2x();

				data[1][i-1][0] = Double.valueOf(i);
				data[1][i-1][1] = valueNote.getT1y();
				data[1][i-1][2] = valueNote.getT2y();
				data[1][i-1][3] = valueNote.getT3y();
				data[1][i-1][4] = valueNote.getT4y();
				data[1][i-1][5] = valueNote.getT5y();
				data[1][i-1][6] = valueNote.getR1y();
				data[1][i-1][7] = valueNote.getR2y();
			}
		}

		//楼层不连续
		int count = 0;
		if(map.size() < maxFloor){
			for (int i = 1; i <= maxFloor; i++) {
				if(!map.containsKey(i)) {
					continue;
				}
				valueNote = map.get(i);
				data[0][count][0] = Double.valueOf(i);
				data[0][count][1] = valueNote.getT1x();
				data[0][count][2] = valueNote.getT2x();
				data[0][count][3] = valueNote.getT3x();
				data[0][count][4] = valueNote.getT4x();
				data[0][count][5] = valueNote.getT5x();
				data[0][count][6] = valueNote.getR1x();
				data[0][count][7] = valueNote.getR2x();

				data[1][count][0] = Double.valueOf(i);
				data[1][count][1] = valueNote.getT1y();
				data[1][count][2] = valueNote.getT2y();
				data[1][count][3] = valueNote.getT3y();
				data[1][count][4] = valueNote.getT4y();
				data[1][count][5] = valueNote.getT5y();
				data[1][count][6] = valueNote.getR1y();
				data[1][count][7] = valueNote.getR2y();
				count++;
			}
		}
		return data;
	}


	/**
	 * 对两个值进行相乘，在除以2
	 * @param value1
	 * @param value2
	 * @return
	 */
	public static String multiplyAndHalf(String value1,String value2){
		BigDecimal v1 = new BigDecimal(value1);
		BigDecimal v2 = new BigDecimal(value2);
		BigDecimal v = v1.multiply(v2); 
		return getPrecisionString(v.divide(new BigDecimal(2)).toString(),0);
	}

	/**
	 * 计算两个数的差值 在除以第一个数
	 * @param value1
	 * @param precision
	 * @return
	 */
	public static String subAndDiv(String value1,String value2,int precision){
		double v1 = Double.valueOf(value1);
		double v2 = Double.valueOf(value2);
		double v = Math.abs(v1 - v2) * 100 / v1;
		return getPrecisionString(v, 2);
	}

	/**
	 * 层间剪力/位移比
	 * @param arrays1
	 * @param arrays2
	 * @return
	 */
	public static String[][][] getArrayProportion(String[][][] arrays1,String[][][] arrays2){
		int floor = Math.min(arrays1[0].length,arrays2[0].length);
		String[][][] arrays = new String[2][floor][8];
		Double proX = 0d;
		Double proY = 0d;
		double sumX = 0d;
		Double sumY = 0d;
		for (int i = 0; i < floor; i++) {
			for (int j = 0; j < 7; j++) {
				proX = Double.valueOf(arrays1[0][i][j])/Double.valueOf(arrays2[0][i][j]);
				arrays[0][i][j] = getPrecisionString(proX,2);
				sumX += proX;

				proY = Double.valueOf(arrays1[1][i][j])/Double.valueOf(arrays2[1][i][j]);
				arrays[1][i][j] = getPrecisionString(proY,2);
				sumY +=proY;
			}
			arrays[0][i][7] = getPrecisionString(sumX/7d, 2);
			arrays[1][i][7] = getPrecisionString(sumY/7d, 2);
			sumX = 0d;
			sumY = 0d;
		}
		return arrays;
	}

	/**
	 * 层高/层位移
	 * @param displace
	 * @param floorHight
	 * @return
	 */
	public static String[][][] getDisplaceAngle(String[][][] displace,List<String> floorHight){
		String[][][] returnValue = new String[2][displace[0].length][displace[0][0].length];
		Double floorH;
		int floor = displace[0].length;
		//楼层循环
		for (int i = 0; i < displace[0].length; i++) {
			floorH = Double.valueOf(floorHight.get(displace[0].length - i - 1));
			//循环列
			for (int j = 0; j < displace[0][0].length ; j++) {
				returnValue[0][floor - i - 1][j] = String.valueOf(floorH/Double.valueOf(displace[0][floor - i - 1][j]));
				returnValue[1][floor - i - 1][j] = String.valueOf(floorH/Double.valueOf(displace[1][floor - i - 1][j]));
			}
		}
		return returnValue;
	}

	/**
	 * 针对于 结构各层阻尼器最大出力及位移包络值汇总表
	 * 中间部分的三列数 不确定是哪三列数
	 * 该方法根据对应的数值如果不为0则表示为有效列
	 * @param arrays
	 * @return
	 */
	public  static  Integer[] getValueCol(Double[][] arrays){
		Integer[] valueCol = new Integer[3];
		double zero = 0d;
		int count = 0;
		for (int i = 0 ; i < arrays.length ; i++){
			for (int j = 1; j < arrays[i].length ; j++){
				if ( zero != arrays[i][j]){
					valueCol[count] = j;
					if(count++ == 2){
						return valueCol;
					}
				}
			}
		}
		return null;
	}

	public  static  Integer[] getValueCol(String[][] arrays){
		Integer[] valueCol = new Integer[3];
		double zero = 0d;
		int count = 0;
		for (int i = 0 ; i < arrays.length ; i++){
			for (int j = 0; j < arrays[i].length ; j++){
				if ( zero != Double.valueOf(arrays[i][j])){
					valueCol[count] = j;
					if(count++ == 2){
						return valueCol;
					}
				}
			}
		}
		return null;
	}

	/**
	 * 获取excel单元格里的值
	 * @param cell
	 * @return
	 */
	public static String getValueFromXssfcell(XSSFCell cell){
		if (cell == null) return null;
		try {
			return cell.getNumericCellValue()+"";
		}catch (Exception e){
			try {
				return cell.getStringCellValue();
			}catch (Exception e1){
				return cell.getRawValue();
			}
		}
	}


	/**
	 * 将值插入到word表格的单元格内
	 * @param cell
	 * @param text
	 */
	public static void insertValueToCell(XWPFTableCell cell, String text) {
		dealCell(cell, text, 10);
	}

	/**
	 * 将值插入到单元格内
	 *
	 * @param cell
	 * @param text
	 */
	private static void dealCell(XWPFTableCell cell, String text, int fontSize) {
		if (cell == null) {
			return;
		}
		cell.removeParagraph(0);
		XWPFParagraph pr = cell.addParagraph();
		XWPFRun rIO = pr.createRun();
		rIO.setFontFamily("Times New Roman");
		rIO.setColor("000000");
		rIO.setFontSize(fontSize);
		rIO.setText(text);
		cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
		pr.setAlignment(ParagraphAlignment.CENTER);
	}

}
