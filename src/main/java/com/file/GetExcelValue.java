package com.file;

import com.excel.sheet.*;
import com.util.Util;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 该类主要用来处理读取excel文件，进行流的操作
 * @author 
 *
 */
public class GetExcelValue {

	/**
	 * 1.模型对比
	 * @throws IOException
	 */
	public static Map<Integer, Object> getModel(String path) {
		FileInputStream e = null;
		try {
			e = new FileInputStream(path);
			XSSFWorkbook excel = new XSSFWorkbook(e);
			
			String value1 = ExcleModel.get_DEAD_LIVE(excel.getSheetAt(0));
			System.out.println("==========================================");
			System.out.println("结构质量对比：");
			System.out.println(value1);
			
			String[] value2 = ExcleModel.getMODAL(excel.getSheetAt(1));
			System.out.println("==========================================");
			System.out.println("周期对比：");
			arrayToString(value2);
			
			String[][] value3 = ExcleModel.getExEy(excel.getSheetAt(2));
			System.out.println("==========================================");
			System.out.println("地震剪力对比:");
			arrayToString(value3[0]);
			arrayToString(value3[1]);
			
			Map<Integer, Object> map = new HashMap<>();
			map.put(1, value1);
			map.put(2, value2);
			map.put(3, value3);
			return map;
			
		} catch (FileNotFoundException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"没找到");
			return null;
		} catch (IOException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"处理异常");
			return null;
		}
		finally {
			if(e != null){
				try {
					e.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}



	/**
	 * 从材料数据文件里获取周期，减震/非减震 剪力对比,质量
	 * @throws IOException
	 */
	public static Map<Integer, Object> getCycleAndFxFy(String path) {
		System.out.println("==========================================");
		System.out.println(path);
		FileInputStream e = null;
		try {
			e = new FileInputStream(path);
			XSSFWorkbook excel = new XSSFWorkbook(e);
			//周期获取
			String[] cecle = ExcelCaculateParams.getCycle(excel.getSheetAt(5));
			System.out.println("周期对比：");
			arrayToString(cecle);

			//质量获取
			String quality = ExcelCaculateParams.getQuality(excel.getSheetAt(4));
			System.out.println("==========================================");
			System.out.println("质量： "+quality);

			//减震剪力
			List[] fxFy = ExcelCaculateParams.getEarthquakeAndShear(excel.getSheetAt(7));
			System.out.println("减震剪力：");
			System.out.println("X方向 ：");
			System.out.println(fxFy[0]);
			System.out.println("Y方向 ：");
			System.out.println(fxFy[1]);

			//非减震剪力
			List[] notFxFy = ExcelCaculateParams.getEarthquakeAndShear(excel.getSheetAt(6));
			System.out.println("非减震剪力：");
			System.out.println("X方向 ：");
			System.out.println(fxFy[0]);
			System.out.println("Y方向 ：");
			System.out.println(fxFy[1]);

			Map<Integer,Object> map = new HashMap<>();
			map.put(1, cecle);
			map.put(2, quality);
			map.put(3, fxFy);
			map.put(4,notFxFy);
			return map;
		} catch (FileNotFoundException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"没找到");
			return null;
		} catch (IOException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"处理异常");
			return null;
		}
		finally {
			if(e != null){
				try {
					e.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}






	/**
	 * 基低剪力对比
	 * @param path
	 * @return
	 */
	public static String[][] getE2_T5_R2(String path){
		FileInputStream e = null;
		try {
			e = new FileInputStream(path);
			XSSFWorkbook excel = new XSSFWorkbook(e);
			String[][] value =  ExcelBaseShear.getE2_T5_R2(excel.getSheetAt(0));
			arrayToString(value[0]);
			arrayToString(value[1]);
			return value;
		} catch (FileNotFoundException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"没找到");
			return null;
		} catch (IOException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"处理异常");
			return null;
		}
		finally {
			if(e != null){
				try {
					e.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}


	/**
	 * 地震波信息
	 * @param path
	 * @return
	 */
	public static Map<String,String[]> getEarthquakeWaveInfo(String path,String[] number ){
		FileInputStream e = null;
		try {
			System.out.println("\n"+path);
			e = new FileInputStream(path);
			XSSFWorkbook excel = new XSSFWorkbook(e);
			Map<String,String[]> value =  ExcelEarthquakeWave.getEarthquakeWaveInfo(excel.getSheetAt(0),number);
			return value;
		} catch (FileNotFoundException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"没找到");
			return null;
		} catch (IOException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"处理异常");
			return null;
		}
		finally {
			if(e != null){
				try {
					e.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}


	/**
	 * 层间剪力
	 * @param path
	 * @return
	 */
	public static String[][][] getShear(String path,int sheet){
		FileInputStream e = null;
		try {
			e = new FileInputStream(path);
			XSSFWorkbook excel = new XSSFWorkbook(e);
			//原来sheet为0
			String[][][] value = ExcelFloorDisplaceShear.getShear(excel.getSheetAt(sheet),0);
			System.out.println("\n"+path);
			System.out.println("=================== 层间剪力 =========================");
			printArrayDisplace(value);
			return value;
		} catch (FileNotFoundException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"没找到");
			return null;
		} catch (IOException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"处理异常");
			return null;
		}
		finally {
			if(e != null){
				try {
					e.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}

	
	/**
	 * 层间位移
	 * @param path
	 * @return
	 */
	public static String[][][] getDisplace(String path,int sheet){
		FileInputStream e = null;
		try {
			e = new FileInputStream(path);
			XSSFWorkbook excel = new XSSFWorkbook(e);
			//原来sheet为0
			String[][][] value = ExcelFloorDisplaceShear.getDisplace(excel.getSheetAt(sheet),5);
			System.out.println("\n"+path);
			System.out.println("=================== 层间位移 =========================");
			printArrayDisplace(value);
			return value;
		} catch (FileNotFoundException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"没找到");
			return null;
		} catch (IOException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"处理异常");
			return null;
		}
		finally {
			if(e != null){
				try {
					e.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}


	/**
	 * X方向各地震波下X方向阻尼器耗能
	 * @param path
	 * @return
	 */
	public static Double[][][] getEarthquakeDamperDisEnergyX(String path){
		System.out.println("================================================================");
		System.out.println("\n"+path);
		System.out.println("X方向各地震波下X方向阻尼器耗能");
		Double[][][][] value = getEarthquakeDamperDisEnergy(path,6,7);
		Double[][] shape = value[0][0];
		Double[][] force = value[1][0];
//		return  new Double[][][]{shape,force};
		return  new Double[][][]{getHarf(shape,"X"),getHarf(force,"X")};
	}

	/**
	 * Y方向各地震波下Y方向阻尼器耗能
	 * @param path
	 * @return
	 */
	public static Double[][][] getEarthquakeDamperDisEnergyY(String path){
		System.out.println("================================================================");
		System.out.println("\n"+path);
		System.out.println("Y方向各地震波下Y方向阻尼器耗能");
		Double[][][][] value = getEarthquakeDamperDisEnergy(path,7,8);
		Double[][] shape = value[0][1];
		Double[][] force = value[1][1];
//		return  new Double[][][]{shape,force};
		return  new Double[][][]{getHarf(shape,"Y"),getHarf(force,"Y")};
	}

	//原来是x和y是分开在两个文件里，现在合并在一个文件里边
	//x取前一半，y取后一半
	private static Double[][] getHarf(Double[][] array,String xOy){
		int length = array.length;
		Double[][] data = new Double[length/2][];
		if (length % 2 == 0){
			int start = 0;
			if ("X".equals(xOy)) start = 0;
			if ("Y".equals(xOy)) start = length/2;
			for (int i = 0; i < length/2 ; i++){
				data[i] = array[start++];
			}
		}else {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ X方向和Y方向的总数不是偶数 $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
			return null;
		}
		return data;
	}

	/**
 	 * 各地震波下X/Y方向阻尼器耗能
	 * @param path
	 * @return
	 */
	private static Double[][][][] getEarthquakeDamperDisEnergy(String path, int valuePositionShape, int valuePositionForce){
		FileInputStream e = null;
		try {
			e = new FileInputStream(path);
			XSSFWorkbook excel = new XSSFWorkbook(e);
			//阻尼器变形
			Double[][][] shape = ExcelDamper.getDamperValue(excel.getSheetAt(0),2,valuePositionShape);
			//阻尼器内力
			Double[][][] force = ExcelDamper.getDamperValue(excel.getSheetAt(1),3,valuePositionForce);
			System.out.println("=================== 阻尼器变形  =========================");
			printArrayShear1(shape);
			System.out.println("=================== 阻尼器内力  =========================");
			printArrayDisplace1(force);
			Double[][][][] value = {shape,force};
			return value;
		}catch (FileNotFoundException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"没找到");
			return null;
		} catch (IOException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"处理异常");
			return null;
		}
		finally {
			if(e != null){
				try {
					e.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}
	
	/**
	 * 层间位移角
	 * @param path
	 * @return
	 */
	public static String[][][] getDisplaceAngle(String path,int sheet){
		FileInputStream e = null;
		try {
			e = new FileInputStream(path);
			XSSFWorkbook excel = new XSSFWorkbook(e);
			//原来sheet为0
			String[][][] value = ExcelFloorDisplaceShear.getDisplace(excel.getSheetAt(sheet),3);
			System.out.println("\n"+path);
			System.out.println("=================== 层间位移角 =========================");
			printArrayShear(value);
			return value;
		} catch (FileNotFoundException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"没找到");
			return null;
		} catch (IOException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"处理异常");
			return null;
		}
		finally {
			if(e != null){
				try {
					e.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}

	/**
	 * 获取楼层层高(或者累计层高)
	 * @param path
	 * @return
	 */
	public static List<String> getFloorHigh(String path, int col){
		System.out.println("\n"+path);
		FileInputStream e = null;
		try {
			e = new FileInputStream(path);
			XSSFWorkbook excel = new XSSFWorkbook(e);
			return ExcelFloorHigh.getFloorH(excel.getSheetAt(0),col);
		} catch (FileNotFoundException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"没找到");
			return null;
		} catch (IOException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$获取层高（累计高度）处理异常");
			return null;
		}
		finally {
			if(e != null){
				try {
					e.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}


	/**
	 * 根据参数，从excle里获取计算所需的数值
	 * 子结构框架梁
	 *
	 * @param path
	 * @return
	 */
	public static Map<String,Object>[] getCaculateParams(String path){
		FileInputStream e = null;
		try {
			System.out.println("\n"+path);
			System.out.println("=================== 计算最后几个表格 **  参数的获取 =========================");
			e = new FileInputStream(path);
			XSSFWorkbook excel = new XSSFWorkbook(e);

			//梁
			Map<String,Object> girderParams = ExcelCaculateParams.getParamsOfGirder(excel);
			//柱
			Map<String,Object> pillarParams = ExcelCaculateParams.getParamsOfPillar(excel);
			//悬臂
			Map<String,Object> cantileverParams = ExcelCaculateParams.getParamsOfCantilever(excel);

			return new Map[]{girderParams,pillarParams,cantileverParams};
		} catch (FileNotFoundException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"没找到");
			return null;
		} catch (IOException e1) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+path+"处理异常");
			return null;
		}
		finally {
			if(e != null){
				try {
					e.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}


	private static void printArrayDisplace(String[][][] array){
		String[][] x = array[0];
		String[][] y = array[1];
		System.out.println("     T1X   T2X   T3X    T4X   T5X    R1X   R2X");
		for (int i = 0; i < x.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString(x[i]);
		}
		System.out.println();
		System.out.println("     T1Y   T2Y   T3Y    T4Y   T5Y    R1Y   R2Y");
		for (int i = 0; i < y.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString(y[i]);
		}
	}
	
	private static void printArrayShear(String[][][] array){
		String[][] x = array[0];
		String[][] y = array[1];
		System.out.println("     T1X     T2X      T3X    T4X     T5X      R1X    R2X");
		for (int i = 0; i < x.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString(x[i]);
		}
		System.out.println();
		System.out.println("     T1X     T2X      T3X    T4X     T5X      R1X    R2X");
		for (int i = 0; i < y.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString(y[i]);
		}
	}
	
	private static void printArrayShear1(Double[][][] array){
		Double[][] x = array[0];
		Double[][] y = array[1];
		System.out.println("         T1X    T2X     T3X    T4X    T5X    R1X   R2X");
		for (int i = 0; i < x.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString1(x[i]);
		}
		System.out.println();
		System.out.println("         T1X    T2X     T3X    T4X    T5X    R1X   R2X");
		for (int i = 0; i < y.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString1(y[i]);
		}
	}

	private static void printArrayDisplace1(Double[][][] array){
		Double[][] x = array[0];
		Double[][] y = array[1];
		System.out.println("     T1X   T2X   T3X    T4X   T5X    R1X   R2X");
		for (int i = 0; i < x.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString1(x[i]);
		}
		System.out.println();
		System.out.println("     T1Y   T2Y   T3Y    T4Y   T5Y    R1Y   R2Y");
		for (int i = 0; i < y.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString1(y[i]);
		}
	}
	
	private static void arrayToString(String[] array){
		for (int i = 0; i < array.length; i++) {
			System.out.print(array[i] + ",  ");
		}
		System.out.println();
	}

	private static void arrayToString1(Double[] array){
		for (int i = 0; i < array.length; i++) {
			System.out.print(array[i] + ",  ");
		}
		System.out.println();
	}
}
