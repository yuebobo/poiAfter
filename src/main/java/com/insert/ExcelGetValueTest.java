package com.insert;

import com.excel.sheet.ExcelBaseShear;
import com.excel.sheet.ExcelDamper;
import com.excel.sheet.ExcelFloorDisplaceShear;
import com.excel.sheet.ExcleModel;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelGetValueTest {

	public static void main(String[] args) throws IOException{
		Long st = System.currentTimeMillis();
		excle1();
		System.out.println("===================================================");
		System.out.println();
		Long ed = System.currentTimeMillis();
		System.out.println(ed - st);
		System.out.println("===================================================");
		
		excel2();
		System.out.println("===================================================");
		System.out.println();
		ed = System.currentTimeMillis();
		System.out.println(ed - st);
		System.out.println("===================================================");
		
		excel3();
		System.out.println("===================================================");
		System.out.println();
		ed = System.currentTimeMillis();
		System.out.println(ed - st);
		System.out.println("===================================================");
		
		excel4();
		System.out.println("===================================================");
		System.out.println();
		ed = System.currentTimeMillis();
		System.out.println(ed - st);
		System.out.println("===================================================");
		
		excel5();
		System.out.println("===================================================");
		System.out.println();
		ed = System.currentTimeMillis();
		System.out.println(ed - st);
		System.out.println("===================================================");
		
		excel6();
		System.out.println("===================================================");
		System.out.println();
		ed = System.currentTimeMillis();
		System.out.println(ed - st);
		System.out.println("===================================================");
		
		excel7();
		System.out.println("===================================================");
		System.out.println();
		ed = System.currentTimeMillis();
		System.out.println(ed - st);
		System.out.println("===================================================");
		
		excel8();
		System.out.println("===================================================");
		System.out.println();
		ed = System.currentTimeMillis();
		System.out.println(ed - st);
		System.out.println("===================================================");
	
//		excel9();
		System.out.println("===================================================");
		System.out.println();
		ed = System.currentTimeMillis();
		System.out.println(ed - st);
		System.out.println("===================================================");
		
//		excel10();
		System.out.println("===================================================");
		System.out.println();
		ed = System.currentTimeMillis();
		System.out.println(ed - st);
		System.out.println("===================================================");
	}
	
	
	/**
	 * 1.模型对比
	 * @throws IOException
	 */
	private static void excle1() throws IOException{
		String path1 = "C:/Users/YYong/Desktop/李祖玮工作资料/excle/工作簿1.xlsx";
		FileInputStream e = new FileInputStream(path1);
		XSSFWorkbook excel = new XSSFWorkbook(e);
		
		String value1 = ExcleModel.get_DEAD_LIVE(excel.getSheetAt(0));
		System.out.println("结构质量对比：");
		System.out.println(value1);
		
		String[] value2 = ExcleModel.getMODAL(excel.getSheetAt(1));
		System.out.println("周期对比：");
		arrayToString(value2);
		
		String[][] value3 = ExcleModel.getExEy(excel.getSheetAt(2));
		System.out.println("地震剪力对比:");
		arrayToString(value3[0]);
		arrayToString(value3[1]);
		e.close();
	}
	
	/**
	 * 2.基底剪力对比
	 * @throws IOException
	 */
	public static void excel2() throws IOException{
		String path1 = "C:/Users/YYong/Desktop/李祖玮工作资料/excle/工作簿2.xlsx";
		FileInputStream e = new FileInputStream(path1);
		XSSFWorkbook excel = new XSSFWorkbook(e);
		String[][] value = ExcelBaseShear.getE2_T5_R2(excel.getSheetAt(0));
		System.out.println("基低剪力对比");
		arrayToString(value[0]);
		arrayToString(value[1]);
		e.close();
	}
	
	/**
	 * 非减震结构层间位移
	 * @throws IOException
	 */
	public static void excel3() throws IOException {
		//原来工作簿5
		String path1 = "C:/Users/YYong/Desktop/李祖玮工作资料/excle/工作簿3.xlsx";
		FileInputStream e = new FileInputStream(path1);
		XSSFWorkbook excel = new XSSFWorkbook(e);
		String[][][] value = ExcelFloorDisplaceShear.getDisplace(excel.getSheetAt(0),2);
		System.out.println("非减震结构层间位移");
		printArrayDisplace(value);
		e.close();
	}
	
	/**
	 * 减震结构层间位移
	 * @throws IOException
	 */
	public static void excel4() throws IOException {
		//原来工作簿6
		String path1 = "C:/Users/YYong/Desktop/李祖玮工作资料/excle/工作簿4.xlsx";
		FileInputStream e = new FileInputStream(path1);
		XSSFWorkbook excel = new XSSFWorkbook(e);
		String[][][] value = ExcelFloorDisplaceShear.getDisplace(excel.getSheetAt(2),2);
		System.out.println("减震结构层间位移XY");
		printArrayDisplace(value);
		e.close();
	}
	
	/**
	 * 非减震结构层间剪力
	 * @throws IOException
	 */
	public static void excel5() throws IOException {
		String path1 = "C:/Users/YYong/Desktop/李祖玮工作资料/excle/工作簿3.xlsx";
		FileInputStream e = new FileInputStream(path1);
		XSSFWorkbook excel = new XSSFWorkbook(e);
		//原来sheet是0
		String[][][] value = ExcelFloorDisplaceShear.getShear(excel.getSheetAt(1),2);
		System.out.println("非减震结构层间剪力");
		printArrayShear(value);
		e.close();
	}
	
	/**
	 * 减震结构层间剪力
	 * @throws IOException
	 */
	public static void excel6() throws IOException {
		String path1 = "C:/Users/YYong/Desktop/李祖玮工作资料/excle/工作簿4.xlsx";
		FileInputStream e = new FileInputStream(path1);
		XSSFWorkbook excel = new XSSFWorkbook(e);
		String[][][] value = ExcelFloorDisplaceShear.getShear(excel.getSheetAt(3),2);
		System.out.println("减震结构层间剪力XY");
		printArrayShear(value);
		e.close();
	}
	
	/**
	 * 非减震结构层间位移角
	 * @throws IOException
	 */
	public static void excel7() throws IOException {
		String path1 = "C:/Users/YYong/Desktop/李祖玮工作资料/excle/工作簿9.xlsx";
		FileInputStream e = new FileInputStream(path1);
		XSSFWorkbook excel = new XSSFWorkbook(e);
		String[][][] value = ExcelFloorDisplaceShear.getDisplace(excel.getSheetAt(0),2);
		System.out.println("非减震结构层间角位移");
		printArrayDisplace(value);
		e.close();
	}
	
	/**
	 * 减震结构层间角位移
	 * @throws IOException
	 */
	public static void excel8() throws IOException {
		//原来工作簿10
		String path1 = "C:/Users/YYong/Desktop/李祖玮工作资料/excle/工作簿5.xlsx";
		FileInputStream e = new FileInputStream(path1);
		XSSFWorkbook excel = new XSSFWorkbook(e);
		//原来sheet为0
		String[][][] value = ExcelFloorDisplaceShear.getDisplace(excel.getSheetAt(2),2);
		System.out.println("减震结构层间角位移");
		printArrayDisplace(value);
		e.close();
	}
	
	
	/**
	 * 阻尼器 内力
	 * @throws IOException 
	 */
//	public static void excel9() throws IOException{
//		String path1 = "C:/Users/YYong/Desktop/李祖玮工作资料/excle/工作簿7.xlsx";
//		FileInputStream e = new FileInputStream(path1);
//		XSSFWorkbook excel = new XSSFWorkbook(e);
//		String[][][] value = ExcelDamper.getInnerForce(excel.getSheetAt(1),2);
//		System.out.println("阻尼器内力");
//		printArrayShear(value);
//		e.close();
//	}
	
	/**
	 * 阻尼器变形
	 * @throws IOException 
//	 */
//	public static void excel10() throws IOException{
//		String path1 = "C:/Users/YYong/Desktop/李祖玮工作资料/excle/工作簿7.xlsx";
//		FileInputStream e = new FileInputStream(path1);
//		XSSFWorkbook excel = new XSSFWorkbook(e);
//		String[][][] value = ExcelDamper.getShapeChange(excel.getSheetAt(0),2);
//		System.out.println("阻尼器形变");
//		printArrayDisplace(value);
//		e.close();
//	}
	
	private static void printArrayDisplace(String[][][] array){
		String[][] x = array[0];
		String[][] y = array[1];
		System.out.println("      T1X    T2X    T3X    T4X    T5X    R1X    R2X");
		for (int i = 0; i < x.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString(x[i]);
		}
		System.out.println();
		System.out.println("      T1Y    T2Y    T3Y    T4Y    T5Y    R1Y    R2Y");
		for (int i = 0; i < y.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString(y[i]);
		}
	}
	
	private static void printArrayShear(String[][][] array){
		String[][] x = array[0];
		String[][] y = array[1];
		System.out.println("      T1X     T2X      T3X       T4X       T5X       R1X       R2X");
		for (int i = 0; i < x.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString(x[i]);
		}
		System.out.println();
		System.out.println("      T1Y     T2Y      T3Y       T4Y       T5Y       R1Y       R2Y");
		for (int i = 0; i < y.length; i++) {
			System.out.print("第"+ (i + 1) + "层    ");
			arrayToString(y[i]);
		}
	}
	
	
	private static void arrayToString(String[] array){
		for (int i = 0; i < array.length; i++) {
			System.out.print(array[i] + ",  ");
		}
		System.out.println();
	}
}
