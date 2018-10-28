package com.insert;

import com.file.GetExcelValue;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

/**
 * 读取数据，插入到excel模型中
 * @author 
 *
 */
public class InsertToExcel {

	private static String basePath;
	public static void getValueInsertExcel(String path) {
		basePath = path;
		try {
			FileInputStream shearFile = new FileInputStream(basePath + "\\工作簿0.xlsx");
			XSSFWorkbook shear = new XSSFWorkbook(shearFile);
			insertShear(shear.getSheetAt(0),shear.getSheetAt(1));
		} catch (Exception e) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+"楼层间剪力在插入到excel表格的过程中出现异常");
			e.printStackTrace();
		}

		try {
			FileInputStream displaceFile = new FileInputStream(basePath + "\\工作簿1.xlsx");
			XSSFWorkbook displace = new XSSFWorkbook(displaceFile);
			insertDisplace(displace.getSheetAt(0), displace.getSheetAt(1));
		} catch (Exception e) {
			System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"+"楼层间位移在插入到excel表格的过程中出现异常");
			e.printStackTrace();
		} 

	}
	
	/**
	 * 结构层间剪力
	 * @param sheetX
	 * @param sheetY
	 */
	private static void insertShear(XSSFSheet sheetX,XSSFSheet sheetY){
		//非减震结构层间剪力
		String[][][] shearNot = GetExcelValue.getShear(basePath+"\\excle\\工作簿3.xlsx",1);
		// 减震结构层间剪力
		String[][][] shear = GetExcelValue.getShear(basePath+"\\excle\\工作簿4.xlsx",3);

		int floor = Math.min(shear[0].length, shearNot[1].length);
		for (int i = 0; i < floor; i++) {
			sheetX.getRow(i+4).getCell(0).setCellValue(floor-i);
			sheetY.getRow(i+4).getCell(0).setCellValue(floor-i);
			for (int j = 0; j < 7; j++) {
				sheetX.getRow(i+4).getCell(j+1).setCellValue(shearNot[0][i][j]);
				sheetX.getRow(i+4).getCell(j+8).setCellValue(shear[0][i][j]);
				sheetY.getRow(i+4).getCell(j+1).setCellValue(shearNot[1][i][j]);
				sheetY.getRow(i+4).getCell(j+8).setCellValue(shear[1][i][j]);
			}
		}
	}

	/**
	 * 层间位移
	 * @param sheetX
	 * @param sheetY
	 */
	private static void insertDisplace(XSSFSheet sheetX,XSSFSheet sheetY){
		//非减震结构层间位移
		//原来工作簿5
		String[][][] displaceNot = GetExcelValue.getDisplace(basePath+"\\excle\\工作簿3.xlsx",0);
		// 减震结构层间位移
		//原来工作簿6
		String[][][] displace = GetExcelValue.getDisplace(basePath+"\\excle\\工作簿4.xlsx",2);
		
		int floor = Math.min(displace[0].length, displaceNot[1].length);
		for (int i = 0; i < floor; i++) {
			sheetX.getRow(i+4).getCell(0).setCellValue(floor-i);
			sheetY.getRow(i+4).getCell(0).setCellValue(floor-i);
			for (int j = 0; j < 7; j++) {
				sheetX.getRow(i+4).getCell(j+1).setCellValue(displaceNot[0][i][j]);
				sheetX.getRow(i+4).getCell(j+8).setCellValue(displace[0][i][j]);
				sheetY.getRow(i+4).getCell(j+1).setCellValue(displaceNot[1][i][j]);
				sheetY.getRow(i+4).getCell(j+8).setCellValue(displace[1][i][j]);
			}
		}
	}
}
