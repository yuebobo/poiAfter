package com.mine;

import com.insert.InsertToWord;

import java.io.File;
import java.io.FileNotFoundException;

/**
 * Gui获取文件基本路径成功后，调用此类的方法
 *
 * @author
 */
public class MainDeal {

    /**
     * 获取基本路径
     * 拼接出其余文件的路径
     *
     * @param wordPath
     * @throws FileNotFoundException
     */
    public static void getBasePath(String wordPath) throws FileNotFoundException {
        String basePath;
        System.out.println(wordPath);
        if (wordPath.contains("\\")) {
            basePath = wordPath.substring(0, wordPath.lastIndexOf("\\"));
        } else if (wordPath.contains("/")) {
            basePath = wordPath.substring(0, wordPath.lastIndexOf("/"));
        } else {
            throw new FileNotFoundException();
        }
        System.out.println("基本路径：" + basePath);

//        boolean excel0IsExist = false;
        boolean excelDirectoryIsExist = false;
        boolean txtDirectoryIsExist = false;

        File file = new File(basePath);
        File[] fileList = file.listFiles();
        System.out.println("基本路径下的文件有：");
        for (File file2 : fileList) {
            System.out.println(file2.getPath());
            //将值直接插入到excel里，暂时不做
//			if(file2.isFile()){
//				if((basePath+"\\excel0.xlsx").equals(file2.getPath())) excel0IsExist = true;
//			}
            if (file2.isDirectory()) {
                if ((basePath + "\\excel").equals(file2.getPath())) {
                    excelDirectoryIsExist = true;
                }
                if ((basePath + "\\txt").equals(file2.getPath())) {
                    txtDirectoryIsExist = true;
                }
            }
        }

        if (excelDirectoryIsExist) {
            File file3 = new File(basePath);
            File[] file3s = file3.listFiles();
            if (file3s.length < 1) {
                System.out.println("excel文件夹里没有文件");
                return;
            }
        } else {
            System.out.println("缺少excel的文件夹");
        }

        if (txtDirectoryIsExist) {
            File file3 = new File(basePath);
            File[] file3s = file3.listFiles();
            if (file3s.length < 1) {
                System.out.println("txt文件夹里没有文件");
                return;
            }
        }else {
            System.out.println("缺少txt文件夹");
        }
        InsertToWord.getValueInsertWord(basePath, wordPath);
//		if(excel0IsExist){
//			try {
//				InsertToExcel.getValueInsertExcel(basePath);
//			} catch (Exception e) {
//				e.printStackTrace();
//			}
//		}
    }
}
