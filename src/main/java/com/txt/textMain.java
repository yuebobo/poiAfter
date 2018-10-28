package com.txt;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;

public class textMain {
	public static void main(String[] args){
		
		String path1 = "C:/Users/YYong/Desktop/pp/李祖玮工作资料/工作空间/txt/WMASS.txt";
//		String path2 = "C:/Users/YYong/Desktop/pp/李祖玮工作资料/工作空间/txt/WZQ.txt";
		String path2 = "C:\\Users\\lizhongxiang\\Desktop\\最终\\WZQ(2).txt";
		String path3 = "C:\\Users\\lizhongxiang\\Desktop\\最终\\wzq(3).txt";
		getLineNoteCount(path2);
		Long st = System.currentTimeMillis();
//		TxtGetValue.getValueFor1(path1);
		TxtGetValue.getValueFor3(path3);
//		List<String> j = TxtGetValue.getValueForFloorHeigh(path1);
//		System.out.println(j);
		
		Long ed = System.currentTimeMillis();
		System.out.println("用时："+(ed - st));
	}
	
	/**
	 * 获取每一行对应的元素数量，按照空格进行划分
	 * @param path
	 * @throws IOException
	 */
	public static void getLineNoteCount(String path) {
		try {
			FileInputStream fileIn = new FileInputStream(path);
			BufferedReader br = new BufferedReader(new InputStreamReader(fileIn,"GBK"));  
			String line;
			String[] strs;
			int i = 1;
			while ((line = br.readLine()) != null) {
				strs = line.trim().replaceAll(" +", " ").split(" ");
				System.out.println(i++ + " :"+strs.length);
			}
			br.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
