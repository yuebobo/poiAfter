package com.txt;

import com.entity.Constants;
import com.util.Util;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TxtGetValue {

	/**
	 * 获取WZQ.txt里的三组数据
	 * <p>
	 * <p>
	 * 第一个 表获取数据方式：
	 * 获取每一行的元素个数（多个空格合并成一个空格，按空格进行分开获得元素数量）
	 * 按照从上往下的方式逐行获取，当个数为8时，则到达第一个表的位置，if(i == 1 && strs.length == 8)
	 * 首个个数为8的行为表头，数据不进行留存，获取前三行数据的第二列元素  if(0 < count && count < 4 )
	 * 相应的计数变量增加
	 * <p>
	 * 第二个表数据的获取方式
	 * 在获取到第一的表数据的基础之上，再继续向下寻找个数为9的行  else if (i == 2 && strs.length == 9)
	 * 首个个数为9的行为表头行 在加上限制条件，首列为数字，这样确定数据的位置  if(strs.length > 0 && strs[0].matches("^[0-9]*$"))
	 * 知道获取到不符合条件为止
	 * 相应的计数变量增加
	 * <p>
	 * 第三个表的数据获取方式与第二个表相同
	 *
	 * @param txtPath
	 * @return
	 */
	public static Map<Integer, List<String>> getValueFor3(String txtPath) {

		List<String> theFirst = new ArrayList<>();
		List<String> theSecond = new ArrayList<>();
		List<String> theThird = new ArrayList<>();
		FileInputStream fileIn = null;
		BufferedReader br = null;
		try {
			fileIn = new FileInputStream(txtPath);
			br = new BufferedReader(new InputStreamReader(fileIn, "GBK"));
			String flageStr;
			String type = null;

			//根据前前三行的内容进行判断是那种类型的 WZQ
			for (int i = 0; i < 3 ; i++){
				flageStr = br.readLine();
				if (flageStr.lastIndexOf(Constants.TYPE_ONE) > 0){
					type = Constants.TYPE_ONE;
					break;
				}
				if (flageStr.lastIndexOf(Constants.TYPE_TWO) > 0){
					type = Constants.TYPE_TWO;
					break;
				}
			}
			if (Constants.TYPE_ONE.equals(type)){
				typeOne(br,theFirst,theSecond,theThird);
			} else if (Constants.TYPE_TWO.equals(type)){
				typeTwo(br,theFirst,theSecond,theThird);
			} else {
				System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$  无法确定WZQ是那种类型的 $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
				throw  new RuntimeException("无法确定WZQ是那种类型的");
			}

			Map<Integer, List<String>> values = new HashMap<>();
			System.out.println(txtPath);
			values.put(1, theFirst);
			values.put(2, theSecond);
			values.put(3, theThird);
			System.out.println();
			System.out.println("=======================================================");
			System.out.println("周期对比:");
			System.out.println(theFirst);
			System.out.println("地震剪力对比Fx:");
			System.out.println(theSecond);
			System.out.println("地震剪力对比Fy:");
			System.out.println(theThird);
			return values;
		} catch (Exception e) {
			System.out.println("$$$$$$$$$$$$" + txtPath + "出现异常");
			e.printStackTrace();
			return null;
		} finally {
			if (fileIn != null) {
				try {
					fileIn.close();
				} catch (IOException e) {
				}
			}
			if (br != null) {
				try {
					br.close();
				} catch (IOException e) {
				}
			}
		}
	}

	/**
	 * 获取WZQ里的数据
	 * 第一种情况
	 *  文件内容空头几行里包含了  ////////////////////////////////
	 * @param br
	 * @param theFirst
	 * @param theSecond
	 * @param theThird
	 * @throws IOException
	 */
	private static void typeOne(BufferedReader br,List<String> theFirst ,List<String> theSecond,List<String> theThird) throws IOException {
		String line;
		String[] strs;
		int i = 1;
		int count = 0;
		while ((line = br.readLine()) != null) {
			strs = line.trim().replaceAll(" +", " ").split(" ");
			if (i == 1 && strs.length == 8) {
				//获取第一个表里的数据  （周期三个数）
				if (0 < count && count < 4) {
					theFirst.add(strs[1]);
				}
				if (count > 3) {
					i++;
				}
				count++;
			} else if (i == 2 && strs.length == 9) {
				//获取第二个表里的数据
				if (strs.length > 0 && strs[0].matches("^[0-9]*$")) {
					theSecond.add(strs[3].substring(0, strs[3].length() - 1));
				} else if (theSecond.size() > 0) {
					i++;
				}
			}else if (i == 2 && theSecond.size() > 0){
				i++;
			} else if (i == 3 && strs.length == 9) {
				//获取第三个表里的数据
				if (strs.length > 0 && strs[0].matches("^[0-9]*$")) {
					theThird.add(strs[3].substring(0, strs[3].length() - 1));
				} else if (theThird.size() > 0) {
					i++;
				}
			} else if (i == 3 && theThird.size() > 0) {
				break;
			}else if (i == 4) {
				break;
			}
		}
	}


	/**
	 * 获取WZQ里的数据
	 * 第二种情况
	 *  文件内容空头几行里包含了  *****************************
	 * @param br
	 * @param theFirst
	 * @param theSecond
	 * @param theThird
	 * @throws IOException
	 */
	private static void typeTwo(BufferedReader br,List<String> theFirst ,List<String> theSecond,List<String> theThird) throws IOException {
		String line;
		String[] strs;
		int i = 1;
		int count = 0;
		boolean flage = false;
		while ((line = br.readLine()) != null) {
			strs = line.trim().replaceAll(" +", " ").split(" ");

			if (i == 1 && strs.length == 5) {
				//获取第一个表里的数据  （周期三个数）
				if (0 < count && count < 4) {
					theFirst.add(strs[1]);
				}
				if (count > 3) {
					i++;
				}
				count++;
				continue;
			}

			if (i == 2 && !flage){
				if (strs.length == 8){
					flage = true;
				}
				continue;
			}
			if (i == 2 && strs.length == 6) {
				//获取第二个表里的数据
				if (strs.length > 0 && strs[0].matches("^[0-9]*$")) {
					theSecond.add(strs[3].substring(0,strs[3].indexOf("(")));
				} else if (theSecond.size() > 0) {
					flage = false;
					i++;
				}
				continue;
			}
			if (i == 2 && theSecond.size() > 0){
				flage = false;
				i++;
				continue;
			}

			if (i == 3 && !flage){
				if (strs.length == 8){
					flage = true;
				}
				continue;
			}

			if (i == 3 && strs.length == 6) {
				//获取第三个表里的数据
				if (strs.length > 0 && strs[0].matches("^[0-9]*$")) {
					theThird.add(strs[3].substring(0, strs[3].indexOf("(")));
				} else if (theThird.size() > 0) {
					i++;
				}
				continue;
			}
			if (i == 3 && theThird.size() > 0){
				break;
			}else if (i == 4) {
				break;
			}
		}
	}

	/**
	 * 质量结构对比
	 * 获取WMASS.txt文件的一个数据
	 * 结构的总质量
	 * <p>
	 * 获取每一行的元素个数（多个空格合并成一个空格，按空格进行分开获得元素数量）
	 * 按照从上往下的方式逐行获取，当元素个数为12，则到达该数据位置上方表的表头  if(strs.length == 12)
	 * 再次往下获取 第四个元素个数为3 的行 ，并获取第三列数据，此数据就是目标数据 if(flag && strs.length == 3)   if(count == 4)
	 *
	 * @param txtPath
	 * @return
	 * @throws IOException
	 */
	public static String getValueFor1(String txtPath) {
		boolean flag = false;
		String line;
		String[] strs;
		int count = 1;
		FileInputStream fileIn = null;
		BufferedReader br = null;
		try {
			fileIn = new FileInputStream(txtPath);
			br = new BufferedReader(new InputStreamReader(fileIn, "GBK"));
			while ((line = br.readLine()) != null) {
				strs = line.trim().replaceAll(" +", " ").split(" ");
				if (strs.length == 12) {
					flag = true;
				}
				if (flag && strs.length == 3) {
					if (count == 4) {
						System.out.println("==========================================");
						System.out.println(txtPath);
						System.out.println("质量结构对比");
						System.out.println(strs[2]);
						return strs[2];
					} else {
						count++;
					}
				}
			}
			System.out.println("$$$$$$$$$$$$" + txtPath + "WMASS.txt文件里的    结构的总质量   数据没有获取到");
			return null;
		} catch (FileNotFoundException e) {
			System.out.println("$$$$$$$$$$$$" + txtPath + "没有找到");
			return null;
		} catch (IOException e) {
			System.out.println("$$$$$$$$$$$$" + txtPath + "处理异常");
			return null;
		} finally {
			if (fileIn != null) {
				try {
					fileIn.close();
				} catch (IOException e) {
				}
			}
			if (br != null) {
				try {
					br.close();
				} catch (IOException e) {
				}
			}
		}
	}

	/**
	 * 获取WMASS.txt文件的 层高
	 * 元素个数为15的行
	 *
	 * @param txtPath
	 * @return
	 */
	public static List<String> getValueForFloorHeigh(String txtPath) {
		System.out.println();
		System.out.println("获取层高：");
		String line;
		List<String> floorH = new ArrayList<>();
		String[] strs;
		FileInputStream fileIn = null;
		BufferedReader br = null;
		try {
			fileIn = new FileInputStream(txtPath);
			br = new BufferedReader(new InputStreamReader(fileIn, "GBK"));
			while ((line = br.readLine()) != null) {
				strs = line.trim().replaceAll(" +", " ").split(" ");
				if (strs.length == 15) {
					floorH.add(Util.getPrecisionString(Double.valueOf(strs[13])*1000,0));
				} else if (floorH.size() > 0) {
					return floorH;
				}
			}
			System.out.println("$$$$$$$$$$$$" + txtPath + "WMASS.txt文件里的   楼层高度  数据没有获取到");
			return null;
		} catch (FileNotFoundException e) {
			System.out.println("$$$$$$$$$$$$" + txtPath + "没有找到");
			return null;
		} catch (IOException e) {
			System.out.println("$$$$$$$$$$$$" + txtPath + "处理异常");
			return null;
		} finally {
			if (fileIn != null) {
				try {
					fileIn.close();
				} catch (IOException e) {
				}
			}
			if (br != null) {
				try {
					br.close();
				} catch (IOException e) {
				}
			}
		}
	}

	/**
	 * 获取地震波持时表 里的数据
	 *
	 * @param txtPath
	 * @return
	 */
	public static String[] earthquakeWave(String txtPath) {
		double dt = 0;
		FileInputStream fileIn = null;
		BufferedReader br = null;
		String line;
		int count = 0;
		double min = 0;
		double max = 0;
		boolean flag = false;
		try {
			fileIn = new FileInputStream(txtPath);
			br = new BufferedReader(new InputStreamReader(fileIn, "GBK"));
			//获取地震波时间间隔
			if ((line = br.readLine()) != null) {
				dt = Double.valueOf(line.substring(line.lastIndexOf("=") + 1, line.lastIndexOf(",")));
			}
			while ((line = br.readLine()) != null) {
				count++;
				if (isLessThen(line)) continue;
				max = count * dt;
				if (flag) continue;
				min = count * dt;
				flag = true;
			}
			return new String[]{String.valueOf(dt), String.valueOf(min), String.valueOf(max)};
		} catch (FileNotFoundException e) {
			System.out.println("$$$$$$$$$$$$" + txtPath + "没有找到");
			return null;
		} catch (IOException e) {
			System.out.println("$$$$$$$$$$$$" + txtPath + "处理异常");
			return null;
		} finally {
			if (fileIn != null) {
				try {
					fileIn.close();
				} catch (IOException e) {
				}
			}
			if (br != null) {
				try {
					br.close();
				} catch (IOException e) {
				}
			}
		}
	}

	/**
	 * 获取地震波信息里的第一行的两个数据
	 * @param txtPath
	 * @return
	 */
	public static String[] eathquakeWave1(String txtPath){
		String dt = null;
		String bt = null;
		FileInputStream fileIn = null;
		BufferedReader br = null;
		String line;
		try {
			fileIn = new FileInputStream(txtPath);
			br = new BufferedReader(new InputStreamReader(fileIn, "GBK"));
			//获取地震波时间间隔
			if ((line = br.readLine()) != null) {
				dt = line.substring(line.lastIndexOf("=") + 1, line.lastIndexOf(","));
				bt = line.split(",")[1].trim();
			}
			return new String[]{dt,bt};
		} catch (FileNotFoundException e) {
			System.out.println("$$$$$$$$$$$$" + txtPath + "没有找到");
			return null;
		} catch (IOException e) {
			System.out.println("$$$$$$$$$$$$" + txtPath + "处理异常");
			return null;
		} finally {
			if (fileIn != null) {
				try {
					fileIn.close();
				} catch (IOException e) {
				}
			}
			if (br != null) {
				try {
					br.close();
				} catch (IOException e) {
				}
			}
		}
	}

	// 对于  2.86E-07 类型的数进行判断 大小是否小于 0.1  小于返回true
	private static boolean isLessThen(String value) {
		value = value.trim().replaceAll(" +", "");
		if (value.contains("E-")) {
			String[] datas = value.split("E-");
			Double basic = Math.abs(Double.valueOf(datas[0]));
			int index = Integer.valueOf(datas[1]);
			if (index == 1) return 1 > basic;
			if (index > 1) return true;
			if (basic != 0) return false;
			return true;
		} else if (value.contains("E+")) {
			String[] datas = value.split("E+");
			Double basic = Math.abs(Double.valueOf(datas[0]));
			if (basic == 0) return true;
			return false;
		} else {
			return 0.1 > Double.valueOf(value);
		}
	}

	/**
	 * 获取第一个周期的数值
	 *
	 * @param txtPath
	 * @return
	 */
	public static String getSingleT(String txtPath) {
		String line;
		String[] strs;
		boolean flag = false;
		FileInputStream fileIn = null;
		BufferedReader br = null;
		try {
			fileIn = new FileInputStream(txtPath);
			br = new BufferedReader(new InputStreamReader(fileIn, "GBK"));
			while ((line = br.readLine()) != null) {
				strs = line.trim().replaceAll(" +", " ").split(" ");
				if (strs.length == 8) {
					if (flag)  return strs[1];
					flag = true;
				}
			}
			return null;
		} catch (Exception e) {
			System.out.println("$$$$ 获取单个周期的数值 $$$$$$$$" + txtPath + "出现异常");
			e.printStackTrace();
			return null;
		} finally {
			if (fileIn != null) {
				try {
					fileIn.close();
				} catch (IOException e) {
				}
			}
			if (br != null) {
				try {
					br.close();
				} catch (IOException e) {
				}
			}
		}
	}

}

