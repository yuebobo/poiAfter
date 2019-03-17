package com.insert;

import com.entity.BaseDate;
import com.file.GetExcelValue;
import com.txt.TxtGetValue;
import com.util.Util;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import sun.nio.cs.ext.MacHebrew;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.function.DoubleToLongFunction;

public class InsertToWord {

    private static String basePath;

    private static BaseDate data;

    //原来的模型中的CAD编号
    private static String[][] modelNo;

    //SAP模型中编号
    private static int[] SPANo;


    /**
     * 对word文件插入值
     *
     * @param path
     * @param wordPath
     */
    public static void getValueInsertWord(String path, String wordPath) {
        basePath = path;
        FileInputStream in = null;
        XWPFDocument word = null;
        FileOutputStream out = null;
        try {
            in = new FileInputStream(wordPath);
            word = new XWPFDocument(in);
            out = new FileOutputStream(basePath + "\\out" + System.currentTimeMillis() + ".docx");
            List<XWPFTable> tables = word.getTables();

            //初始化，获取初始数据
            init();

            //CAD模型编号表（SAP模型中编号）
            insertCADModelNo(tables.get(4));

            //模型对比三个表
            insertModelCompare(tables.get(6), tables.get(7), tables.get(8));

            //基低剪力对比          （非减震结构底部剪力对比表）
            insertBaseShearCopmpare(tables.get(9));

            //地震波信息
            insertEarthquakeWaveInfo(tables.get(10));
//
//            //地震波持时
            insertEarthquakeWave(tables.get(11));
//
//            //层间剪力对比
            insertFloorShearCopmare(tables.get(13));
//
//            //层间位移对比
            insertFloorDisplaceCompare(tables.get(14), tables.get(15));
//
//            //地震波下结构X/Y方向的弹性能
            insertElasticPropertyOfBaseEarthquake(tables.get(17), tables.get(18));
//
//            //各地震波下X/Y方向阻尼器耗能
            insertEarthquakeDamperDisEnergy(tables.get(4), tables.get(19), tables.get(20));
//
//            //结构附加阻尼比计算  该表的数据依赖与上边四个表的数据(此表要后处理)
            insertAnnexDamperRatio(tables.get(16), tables.get(17), tables.get(18), tables.get(19), tables.get(20));
//
//            //阻尼器出力与楼层剪力占比
            insertDamperFloorRatio(tables.get(21), tables.get(22), tables.get(4));
//
//            //层间位移角
            insertFloorDisplaceAngle(tables.get(24), tables.get(25));
//
//
//            //结构各层阻尼器最大出力及位移包络值汇总
//            //粘滞阻尼器性能规格表
            maxEarthquakeDapmerForceDisplace(tables.get(26), tables.get(27), tables.get(3));
//
            //金属阻尼器 表5
            insertMetalDamper(tables.get(5), tables.get(4));

//            //计算最后几个表里的值
//            //减震器周边子结构的设计计算方法
            calculateTable(tables.get(28), tables.get(29), tables.get(30));

        } catch (FileNotFoundException e) {
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + wordPath + "没找到");
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + wordPath + "处理异常");
            e.printStackTrace();
        } finally {
            try {
                word.write(out);
            } catch (IOException e) {
                System.out.println("输出流异常");
                e.printStackTrace();
            }
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * CAD模型编号表（SAP模型中编号）
     *
     * @param table4
     */
    private static void insertCADModelNo(XWPFTable table4) {
        System.out.println("=================================================");
        System.out.println("\n处理 CAD模型编号表（SAP模型中编号）");

        String[][] x_CAD = data.CAD_MODEL_X;
        String[][] y_CAD = data.CAD_MODEL_Y;
        List<String> listX = new ArrayList<>();
        List<String> listY = new ArrayList<>();

//      x方向
        boolean flage = true;
        String value;
        XWPFTableRow row;
        int countX = 0;
        for (int i = 0; flage; i++) {
            flage = false;

            for (int k = 0; k < x_CAD.length; k++) {
                try {
                    value = x_CAD[k][i];
                    if (null != value && !"".equals(value)) {
                        row = table4.createRow();
                        dealCellSM(row.getCell(0), value);
                        dealCellSM(row.getCell(1), ++countX + "");
                        listX.add(value);
                        flage = true;
                    }
                } catch (Exception e) {
                }
            }
        }
        //y方向
        int countY = 0;
        int count = countX;
        flage = true;
        for (int i = 0; flage; i++) {
            flage = false;
            for (int k = 0; k < y_CAD.length; k++) {
                try {
                    value = y_CAD[k][i];
                    if (null != value && !"".equals(value)) {
                        row = table4.getRow(countY + 1);
                        dealCellSM(row.getCell(2), value);
                        dealCellSM(row.getCell(3), ++count + "");
                        countY++;
                        listY.add(value);
                        if (countY > countX) {
                            System.out.println("$$$$$$$$$$$$$$$$$$$$ CAD模型里X方向和Y方向的数量不一致 $$$$$$$$$$$$$$$$$$$$$$$$$");
                            flage = false;
                            break;
                        }
                        flage = true;
                    }
                } catch (Exception e) {
                }
            }
        }
        if (countX != countX) {
            System.out.println("$$$$$$$$$$$$$$$$$$$$ CAD模型里X方向和Y方向的数量不一致 $$$$$$$$$$$$$$$$$$$$$$$$$");
        }
        //CAD模型中的编号
        modelNo = new String[2][Math.max(listX.size(), listY.size())];
        listX.toArray(modelNo[0]);
        listY.toArray(modelNo[1]);
        //SAP模型中编号
        SPANo = new int[countX + countY];
        for (int i = 0; i < SPANo.length; ) {
            SPANo[i] = ++i;
        }
        System.out.println("============== CAD 编号 ====================");
        System.out.println(listX);
        System.out.println(listY);
        System.out.println("============== SPA 编号 ====================");
        System.out.println(Arrays.asList(SPANo));
    }

    /**
     * 金属阻尼器表格的处理
     *
     * @param table5
     * @param table2
     */
    private static void insertMetalDamper(XWPFTable table5, XWPFTable table2) {
        System.out.println("=================================================");
        System.out.println("金属阻尼器表格的处理");
        try {
            //获取CAD 编号
            String[][] modelValue = getModelNo(table2);
            String[][] x_CAD = data.CAD_MODEL_X;
            String[][] y_CAD = data.CAD_MODEL_Y;

            //获取每一层对应的编号位置  位置从0开始
            Map<Integer, List<Integer>> map = getFloorOnPositionOfModelNO(modelValue);

            //X方向
            //原来是工作簿4
            Double[][][] valueX = GetExcelValue.getEarthquakeDamperDisEnergyX(basePath + "\\excel\\工作簿3.xlsx");
            //阻尼器形变
            Double[][] shapeX = valueX[0];
            //阻尼器内力
            Double[][] forceX = valueX[1];

            //Y方向
            //原来是工作簿4
            Double[][][] valueY = GetExcelValue.getEarthquakeDamperDisEnergyY(basePath + "\\excel\\工作簿3.xlsx");
            //阻尼器形变
            Double[][] shapeY = valueY[0];
            //阻尼器内力
            Double[][] forceY = valueY[1];

            //获取层高数据  此处数值单位为 豪米
            Double[] floorH = data.FLOOR_H;
            if (map.size() != floorH.length) {
                System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$  金属阻尼器表格  CAD模型编号的楼层数量与层高表里的楼层数量不一致   $$$$$$$$$$$$$  ");
            }
            Integer floor = floorH.length;

            //每一层对应的金属阻尼器弹性时程平均出力 和	金属阻尼器弹性时程平均位移
            Map<Integer, Double> forceXAvg = getAvgValueGroupByFloorFromTable( map,forceX);
            Map<Integer, Double> forceYAvg = getAvgValueGroupByFloorFromTable( map,forceY);
            Map<Integer, Double> shapeXAvg = getAvgValueGroupByFloorFromTable(map,shapeX);
            Map<Integer, Double> shapeYAvg = getAvgValueGroupByFloorFromTable(map,shapeY);
//
//表头四行，
            //数据行以表格第五行数据为模版进行加入数据
            //新加入的行都插入到第六行
            //最后模板行在数据行的最下边，数据插入完成将其删除
            XWPFTableRow row500 = table5.getRow(4);
            XWPFTableRow row;
            // Y方向
            int z = forceYAvg.size();
            for (Integer i = floor; i >= 1; i--) {
                for (int k = y_CAD[i - 1].length - 1; k >= 0; k--) {
                    table5.addRow(row500, 5);
                    row = table5.getRow(5);
                    dealCellSM(row.getCell(0), y_CAD[i - 1][k]);
                    insertTable(row, i, --z,floorH, forceYAvg, shapeYAvg);
                }
            }
            z = forceXAvg.size();
            //X方向
            for (Integer i = floor; i >= 1; i--) {
                for (int k = x_CAD[i - 1].length - 1; k >= 0; k--) {
                    table5.addRow(row500, 5);
                    row = table5.getRow(5);
                    dealCellSM(row.getCell(0), x_CAD[i - 1][k]);
                    insertTable(row, i, --z,floorH, forceXAvg, shapeXAvg);
                }
            }
            table5.removeRow(table5.getRows().size() - 1);
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$金属阻尼器表格的处理时发生异常");
        }
    }

    private static void insertTable(XWPFTableRow row, int i, int k,Double[] floorH, Map<Integer, Double> forceAvg, Map<Integer, Double> shapeAvg) {
        Double v;
        dealCellSM(row.getCell(1), i + ""); //楼层
        dealCellSM(row.getCell(2), floorH[i - 1].intValue() + "");
        dealCellSM(row.getCell(3), Util.getPrecisionString(forceAvg.get(k), 0));
        dealCellSM(row.getCell(4), Util.getPrecisionString(shapeAvg.get(k), 2));
        dealCellSM(row.getCell(5), Util.getPrecisionString(Double.valueOf(row.getCell(3).getText()) / Double.valueOf(row.getCell(4).getText()), 0));

        v = 0.4 * Double.valueOf(row.getCell(6).getText()) * Double.valueOf(row.getCell(7).getText()) * Double.valueOf(row.getCell(8).getText()) / (600 * (floorH[i - 1] - 600));
        dealCellSM(row.getCell(9), Util.getPrecisionString(v, 0));

        v = 3 * Double.valueOf(row.getCell(6).getText()) * Double.valueOf(row.getCell(7).getText()) * Math.pow(Double.valueOf(row.getCell(8).getText()), 3) / (12000 * Math.pow(((floorH[i - 1] - 600) / 2), 3));
        dealCellSM(row.getCell(10), Util.getPrecisionString(v, 0));

        v = 1 / ((1 / Double.valueOf(row.getCell(9).getText())) + (1 / Double.valueOf(row.getCell(10).getText())));
        dealCellSM(row.getCell(11), Util.getPrecisionString(v, 0));
        v = Double.valueOf(row.getCell(11).getText()) / 2;
        dealCellSM(row.getCell(12), Util.getPrecisionString(v, 0));

        v = 1 / ((1 / Double.valueOf(row.getCell(5).getText())) + (1 / Double.valueOf(row.getCell(12).getText())));
        dealCellSM(row.getCell(14), Util.getPrecisionString(v, 0));

        v = 1000 * Double.valueOf(row.getCell(14).getText()) * (
                (Math.pow(floorH[i - 1], 3) / (Double.valueOf(row.getCell(13).getText()) * Math.pow(Double.valueOf(row.getCell(16).getText()), 3))) +
                        (1.2 * floorH[i - 1] / (0.4 * Double.valueOf(row.getCell(13).getText()) * Double.valueOf(row.getCell(16).getText()))));
        dealCellSM(row.getCell(15), Util.getPrecisionString(v, 0));

        v = Double.valueOf(row.getCell(13).getText()) * 0.4 * Double.valueOf(row.getCell(16).getText()) * Double.valueOf(row.getCell(15).getText()) / (1200 * floorH[i - 1]);
        dealCellSM(row.getCell(17), Util.getPrecisionString(v, 2));

        v = 12 * Double.valueOf(row.getCell(13).getText()) * Double.valueOf(row.getCell(15).getText()) * Math.pow(Double.valueOf(row.getCell(16).getText()), 3) / (12000 * Math.pow(floorH[i - 1], 3));
        dealCellSM(row.getCell(18), Util.getPrecisionString(v, 2));

        v = 1 / ((1 / Double.valueOf(row.getCell(17).getText())) + (1 / Double.valueOf(row.getCell(18).getText())));
        dealCellSM(row.getCell(19), Util.getPrecisionString(v, 2));

        v = Math.abs(Double.valueOf(row.getCell(19).getText()) - Double.valueOf(row.getCell(14).getText()))
                / Math.max(Double.valueOf(row.getCell(19).getText()), Double.valueOf(row.getCell(14).getText()));
        dealCellSM(row.getCell(20), v < 0.05 ? "满足" : "不满足");
    }

    /**
     * 模型对比三个表的值得插入
     *
     * @param table3
     * @param table4
     * @param table5
     */
    private static void insertModelCompare(XWPFTable table3, XWPFTable table4, XWPFTable table5) {
        System.out.println("\n处理模型对比三张表");
        try {
            //1.结构周期对比   //地震剪力对比Fx  //地震剪力对比Fy
            //从记事本WZQ中 获取
//            Map<Integer, List<String>> map = TxtGetValue.getValueFor3(basePath + "\\txt\\2.txt");


            //2.结构质量对比  周期对比    地震剪力对比
            Map<Integer, Object> modelMap = GetExcelValue.getModel(basePath + "\\excel\\工作簿1.xlsx");

            //3.结构质量对比    从记事本WMASS中获取 ：结构的总质量       (表3 PKPM)
//            String qualityOfStructure1 = TxtGetValue.getValueFor1(basePath + "\\txt\\3.txt");

            //  == == 从材料数据文件里获取周期，减震剪力对比,质量
//            Map<Integer, Object> map = GetExcelValue.getCycleAndFxFy(basePath + "\\excel\\材料数据.xlsx");

            //4.周期对比                           （表2   PKPM）
//            String[] cycle1 = (String[]) map.get(1);
            String[] cycle1 = data.CECLE;

            //结构质量对比
//            String qualityOfStructure1 = (String) map.get(2);
            String qualityOfStructure1 = data.QUALITY;
            //
//            List[] fxFy = (List[]) map.get(3);
            List[] fxFy = data.FX_FY;

            //5.地震剪力对比Fx    （表1   PKPM   X）
            List<String> fx = fxFy[0];
            //6."地震剪力对比Fy"  （表1   PKPM   Y）
            List<String> fy = fxFy[1];

            //7.结构质量对比               (表3 SAP2000)
            String qualityOfStructure2 = (String) modelMap.get(1);
            //8.周期对比                    （表2   SAP2000）
            String[] cycle2 = (String[]) modelMap.get(2);
            //9.地震剪力对比 xy  (表1 SAP20000 XY)
            String[][] f = (String[][]) modelMap.get(3);

            //===================================  表3  结构质量对比  ====================================
            XWPFTableRow row3 = table3.getRow(1);
            dealCellBig(row3.getCell(0), Util.getPrecisionString(qualityOfStructure1, 1));
            dealCellBig(row3.getCell(1), Util.getPrecisionString(Double.valueOf(qualityOfStructure2) / 10, 1));
            //计算差值并填入
            dealCellBig(row3.getCell(2), Util.subAndDiv(qualityOfStructure1, String.valueOf(Double.valueOf(qualityOfStructure2) / 10), 2));

            //===================================  表4  结构周期对比  ====================================
            XWPFTableRow row4;
            for (int i = 1; i < 4; i++) {
                row4 = table4.getRow(i);
                dealCellBig(row4.getCell(1), Util.getPrecisionString(cycle1[i - 1], 3));
                dealCellBig(row4.getCell(2), Util.getPrecisionString(cycle2[i - 1], 3));
                dealCellBig(row4.getCell(3), Util.subAndDiv(cycle1[i - 1], cycle2[i - 1], 2));
            }

            //===================================  表5  结构地震剪力对比  ====================================
            int floor5 = Math.min(fx.size(), f[0].length);
            XWPFTableRow row5;
            for (int i = 0; i < floor5; i++) {

                //按照表头的单元格数进行添加
                table5.createRow();
                row5 = table5.getRow(i + 2);

                //表头与表身差三个单元格
                row5.addNewTableCell();
                row5.addNewTableCell();
                row5.addNewTableCell();

                dealCellBig(row5.getCell(0), String.valueOf(floor5 - i));
                dealCellBig(row5.getCell(1), Util.getPrecisionString(fx.get(i), 0));
                dealCellBig(row5.getCell(2), Util.getPrecisionString(fy.get(i), 0));
                dealCellBig(row5.getCell(3), Util.getPrecisionString(f[0][floor5 - i - 1], 0));
                dealCellBig(row5.getCell(4), Util.getPrecisionString(f[1][floor5 - i - 1], 0));
                dealCellBig(row5.getCell(5), Util.subAndDiv(fx.get(i), f[0][floor5 - i - 1], 2));
                dealCellBig(row5.getCell(6), Util.subAndDiv(fy.get(i), f[1][floor5 - i - 1], 2));
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$处理模型对比三张表时发生异常");
        }
    }


    /**
     * 基低剪力对比          （非减震结构底部剪力对比表）
     *
     * @param table6
     */
    private static void insertBaseShearCopmpare(XWPFTable table6) {
        System.out.println("======================================================");
        System.out.println("\n处理 非减震结构底部剪力对比表");
        try {
            //e2T5R2[0] 为x方向  从0到7   依次为反应普  T1-T5  R1-R2
            //e2T5R2[1] 为y方向
            String[][] e2T5R2 = GetExcelValue.getE2_T5_R2(basePath + "\\excel\\工作簿2.xlsx");

//            Map<Integer, Object> map = GetExcelValue.getCycleAndFxFy(basePath + "\\excel\\材料数据.xlsx");
//            List[] notFxFy = (List[]) map.get(4);
            List[] notFxFy = data.NOT_FX_FY;

            List<String> fx = notFxFy[0];
            List<String> fy = notFxFy[1];
            e2T5R2[0][0] = fx.get(fx.size() - 1);
            e2T5R2[1][0] = fy.get(fy.size() - 1);

            //用于计算T1-R2的平均值
            double x = 0d;
            double y = 0d;
            for (int i = 0; i < 8; i++) {
                dealCellSM(table6.getRow(2).getCell(i + 2), Util.getPrecisionString(e2T5R2[0][i], 0));
                dealCellSM(table6.getRow(3).getCell(i + 2), Util.getPrecisionString(e2T5R2[1][i], 0));

                dealCellSM(table6.getRow(4).getCell(i + 2), Util.getPrecisionString(Double.valueOf(e2T5R2[0][i]) / Double.valueOf(e2T5R2[0][0]), 2));
                dealCellSM(table6.getRow(5).getCell(i + 2), Util.getPrecisionString(Double.valueOf(e2T5R2[1][i]) / Double.valueOf(e2T5R2[1][0]), 2));
                x += Double.valueOf(e2T5R2[0][i]);
                y += Double.valueOf(e2T5R2[1][i]);
            }
            x -= Double.valueOf(e2T5R2[0][0]);
            y -= Double.valueOf(e2T5R2[1][0]);
            dealCellSM(table6.getRow(2).getCell(10), Util.getPrecisionString(x / 7d, 0));
            dealCellSM(table6.getRow(3).getCell(10), Util.getPrecisionString(y / 7d, 0));
            dealCellSM(table6.getRow(4).getCell(10), Util.getPrecisionString((x / 7d) / Double.valueOf(e2T5R2[0][0]), 2));
            dealCellSM(table6.getRow(5).getCell(10), Util.getPrecisionString((y / 7d) / Double.valueOf(e2T5R2[1][0]), 2));
        } catch (Exception e) {
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + "处理非减震结构底部剪力对比表发生异常");
            e.printStackTrace();
        }
    }

    /**
     * 地震波信息
     *
     * @param table7
     */
    private static void insertEarthquakeWaveInfo(XWPFTable table7) {
        System.out.println("=================================================");
        System.out.println("\n处理 地震波信息表");
        String path = basePath + "\\excel\\地震波信息.xlsx";
        try {
            //获取word表里的编号数据 T1~T5
            String[] number = new String[5];
            for (int i = 0; i < 5; i++) {
                number[i] = table7.getRow(i + 2).getCell(1).getText();
            }

            String[] value;
            XWPFTableRow row;

            //地点,测震台站,发震时间
            //获取数据
            Map<String, String[]> maps = GetExcelValue.getEarthquakeWaveInfo(path, number);
            for (int i = 0; i < 5; i++) {
                value = maps.get(number[i]);
                if (value == null) {
                    System.out.println("编号为：" + number[i] + " 的数据为空");
                    continue;
                }
                row = table7.getRow(i + 2);
                dealCellSM(row.getCell(2), value[0]);
                dealCellSM(row.getCell(3), value[1]);
                dealCellSM(row.getCell(4), value[2]);
            }

            row = table7.getRow(2);
            String str1 = row.getCell(5).getText();
            String str2 = row.getCell(6).getText();

            //采集间隔,采点数量  以及前两列数据
            for (int i = 1; i < 8; i++) {
                row = table7.getRow(i + 1);
                if (i < 6) {
                    dealCellSM(row.getCell(5), str1);
                    dealCellSM(row.getCell(6), str2);
                    path = basePath + "\\txt\\T" + (i) + ".txt";
                    value = TxtGetValue.eathquakeWave1(path);
                    dealCellSM(row.getCell(7), value[0]);
                    dealCellSM(row.getCell(8), value[1]);
                } else {
                    dealCellSM(row.getCell(3), str1);
                    dealCellSM(row.getCell(4), str2);
                    path = basePath + "\\txt\\R" + (i - 5) + ".txt";
                    value = TxtGetValue.eathquakeWave1(path);
                    dealCellSM(row.getCell(5), value[0]);
                    dealCellSM(row.getCell(6), value[1]);
                }
            }
        } catch (Exception e) {
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + "处理地震波信息表发生异常");
            e.printStackTrace();
        }
    }

    /**
     * 地震波持时表
     *
     * @param table8
     */
    private static void insertEarthquakeWave(XWPFTable table8) {
        System.out.println("\n处理 地震波持时表");
        try {
            //获取周期
//            String singleT = TxtGetValue.getSingleT(basePath + "\\txt\\1.txt");
//            Map<Integer, Object> map = GetExcelValue.getCycleAndFxFy(basePath + "\\excel\\材料数据.xlsx");
//            String[] cycle1 = (String[]) map.get(1);
            String[] cycle1 = data.CECLE;

            String singleT = cycle1[0];

            String path;
            String[] value;
            for (int i = 2; i < 9; i++) {
                if (i < 7) {
                    path = basePath + "\\txt\\T" + (i - 1) + ".txt";
                } else {
                    path = basePath + "\\txt\\R" + (i - 6) + ".txt";
                }
                //到对应文件中获取各个波的数据
                value = TxtGetValue.earthquakeWave(path);
                dealCellSM(table8.getRow(i).getCell(1), Util.getPrecisionString(value[0], 3));
                dealCellSM(table8.getRow(i).getCell(2), Util.getPrecisionString(value[1], 2));
                dealCellSM(table8.getRow(i).getCell(3), Util.getPrecisionString(value[2], 2));
                dealCellSM(table8.getRow(i).getCell(4), Util.getPrecisionString(Double.valueOf(value[2]) - Double.valueOf(value[1]), 2));
                dealCellSM(table8.getRow(i).getCell(5), singleT);
                dealCellSM(table8.getRow(i).getCell(6), Util.getPrecisionString((Double.valueOf(value[2]) - Double.valueOf(value[1])) / Double.valueOf(singleT), 2));
            }
        } catch (NumberFormatException e) {
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + "处理地震波持时表发生异常");
            e.printStackTrace();
        }
    }

    /**
     * 层间剪力对比
     *
     * @param table10
     */
    private static void insertFloorShearCopmare(XWPFTable table10) {
        System.out.println("=========================================================");
        System.out.println("\n处理 楼层剪力对比表");
        try {
            String[][][] shearNot = GetExcelValue.getShear(basePath + "\\excel\\工作簿3.xlsx", 3);
//            Map<Integer, Object> map = GetExcelValue.getCycleAndFxFy(basePath + "\\excel\\材料数据.xlsx");
//            List[] fxFy = (List[]) map.get(3);
            List[] fxFy = data.FX_FY;

            List<String> earthquakeAfterX = fxFy[0];
            List<String> earthquakeAfterY = fxFy[1];

            //楼层数
            int floor = Math.min(shearNot[0].length, shearNot[1].length);
            if (earthquakeAfterY.size() != floor || earthquakeAfterX.size() != floor) {
                System.out.println("楼层剪力对比表的  来自材料文件的楼层数与来自excel里的楼层数不一致");
            }
            floor = Math.min(floor, earthquakeAfterY.size());
            floor = Math.min(floor, earthquakeAfterY.size());

            XWPFTableRow row10;
            for (int i = 0; i < floor; i++) {
                //按照表头的单元格数进行添加
                table10.createRow();

                //表头有3行
                row10 = table10.getRow(i + 3);

                //表头与表身差14个单元格
                for (int j = 0; j < 14; j++) {
                    row10.addNewTableCell();
                }

                //插入值
                dealCellSM(row10.getCell(0), String.valueOf(floor - i));
                for (int j = 1; j < 8; j++) {
                    dealCellSM(row10.getCell(j), shearNot[0][floor - i - 1][j - 1]);
                    dealCellSM(row10.getCell(j + 7), shearNot[1][floor - i - 1][j - 1]);
                }

                //PKPM&YJK
                dealCellSM(row10.getCell(15), Util.getPrecisionString(earthquakeAfterX.get(i), 0));
                dealCellSM(row10.getCell(16), Util.getPrecisionString(earthquakeAfterY.get(i), 0));
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + "处理楼层剪力对比表发生异常");
        }
    }

    /**
     * 层间位移对比
     * 楼层层间位移角
     *
     * @param table12
     * @param table14
     */
    private static void insertFloorDisplaceCompare(XWPFTable table12, XWPFTable table14) {
        System.out.println("\n处理 楼层层间位移对比表与楼层层间位移角");
        try {
            //非减震结构层间位移
            String[][][] displaceNot = GetExcelValue.getDisplace(basePath + "\\excel\\工作簿3.xlsx", 2);
            //获取层高数据  此处数值单位为 豪米
//            List<String> floorHight = GetExcelValue.getFloorHigh(basePath + "\\excel\\floorH.xlsx", 0);
            Double[] floorHight = data.FLOOR_H;
            //楼层层间位移角
            //非减震结构层间位移
            String[][][] displaceNotAngle = Util.getDisplaceAngle(displaceNot, floorHight);

            //楼层数
            int floor = Math.min(floorHight.length, displaceNot[0].length);
            XWPFTableRow row12;
            XWPFTableRow row14;

            //楼层层间位移角  的和
            Double notXSum = 0d;
            Double notYSum = 0d;
            for (int i = 0; i < floor; i++) {
                notXSum = 0d;
                notYSum = 0d;

                //按照表头的单元格数进行添加
                table12.createRow();
                table14.createRow();

                //表头有3行
                row12 = table12.getRow(i + 3);
                row14 = table14.getRow(i + 3);

                //表头与表身差12个单元格
                for (int j = 0; j < 12; j++) {
                    row12.addNewTableCell();
                }
                //表头与表身差14个单元格
                for (int j = 0; j < 14; j++) {
                    row14.addNewTableCell();
                }

                //插入值
                //楼层
                dealCellSM(row12.getCell(0), String.valueOf(floor - i));
                dealCellSM(row14.getCell(0), String.valueOf(floor - i));

                //层高
                dealCellSM(row12.getCell(1), floorHight[floor - i - 1].toString());
                dealCellSM(row14.getCell(1), floorHight[floor - i - 1].toString());

                for (int j = 2; j < 9; j++) {
                    //楼层层间位移对比表
                    dealCellSM(row12.getCell(j), Util.getPrecisionString(displaceNot[0][floor - i - 1][j - 2], 2));
                    dealCellSM(row12.getCell(j + 7), Util.getPrecisionString(displaceNot[1][floor - i - 1][j - 2], 2));

                    //楼层层间位移角
                    //非减震结构层间位移 x与y
                    dealCellSM(row14.getCell(j), Util.getPrecisionString(displaceNotAngle[0][floor - i - 1][j - 2], 0));
                    dealCellSM(row14.getCell(j + 8), Util.getPrecisionString(displaceNotAngle[1][floor - i - 1][j - 2], 0));

                    //减震结构层间位移 x与y 累计
                    notXSum += Double.valueOf(displaceNotAngle[0][floor - i - 1][j - 2]);
                    notYSum += Double.valueOf(displaceNotAngle[1][floor - i - 1][j - 2]);
                }

                dealCellSM(row14.getCell(9), Util.getPrecisionString(notXSum / 7, 0));
                dealCellSM(row14.getCell(17), Util.getPrecisionString(notXSum / 7, 0));
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + "楼层层间位移对比表与楼层层间位移角");
        }
    }

    /**
     * 地震波下结构X/Y方向的弹性能
     *
     * @param table17
     * @param table18
     */
    private static void insertElasticPropertyOfBaseEarthquake(XWPFTable table17, XWPFTable table18) {
        System.out.println("\n处理 地震波下结构X/Y方向的弹性能表");
        try {
            // 减震结构层间剪力
            String[][][] shear = GetExcelValue.getShear(basePath + "\\excel\\工作簿3.xlsx", 3);
            // 减震结构层间位移
            //原来工作簿6
            String[][][] displace = GetExcelValue.getDisplace(basePath + "\\excel\\工作簿3.xlsx", 2);
            //楼层数
            int floor = Math.min(displace[0].length, displace[0].length);
            XWPFTableRow row17;
            XWPFTableRow row18;

            //表头四行，表低一行格式固定，数据加在中间
            //数据行以表格第五行数据为模版进行加入数据
            //新加入的行都插入到第六行
            //最后模板行在数据行的最下边，数据插入完成将其删除
            XWPFTableRow row170 = table17.getRow(4);
            XWPFTableRow row180 = table18.getRow(4);

            //表格最后一行的求和
            Double[] sumX = {0d, 0d, 0d, 0d, 0d, 0d, 0d};
            Double[] sumY = {0d, 0d, 0d, 0d, 0d, 0d, 0d};

            for (int i = 0; i < floor; i++) {

                table17.addRow(row170, 5);
                table18.addRow(row180, 5);

                row17 = table17.getRow(5);
                row18 = table18.getRow(5);

                //插入值     楼层
                dealCellSM(row17.getCell(0), String.valueOf(i + 1));
                dealCellSM(row18.getCell(0), String.valueOf(i + 1));

                //数据值插入
                for (int j = 0; j < 7; j++) {
                    dealCellSM(row17.getCell(j + 1), shear[0][i][j]);
                    dealCellSM(row17.getCell(j + 8), Util.getPrecisionString(displace[0][i][j], 1));
                    dealCellSM(row17.getCell(j + 15), Util.multiplyAndHalf(shear[0][i][j], displace[0][i][j]));
                    sumX[j] = sumX[j] + Double.valueOf(Util.multiplyAndHalf(shear[0][i][j], displace[0][i][j]));

                    dealCellSM(row18.getCell(j + 1), shear[1][i][j]);
                    dealCellSM(row18.getCell(j + 8), Util.getPrecisionString(displace[1][i][j], 1));
                    dealCellSM(row18.getCell(j + 15), Util.multiplyAndHalf(shear[1][i][j], displace[1][i][j]));
                    sumY[j] = sumY[j] + Double.valueOf(Util.multiplyAndHalf(shear[1][i][j], displace[1][i][j]));
                }
            }
            //移除模板行
            table17.removeRow(floor + 4);
            table18.removeRow(floor + 4);
            //插入求和行，最后一行
            for (int i = 0; i < 7; i++) {
                dealCellSM(table17.getRow(floor + 4).getCell(i + 1), Util.getPrecisionString(sumX[i].toString(), 0));
                dealCellSM(table18.getRow(floor + 4).getCell(i + 1), Util.getPrecisionString(sumY[i].toString(), 0));
            }
        } catch (NumberFormatException e) {
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + "处理 地震波下结构X/Y方向的弹性能表发生异常");
        }
    }

    /**
     * 各地震波下X/Y方向阻尼器耗能
     *
     * @param table19
     * @param table20
     */
    private static void insertEarthquakeDamperDisEnergy(XWPFTable table2, XWPFTable table19, XWPFTable table20) {
        System.out.println("\n处理 各地震波下X/Y方向阻尼器耗能表");
        try {
            //数据获取

            //X方向
            //原来是工作簿4
            Double[][][] valueX = GetExcelValue.getEarthquakeDamperDisEnergyX(basePath + "\\excel\\工作簿3.xlsx");
            //阻尼器形变
            Double[][] shapeX = valueX[0];
            //阻尼器内力
            Double[][] forceX = valueX[1];

            //Y方向
            //原来是工作簿4
            Double[][][] valueY = GetExcelValue.getEarthquakeDamperDisEnergyY(basePath + "\\excel\\工作簿3.xlsx");
            //阻尼器形变
            Double[][] shapeY = valueY[0];
            //阻尼器内力
            Double[][] forceY = valueY[1];

            XWPFTableRow row19;
            XWPFTableRow row20;

            //表头四行，
            //数据行以表格第五行数据为模版进行加入数据
            //新加入的行都插入到第六行
            //最后模板行在数据行的最下边，数据插入完成将其删除
            XWPFTableRow row190 = table19.getRow(4);
            XWPFTableRow row200 = table20.getRow(4);

            //用于计算阻尼器耗能
            double energyX;
            double energyY;
            double[] energyArrayX = {0d, 0d, 0d, 0d, 0d, 0d, 0d};
            double[] energyArrayY = {0d, 0d, 0d, 0d, 0d, 0d, 0d};

            //屈服力，屈服位移，屈服刚度
            double yieldForceX;
            double yieldDisplacementX;
            double yieldRigidityX;

            double yieldForceY;
            double yieldDisplacementY;
            double yieldRigidityY;

            int floor = Math.min(shapeX.length, forceY.length);
            floor = Math.min(floor, modelNo[0].length);
            if (floor != modelNo[0].length || floor != shapeX.length || floor != forceY.length) {
                System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$ CAD编号数量与原始表格里的数据的数量不一致 $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
            }

            for (int i = 0; i < floor; i++) {

                table19.addRow(row190, 5);
                table20.addRow(row200, 5);

                row19 = table19.getRow(5);
                row20 = table20.getRow(5);

                //插入模型编号
                dealCellSM(row19.getCell(0), modelNo[0][floor - 1 - i]);
                dealCellSM(row20.getCell(0), modelNo[1][floor - 1 - i]);
                dealCellSM(row19.getCell(1), Util.getPrecisionString(forceX[floor - i - 1][0], 0));
                dealCellSM(row20.getCell(1), Util.getPrecisionString(forceY[floor - i - 1][0], 0));

                yieldDisplacementX = Double.valueOf(row19.getCell(3).getText());
                yieldForceX = Double.valueOf(row19.getCell(2).getText());
                yieldRigidityX = Double.valueOf(row19.getCell(4).getText());

                yieldDisplacementY = Double.valueOf(row20.getCell(3).getText());
                yieldForceY = Double.valueOf(row20.getCell(2).getText());
                yieldRigidityY = Double.valueOf(row20.getCell(4).getText());

                //数据值插入
                for (int j = 1; j < 8; j++) {
                    if (shapeX[floor - i - 1][j] > yieldDisplacementX) {
                        //屈服力，屈服位移，屈服刚度
                        energyX = 4 * yieldForceX * (Util.getPrecisionDouble(shapeX[floor - i - 1][j].toString(), 2) - yieldDisplacementX) * (1 - yieldRigidityX);
                    } else {
                        energyX = 0D;
                    }
                    energyArrayX[j - 1] += energyX;

                    dealCellSM(row19.getCell(j + 4), Util.getPrecisionString(forceX[floor - i - 1][j], 0));
                    dealCellSM(row19.getCell(j + 11), Util.getPrecisionString(shapeX[floor - i - 1][j], 2));
                    dealCellSM(row19.getCell(j + 18), Util.getPrecisionString(energyX, 0));

                    if (shapeY[floor - i - 1][j] > yieldDisplacementY) {
                        //屈服力，屈服位移，屈服刚度
                        energyY = 4 * yieldForceY * (Util.getPrecisionDouble(shapeY[floor - i - 1][j].toString(), 2) - yieldDisplacementY) * (1 - yieldRigidityY);
                    } else {
                        energyY = 0D;
                    }
                    energyArrayY[j - 1] += energyY;
                    dealCellSM(row20.getCell(j + 4), Util.getPrecisionString(forceY[floor - i - 1][j], 0));
                    dealCellSM(row20.getCell(j + 11), Util.getPrecisionString(shapeY[floor - i - 1][j], 2));
                    dealCellSM(row20.getCell(j + 18), Util.getPrecisionString(energyY, 0));
                }
            }
            //移除模板行
            table19.removeRow(floor + 4);
            table20.removeRow(floor + 4);

            //插入求和的值
            row19 = table19.getRow(floor + 4);
            row20 = table20.getRow(floor + 4);
            for (int i = 0; i < 7; i++) {
                dealCellSM(row19.getCell(i + 1), Util.getPrecisionString(energyArrayX[i], 0));
                dealCellSM(row20.getCell(i + 1), Util.getPrecisionString(energyArrayY[i], 0));
            }
        } catch (NumberFormatException e) {
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + "处理 各地震波下X/Y方向阻尼器耗能表发生异常");
        }
    }

    /**
     * X/Y方向结构附加阻尼比计算
     * 该表里的数据由word里的其他表里的数据得到
     * 计算公式(v1/(v2*4*pi))
     *
     * @param table16
     * @param table17 结构附加阻尼比计算
     * @param table18
     * @param table19 各地震波下结构  的弹性能
     * @param table20
     */
    private static void insertAnnexDamperRatio(XWPFTable table16, XWPFTable table17, XWPFTable table18, XWPFTable table19, XWPFTable table20) {
        System.out.println("\n处理  X/Y方向结构附加阻尼比计算表");
        System.out.println("==================================================");
        System.out.println();
        try {
            XWPFTableRow row17 = table17.getRow(table17.getRows().size() - 1);
            XWPFTableRow row18 = table18.getRow(table18.getRows().size() - 1);
            XWPFTableRow row19 = table19.getRow(table19.getRows().size() - 1);
            XWPFTableRow row20 = table20.getRow(table20.getRows().size() - 1);
            String ratio;
            Double sumX = 0d;
            Double sumY = 0d;
            for (int i = 1; i < 8; i++) {
                dealCellSM(table16.getRow(2).getCell(i), row17.getCell(i).getText());
                dealCellSM(table16.getRow(3).getCell(i), row19.getCell(i).getText());

                //附加阻尼比
                ratio = Util.getPrecisionString(100 * Double.valueOf(row19.getCell(i).getText()) / (Double.valueOf(row17.getCell(i).getText()) * 4 * Math.PI), 2);
                dealCellSM(table16.getRow(4).getCell(i), ratio + "%");
                sumX += Double.valueOf(ratio);

                dealCellSM(table16.getRow(8).getCell(i), row18.getCell(i).getText());
                dealCellSM(table16.getRow(9).getCell(i), row20.getCell(i).getText());
                //附加阻尼比
                ratio = Util.getPrecisionString(100 * Double.valueOf(row20.getCell(i).getText()) / (Double.valueOf(row18.getCell(i).getText()) * 4 * Math.PI), 2);
                dealCellSM(table16.getRow(10).getCell(i), ratio + "%");
                sumY += Double.valueOf(ratio);
            }
            //平均值
            dealCellSM(table16.getRow(5).getCell(1), Util.getPrecisionString(sumX / 7d, 2) + "%");
            dealCellSM(table16.getRow(11).getCell(1), Util.getPrecisionString(sumY / 7d, 2) + "%");
        } catch (Exception e) {
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + "处理 X/Y方向结构附加阻尼比计算表发生异常");
        }
    }

    /**
     * 阻尼器出力与楼层剪力占比
     * <p>
     * 该表分为三部分
     * 最左边部分的数据由 "减震结构层间剪力" 获得
     * 中间部分的数据由  “阻尼器耗能“表里的阻尼器处理部分获得
     * 中间部分表格的数据获取方式特殊   在“阻尼器耗能表” 的最左边为CAD中的编号，如  X-2-1 ，Y-3-2   中间的数字表示楼层，
     * 当前处理表的中间不分的数据是由相同楼层的数据之和得到的，某些楼层可能空缺
     * 最右边的数据是由中间的数据除以右边的数据得到的，平均值为比值的平均值
     *
     * @param table21
     * @param table22
     */
    private static void insertDamperFloorRatio(XWPFTable table21, XWPFTable table22, XWPFTable table2) {
        System.out.println(" 处理 阻尼器出力与楼层剪力占比 ");
        //X方向
        //原来是工作簿7
        Double[][][] valueX = GetExcelValue.getEarthquakeDamperDisEnergyX(basePath + "\\excel\\工作簿3.xlsx");
        //阻尼器内力
        Double[][] forceX = valueX[1];

        //Y方向
        //原来是工作簿8
        Double[][][] valueY = GetExcelValue.getEarthquakeDamperDisEnergyY(basePath + "\\excel\\工作簿3.xlsx");
        //阻尼器内力
        Double[][] forceY = valueY[1];

        //阻尼器内力  第一数为模型中的编号 如force[0][0][0]，force[0][1][0]
        Double[][][] force = {forceX, forceY};

        //获取出对应楼层的求和的值
        Map<Integer, Double[]>[] maps = getDamperFloorAdd(table2, force);

        // 减震结构层间剪力
        String[][][] shear = GetExcelValue.getShear(basePath + "\\excel\\工作簿3.xlsx", 3);

        //楼层数
        int floor = shear[0].length;
        //用于某楼层 阻尼器出力之和的有无
        boolean flageX = false;
        boolean flageY = false;
        Double sumX = 0d;
        Double sumY = 0d;
        XWPFTableRow row21;
        XWPFTableRow row22;
        for (int i = 0; i < floor; i++) {
            //按照表头的单元格数进行添加
            table21.createRow();
            table22.createRow();

            //表头有4行
            row21 = table21.getRow(i + 4);
            row22 = table22.getRow(i + 4);

            //表头与表身差25个单元格
            for (int j = 0; j < 22; j++) {
                row21.addNewTableCell();
                row22.addNewTableCell();
            }

            if (maps[0].containsKey(floor - i)) {
                flageX = true;
            } else {
                flageX = false;
            }
            if (maps[1].containsKey(floor - i)) {
                flageY = true;
            } else {
                flageY = false;
            }

            //插入值
            dealCellSM(row21.getCell(0), String.valueOf(floor - i));
            dealCellSM(row22.getCell(0), String.valueOf(floor - i));
            sumX = 0d;
            sumY = 0d;
            for (int j = 1; j < 8; j++) {
                //减震结构层间剪力 x与y
                dealCellSM(row21.getCell(j), shear[0][floor - i - 1][j - 1]);
                dealCellSM(row22.getCell(j), shear[1][floor - i - 1][j - 1]);

                //楼层阻尼器出力之和
                if (flageX) {
                    dealCellSM(row21.getCell(j + 7), Util.getPrecisionString(maps[0].get(floor - i)[j - 1], 0));
                    dealCellSM(row21.getCell(j + 14), Util.getPrecisionString(100 * maps[0].get(floor - i)[j - 1] / Double.valueOf(shear[0][floor - i - 1][j - 1]), 2));
                    sumX += 100 * maps[0].get(floor - i)[j - 1] / Double.valueOf(shear[0][floor - i - 1][j - 1]);
                } else {
                    dealCellSM(row21.getCell(j + 7), "\\");
                    dealCellSM(row21.getCell(j + 14), "\\");
                }
                if (flageY) {
                    dealCellSM(row22.getCell(j + 7), Util.getPrecisionString(maps[1].get(floor - i)[j - 1], 0));
                    dealCellSM(row22.getCell(j + 14), Util.getPrecisionString(100 * maps[1].get(floor - i)[j - 1] / Double.valueOf(shear[1][floor - i - 1][j - 1]), 2));
                    sumY += 100 * maps[1].get(floor - i)[j - 1] / Double.valueOf(shear[1][floor - i - 1][j - 1]);
                } else {
                    dealCellSM(row22.getCell(j + 7), "\\");
                    dealCellSM(row22.getCell(j + 14), "\\");
                }
            }
            if (flageX) {
                dealCellSM(row21.getCell(22), Util.getPrecisionString(sumX / 7, 2));
            } else {
                dealCellSM(row21.getCell(22), "\\");
            }

            if (flageY) {
                dealCellSM(row22.getCell(22), Util.getPrecisionString(sumY / 7, 2));
            } else {
                dealCellSM(row22.getCell(22), "\\");
            }

        }
    }

    /**
     * 结构层间位移角
     * 大震下非减震和减震的结构层间位移角
     *
     * @param table23
     * @param table24
     */
    private static void insertFloorDisplaceAngle(XWPFTable table23, XWPFTable table24) {
        System.out.println("\n处理  大震下非减震和减震的结构层间位移角表");
        System.out.println("==================================================");
        try {
            //非减震结构层间位移
            String[][][] displaceAngleNot = GetExcelValue.getDisplaceAngle(basePath + "\\excel\\工作簿4.xlsx", 0);
            // 减震结构层间位移
            String[][][] displaceAngle = GetExcelValue.getDisplaceAngle(basePath + "\\excel\\工作簿5.xlsx", 2);
            Double[] floorh = data.FLOOR_H;
            //获取有效列
            Integer[] valueCol = Util.getValueCol(displaceAngleNot[0]);
            if (valueCol == null) {
                System.out.println("有效的列无法确定");
            }

            //更改表头的有效列的名称
            XWPFTableRow row23 = table23.getRow(2);
            XWPFTableRow row24 = table24.getRow(2);
//            String name = null;
//            for (int i = 0; i < 7; i++) {
//                name = getName1(valueCol, i);
//                dealCellSM(row23.getCell(1 + i), name);
//                dealCellSM(row23.getCell(9 + i), name);
//                dealCellSM(row24.getCell(1 + i), name);
//                dealCellSM(row24.getCell(9 + i), name);
//            }


            int floor = Math.min(displaceAngleNot[0].length, displaceAngle[0].length);
            floor = Math.min(floor, floorh.length);

            //包络值
            Double envelopeX = null;
            Double envelopeXNot = null;
            Double envelopeY = null;
            Double envelopeYNot = null;

            //对于包络值得列取最小值
            Double minEnvelopeX = null;
            Double minEnvelopeXNot = null;
            Double minEnvelopeY = null;
            Double minEnvelopeYNot = null;

            //表头三行，表低两行格式固定，数据加在中间
            //数据行以表格第四行数据为模版进行加入数据
            //新加入的行都插入到第五行
            //最后模板行在数据行的最下边，将其删除
            XWPFTableRow row210 = table23.getRow(3);
            XWPFTableRow row220 = table24.getRow(3);
            for (int i = 0; i < floor; i++) {

                table23.addRow(row210, 4);
                table24.addRow(row220, 4);

                row23 = table23.getRow(4);
                row24 = table24.getRow(4);

                //插入值
                dealCellSM(row23.getCell(0), String.valueOf(i + 1));
                dealCellSM(row24.getCell(0), String.valueOf(i + 1));

                //数据值插入
                for (int j = 0; j < 7; j++) {
                    //非减震结构层间位移 x与y
                    double ddd = Double.valueOf(displaceAngleNot[0][i][valueCol[j]]);
                    double sss = floorh[i] / Double.valueOf(displaceAngleNot[0][i][valueCol[j]]);
                    dealCellSM(row23.getCell(j + 1), Util.getPrecisionString(floorh[i] / Double.valueOf(displaceAngleNot[0][i][valueCol[j]]), 0));
                    dealCellSM(row24.getCell(j + 1), Util.getPrecisionString(floorh[i] / Double.valueOf(displaceAngleNot[1][i][valueCol[j]]), 0));
                    //减震结构层间位移 x与y
                    dealCellSM(row23.getCell(j + 9), Util.getPrecisionString(floorh[i] / Double.valueOf(displaceAngle[0][i][valueCol[j]]), 0));
                    dealCellSM(row24.getCell(j + 9), Util.getPrecisionString(floorh[i] / Double.valueOf(displaceAngle[1][i][valueCol[j]]), 0));
                }

                //计算包络值
                //包络值为该行的  这7个数值的平均值

                envelopeX = floorh[i] / Util.getAvg(
                        Double.valueOf(displaceAngle[0][i][valueCol[0]]),
                        Double.valueOf(displaceAngle[0][i][valueCol[1]]),
                        Double.valueOf(displaceAngle[0][i][valueCol[2]]),
                        Double.valueOf(displaceAngle[0][i][valueCol[3]]),
                        Double.valueOf(displaceAngle[0][i][valueCol[4]]),
                        Double.valueOf(displaceAngle[0][i][valueCol[5]]),
                        Double.valueOf(displaceAngle[0][i][valueCol[6]]));

                envelopeXNot = floorh[i] / Util.getAvg(
                        Double.valueOf(displaceAngleNot[0][i][valueCol[0]]),
                        Double.valueOf(displaceAngleNot[0][i][valueCol[1]]),
                        Double.valueOf(displaceAngleNot[0][i][valueCol[2]]),
                        Double.valueOf(displaceAngleNot[0][i][valueCol[3]]),
                        Double.valueOf(displaceAngleNot[0][i][valueCol[4]]),
                        Double.valueOf(displaceAngleNot[0][i][valueCol[5]]),
                        Double.valueOf(displaceAngleNot[0][i][valueCol[6]]));

                envelopeY = floorh[i] / Util.getAvg(
                        Double.valueOf(displaceAngle[1][i][valueCol[0]]),
                        Double.valueOf(displaceAngle[1][i][valueCol[1]]),
                        Double.valueOf(displaceAngle[1][i][valueCol[2]]),
                        Double.valueOf(displaceAngle[1][i][valueCol[3]]),
                        Double.valueOf(displaceAngle[1][i][valueCol[4]]),
                        Double.valueOf(displaceAngle[1][i][valueCol[5]]),
                        Double.valueOf(displaceAngle[1][i][valueCol[6]]));

                envelopeYNot = floorh[i] / Util.getAvg(
                        Double.valueOf(displaceAngleNot[1][i][valueCol[0]]),
                        Double.valueOf(displaceAngleNot[1][i][valueCol[1]]),
                        Double.valueOf(displaceAngleNot[1][i][valueCol[2]]),
                        Double.valueOf(displaceAngleNot[1][i][valueCol[3]]),
                        Double.valueOf(displaceAngleNot[1][i][valueCol[4]]),
                        Double.valueOf(displaceAngleNot[1][i][valueCol[5]]),
                        Double.valueOf(displaceAngleNot[1][i][valueCol[6]]));

//                envelopeX = floorh[i] / Math.max(Double.valueOf(displaceAngle[0][i][valueCol[0]]),
//                        Math.max(Double.valueOf(displaceAngle[0][i][valueCol[1]]),
//                                Math.max(Double.valueOf(displaceAngle[0][i][valueCol[2]]),
//                                        Math.max(Double.valueOf(displaceAngle[0][i][valueCol[3]]),
//                                                Math.max(Double.valueOf(displaceAngle[0][i][valueCol[4]]),
//                                                        Math.max(Double.valueOf(displaceAngle[0][i][valueCol[5]]),
//                                                                Double.valueOf(displaceAngle[0][i][valueCol[6]])))))));
//                envelopeXNot = floorh[i] / Math.max(Double.valueOf(displaceAngleNot[0][i][valueCol[0]]),
//                        Math.max(Double.valueOf(displaceAngleNot[0][i][valueCol[1]]),
//                                Math.max(Double.valueOf(displaceAngleNot[0][i][valueCol[2]]),
//                                        Math.max(Double.valueOf(displaceAngleNot[0][i][valueCol[3]]),
//                                                Math.max(Double.valueOf(displaceAngleNot[0][i][valueCol[4]]),
//                                                        Math.max(Double.valueOf(displaceAngleNot[0][i][valueCol[5]]),
//                                                                Double.valueOf(displaceAngleNot[0][i][valueCol[6]])))))));
//                envelopeY = floorh[i] / Math.max(Double.valueOf(displaceAngle[1][i][valueCol[0]]),
//                        Math.max(Double.valueOf(displaceAngle[1][i][valueCol[1]]),
//                                Math.max(Double.valueOf(displaceAngle[1][i][valueCol[2]]),
//                                        Math.max(Double.valueOf(displaceAngle[1][i][valueCol[3]]),
//                                                Math.max(Double.valueOf(displaceAngle[1][i][valueCol[4]]),
//                                                        Math.max(Double.valueOf(displaceAngle[1][i][valueCol[5]]),
//                                                                Double.valueOf(displaceAngle[1][i][valueCol[6]])))))));
//                envelopeYNot = floorh[i] / Math.max(Double.valueOf(displaceAngleNot[1][i][valueCol[0]]),
//                        Math.max(Double.valueOf(displaceAngleNot[1][i][valueCol[1]]),
//                                Math.max(Double.valueOf(displaceAngleNot[1][i][valueCol[2]]),
//                                        Math.max(Double.valueOf(displaceAngleNot[1][i][valueCol[3]]),
//                                                Math.max(Double.valueOf(displaceAngleNot[1][i][valueCol[4]]),
//                                                        Math.max(Double.valueOf(displaceAngleNot[1][i][valueCol[5]]),
//                                                                Double.valueOf(displaceAngleNot[1][i][valueCol[6]])))))));

                //获取包络值列的最小值
                minEnvelopeX = minEnvelopeX == null ? envelopeX : Math.min(minEnvelopeX, envelopeX);
                minEnvelopeXNot = minEnvelopeXNot == null ? envelopeXNot : Math.min(minEnvelopeXNot, envelopeXNot);
                minEnvelopeY = minEnvelopeY == null ? envelopeY : Math.min(minEnvelopeY, envelopeY);
                minEnvelopeYNot = minEnvelopeYNot == null ? envelopeYNot : Math.min(minEnvelopeYNot, envelopeYNot);

                //插入包络值
                dealCellSM(row23.getCell(8), Util.getPrecisionString(envelopeXNot, 0));
                dealCellSM(row23.getCell(16), Util.getPrecisionString(envelopeX, 0));
                dealCellSM(row24.getCell(8), Util.getPrecisionString(envelopeYNot, 0));
                dealCellSM(row24.getCell(16), Util.getPrecisionString(envelopeY, 0));
            }
            table23.removeRow(floor + 3);
            table24.removeRow(floor + 3);

            // 计算位移比
            Double proX = minEnvelopeXNot / minEnvelopeX;
            Double proY = minEnvelopeYNot / minEnvelopeY;

            //插入最小包络值和位移比例
            dealCellSM(table23.getRow(3 + floor).getCell(1), Util.getPrecisionString(minEnvelopeXNot, 0));
            dealCellSM(table23.getRow(3 + floor).getCell(2), Util.getPrecisionString(minEnvelopeX, 0));
            dealCellSM(table23.getRow(3 + floor + 1).getCell(1), Util.getPrecisionString(proX.toString(), 2));
            dealCellSM(table24.getRow(3 + floor).getCell(1), Util.getPrecisionString(minEnvelopeYNot, 0));
            dealCellSM(table24.getRow(3 + floor).getCell(2), Util.getPrecisionString(minEnvelopeY, 0));
            dealCellSM(table24.getRow(3 + floor + 1).getCell(1), Util.getPrecisionString(proY.toString(), 2));
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + "处理 大震下非减震和减震的结构层间位移角表发生异常");
        }
    }

    /**
     * 结构各层阻尼器最大出力及位移包络值汇总
     * 粘滞阻尼器性能规格表
     *
     * @param table25
     * @param table26
     * @param table1  table1 的部分数据来源于table24   和   table25
     *                单独写方法时 table24和table25获取到的各行的对象都一样，无法获取数据，所以直接就写在一个方法里
     */
    private static void maxEarthquakeDapmerForceDisplace(XWPFTable table25, XWPFTable table26, XWPFTable table1) {
        System.out.println("\n处理  结构各层阻尼器最大出力及位移包络值汇总表");
        try {

            //X方向  //原来工作簿11
            Double[][][] valueX = GetExcelValue.getEarthquakeDamperDisEnergyX(basePath + "\\excel\\工作簿5.xlsx");
            //阻尼器形变
            Double[][] shapeX = valueX[0];
            //阻尼器内力
            Double[][] forceX = valueX[1];

            //Y方向
            //原来是工作簿12
            Double[][][] valueY = GetExcelValue.getEarthquakeDamperDisEnergyY(basePath + "\\excel\\工作簿5.xlsx");
            //阻尼器形变
            Double[][] shapeY = valueY[0];
            //阻尼器内力
            Double[][] forceY = valueY[1];

            //获取有效列
            Integer[] valueCol = Util.getValueCol(shapeX);
            if (valueCol == null) {
                System.out.println("有效的列无法确定");
            }

            //更改表头的有效列的名称
            XWPFTableRow row25 = table25.getRow(3);
            XWPFTableRow row26 = table26.getRow(3);
//            String name = null;
//            for (int i = 0; i < 3; i++) {
//                name = getName(valueCol, i);
//                dealCellSM(row25.getCell(4 + i), name);
//                dealCellSM(row25.getCell(7 + i), name);
//                dealCellSM(row26.getCell(4 + i), name);
//                dealCellSM(row26.getCell(7 + i), name);
//            }


            //表头四行，
            //数据行以表格第五行数据为模版进行加入数据
            //新加入的行都插入到第六行
            //最后模板行在数据行的最下边，数据插入完成将其删除
            XWPFTableRow row250 = table25.getRow(4);
            XWPFTableRow row260 = table26.getRow(4);

            //包络值     内力/形变/速度
            double forceEnvelope;
            double shapeEnvelope;
            double speedEnvelope;

            //极限值     内力/形变/速度
            double forceLimit;
            double shapeLimit;
            double speedLimit;

            //各属性的较大值
            Double[] propertyMax = {0d, 0d, 0d, 0d, 0d, 0d};

            int floor = Math.min(shapeX.length, forceX.length);
            floor = Math.min(floor, modelNo[0].length);
            floor = Math.min(floor, modelNo[1].length);
            if (floor != shapeX.length || floor != forceY.length || floor != modelNo[0].length || floor != modelNo[1].length) {
                System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$ CAD编号数量与原始表格里的数据的数量不一致 $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
            }
            for (int i = 0; i < floor; i++) {

                table25.addRow(row250, 5);
                table26.addRow(row260, 5);

                row25 = table25.getRow(5);
                row26 = table26.getRow(5);

                //插入CAD编号
                dealCellSM(row25.getCell(0), modelNo[0][floor - 1 - i]);
                dealCellSM(row26.getCell(0), modelNo[1][floor - 1 - i]);
                //插入模型编号
                dealCellSM(row25.getCell(1), Util.getPrecisionString(forceX[floor - i - 1][0], 0));
                dealCellSM(row26.getCell(1), Util.getPrecisionString(shapeY[floor - i - 1][0], 0));

                //x方向
                //数据值插入
                dealCellSM(row25.getCell(4), Util.getPrecisionString(forceX[floor - i - 1][valueCol[0]], 0));
                dealCellSM(row25.getCell(5), Util.getPrecisionString(forceX[floor - i - 1][valueCol[1]], 0));
                dealCellSM(row25.getCell(6), Util.getPrecisionString(forceX[floor - i - 1][valueCol[2]], 0));
                dealCellSM(row25.getCell(7), Util.getPrecisionString(forceX[floor - i - 1][valueCol[3]], 0));
                dealCellSM(row25.getCell(8), Util.getPrecisionString(forceX[floor - i - 1][valueCol[4]], 0));
                dealCellSM(row25.getCell(9), Util.getPrecisionString(forceX[floor - i - 1][valueCol[5]], 0));
                dealCellSM(row25.getCell(10), Util.getPrecisionString(forceX[floor - i - 1][valueCol[6]], 0));

                dealCellSM(row25.getCell(11), Util.getPrecisionString(shapeX[floor - i - 1][valueCol[0]], 2));
                dealCellSM(row25.getCell(12), Util.getPrecisionString(shapeX[floor - i - 1][valueCol[1]], 2));
                dealCellSM(row25.getCell(13), Util.getPrecisionString(shapeX[floor - i - 1][valueCol[2]], 2));
                dealCellSM(row25.getCell(14), Util.getPrecisionString(shapeX[floor - i - 1][valueCol[3]], 2));
                dealCellSM(row25.getCell(15), Util.getPrecisionString(shapeX[floor - i - 1][valueCol[4]], 2));
                dealCellSM(row25.getCell(16), Util.getPrecisionString(shapeX[floor - i - 1][valueCol[5]], 2));
                dealCellSM(row25.getCell(17), Util.getPrecisionString(shapeX[floor - i - 1][valueCol[6]], 2));


                //y方向
                dealCellSM(row26.getCell(4), Util.getPrecisionString(forceY[floor - i - 1][valueCol[0]], 0));
                dealCellSM(row26.getCell(5), Util.getPrecisionString(forceY[floor - i - 1][valueCol[1]], 0));
                dealCellSM(row26.getCell(6), Util.getPrecisionString(forceY[floor - i - 1][valueCol[2]], 0));
                dealCellSM(row26.getCell(7), Util.getPrecisionString(forceY[floor - i - 1][valueCol[3]], 0));
                dealCellSM(row26.getCell(8), Util.getPrecisionString(forceY[floor - i - 1][valueCol[4]], 0));
                dealCellSM(row26.getCell(9), Util.getPrecisionString(forceY[floor - i - 1][valueCol[5]], 0));
                dealCellSM(row26.getCell(10), Util.getPrecisionString(forceY[floor - i - 1][valueCol[6]], 0));

                dealCellSM(row26.getCell(11), Util.getPrecisionString(shapeY[floor - i - 1][valueCol[0]], 2));
                dealCellSM(row26.getCell(12), Util.getPrecisionString(shapeY[floor - i - 1][valueCol[1]], 2));
                dealCellSM(row26.getCell(13), Util.getPrecisionString(shapeY[floor - i - 1][valueCol[2]], 2));
                dealCellSM(row26.getCell(14), Util.getPrecisionString(shapeY[floor - i - 1][valueCol[3]], 2));
                dealCellSM(row26.getCell(15), Util.getPrecisionString(shapeY[floor - i - 1][valueCol[4]], 2));
                dealCellSM(row26.getCell(16), Util.getPrecisionString(shapeY[floor - i - 1][valueCol[5]], 2));
                dealCellSM(row26.getCell(17), Util.getPrecisionString(shapeY[floor - i - 1][valueCol[6]], 2));

//                dealCellSM(row26.getCell(4), Util.getPrecisionString(forceY[floor - i - 1][valueCol[0]], 0));
//                dealCellSM(row26.getCell(5), Util.getPrecisionString(forceY[floor - i - 1][valueCol[1]], 0));
//                dealCellSM(row26.getCell(6), Util.getPrecisionString(forceY[floor - i - 1][valueCol[2]], 0));
//                dealCellSM(row26.getCell(7), Util.getPrecisionString(shapeY[floor - i - 1][valueCol[0]], 2));
//                dealCellSM(row26.getCell(8), Util.getPrecisionString(shapeY[floor - i - 1][valueCol[1]], 2));
//                dealCellSM(row26.getCell(9), Util.getPrecisionString(shapeY[floor - i - 1][valueCol[0]], 2));

                //x方向
                //包络值
                forceEnvelope = Util.getAvg(
                        forceX[floor - i - 1][valueCol[0]],
                        forceX[floor - i - 1][valueCol[1]],
                        forceX[floor - i - 1][valueCol[2]],
                        forceX[floor - i - 1][valueCol[3]],
                        forceX[floor - i - 1][valueCol[4]],
                        forceX[floor - i - 1][valueCol[5]],
                        forceX[floor - i - 1][valueCol[6]]);

                shapeEnvelope = Util.getAvg(
                        shapeX[floor - i - 1][valueCol[0]],
                        shapeX[floor - i - 1][valueCol[1]],
                        shapeX[floor - i - 1][valueCol[2]],
                        shapeX[floor - i - 1][valueCol[3]],
                        shapeX[floor - i - 1][valueCol[4]],
                        shapeX[floor - i - 1][valueCol[5]],
                        shapeX[floor - i - 1][valueCol[6]]);

//                forceEnvelope = Math.max(forceX[floor - i - 1][valueCol[0]],
//                        Math.max(forceX[floor - i - 1][valueCol[1]],
//                                Math.max(forceX[floor - i - 1][valueCol[2]],
//                                        Math.max(forceX[floor - i - 1][valueCol[3]],
//                                                Math.max(forceX[floor - i - 1][valueCol[4]],
//                                                        Math.max(forceX[floor - i - 1][valueCol[5]],
//                                                                forceX[floor - i - 1][valueCol[6]]))))));
//
//                shapeEnvelope = Math.max(shapeX[floor - i - 1][valueCol[0]],
//                        Math.max(shapeX[floor - i - 1][valueCol[1]],
//                                Math.max(shapeX[floor - i - 1][valueCol[2]],
//                                        Math.max(shapeX[floor - i - 1][valueCol[3]],
//                                                Math.max(shapeX[floor - i - 1][valueCol[4]],
//                                                        Math.max(shapeX[floor - i - 1][valueCol[5]],
//                                                                shapeX[floor - i - 1][valueCol[6]]))))));
                speedEnvelope = Math.pow(forceEnvelope / Double.valueOf(row25.getCell(2).getText()), 1d / Double.valueOf(row25.getCell(3).getText()));
                dealCellSM(row25.getCell(18), Util.getPrecisionString(forceEnvelope, 0));
                dealCellSM(row25.getCell(19), Util.getPrecisionString(shapeEnvelope, 2));
//                dealCellSM(row25.getCell(12), Util.getPrecisionString(speedEnvelope, 0));
                //极限值
                speedLimit = speedEnvelope * 1.2d;
                forceLimit = Math.pow(speedLimit, Double.valueOf(row25.getCell(3).getText())) * Double.valueOf(row25.getCell(2).getText());
                shapeLimit = shapeEnvelope * 1.2d;
//                dealCellSM(row25.getCell(15), Util.getPrecisionString(speedLimit, 0));
                dealCellSM(row25.getCell(20), Util.getPrecisionString(forceLimit, 0));
                dealCellSM(row25.getCell(21), Util.getPrecisionString(shapeLimit, 1));
                //较大值比较选择
                propertyMax[0] = Math.max(propertyMax[0], forceEnvelope);
                propertyMax[1] = Math.max(propertyMax[1], shapeEnvelope);
                propertyMax[2] = Math.max(propertyMax[2], speedEnvelope);
                propertyMax[3] = Math.max(propertyMax[3], forceLimit);
                propertyMax[4] = Math.max(propertyMax[4], shapeLimit);
                propertyMax[5] = Math.max(propertyMax[5], speedLimit);


                //y方向
                //包络值
                forceEnvelope = Util.getAvg(
                        forceY[floor - i - 1][valueCol[0]],
                        forceY[floor - i - 1][valueCol[1]],
                        forceY[floor - i - 1][valueCol[2]],
                        forceY[floor - i - 1][valueCol[3]],
                        forceY[floor - i - 1][valueCol[4]],
                        forceY[floor - i - 1][valueCol[5]],
                        forceY[floor - i - 1][valueCol[6]]);

                shapeEnvelope = Util.getAvg(
                        shapeY[floor - i - 1][valueCol[0]],
                        shapeY[floor - i - 1][valueCol[1]],
                        shapeY[floor - i - 1][valueCol[2]],
                        shapeY[floor - i - 1][valueCol[3]],
                        shapeY[floor - i - 1][valueCol[4]],
                        shapeY[floor - i - 1][valueCol[5]],
                        shapeY[floor - i - 1][valueCol[6]]);

//                forceEnvelope = Math.max(forceY[floor - i - 1][valueCol[0]],
//                        Math.max(forceY[floor - i - 1][valueCol[1]],
//                                Math.max(forceY[floor - i - 1][valueCol[2]],
//                                        Math.max(forceY[floor - i - 1][valueCol[3]],
//                                                Math.max(forceY[floor - i - 1][valueCol[4]],
//                                                        Math.max(forceY[floor - i - 1][valueCol[5]],
//                                                                forceY[floor - i - 1][valueCol[6]]))))));
//
//                shapeEnvelope = Math.max(shapeY[floor - i - 1][valueCol[0]],
//                        Math.max(shapeY[floor - i - 1][valueCol[1]],
//                                Math.max(shapeY[floor - i - 1][valueCol[2]],
//                                        Math.max(shapeY[floor - i - 1][valueCol[3]],
//                                                Math.max(shapeY[floor - i - 1][valueCol[4]],
//                                                        Math.max(shapeY[floor - i - 1][valueCol[5]],
//                                                        shapeY[floor - i - 1][valueCol[6]]))))));

                speedEnvelope = Math.pow(forceEnvelope / Double.valueOf(row26.getCell(2).getText()), 1d / Double.valueOf(row26.getCell(3).getText()));
                dealCellSM(row26.getCell(18), Util.getPrecisionString(forceEnvelope, 0));
                dealCellSM(row26.getCell(19), Util.getPrecisionString(shapeEnvelope, 2));
//                dealCellSM(row26.getCell(12), Util.getPrecisionString(speedEnvelope, 0));
                //极限值
                speedLimit = speedEnvelope * 1.2d;
                forceLimit = Math.pow(speedLimit, Double.valueOf(row26.getCell(3).getText())) * Double.valueOf(row26.getCell(2).getText());
                shapeLimit = shapeEnvelope * 1.2d;
//                dealCellSM(row26.getCell(15), Util.getPrecisionString(speedEnvelope, 0));
                dealCellSM(row26.getCell(20), Util.getPrecisionString(forceLimit, 0));
                dealCellSM(row26.getCell(21), Util.getPrecisionString(shapeLimit, 1));

                //较大值比较选择
                propertyMax[0] = Math.max(propertyMax[0], forceEnvelope);
                propertyMax[1] = Math.max(propertyMax[1], shapeEnvelope);
                propertyMax[2] = Math.max(propertyMax[2], speedEnvelope);
                propertyMax[3] = Math.max(propertyMax[3], forceLimit);
                propertyMax[4] = Math.max(propertyMax[4], shapeLimit);
                propertyMax[5] = Math.max(propertyMax[5], speedLimit);
            }
            //移除模板行
            table25.removeRow(floor + 4);
            table26.removeRow(floor + 4);

            //处理table1
            dealCellSM(table1.getRow(1).getCell(3), Util.getPrecisionString(propertyMax[0], 0));
            dealCellSM(table1.getRow(1).getCell(4), Util.getPrecisionString(propertyMax[1], 0));
            dealCellSM(table1.getRow(1).getCell(5), Util.getPrecisionString(propertyMax[3], 0));
            dealCellSM(table1.getRow(1).getCell(6), Util.getPrecisionString(propertyMax[4], 0));

            dealCellSM(table1.getRow(1).getCell(8), (floor * 2) + "");
            dealCellSM(table1.getRow(2).getCell(1), (floor * 2) + "");
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" + "处理 结构各层阻尼器最大出力及位移包络值汇总表发生异常");
        }
    }


    /**
     * 计算最后几个表里的值
     * 减震器周边子结构的设计计算方法
     *
     * @param table29
     * @param table30
     * @param table31
     */
    private static void calculateTable(XWPFTable table29, XWPFTable table30, XWPFTable table31) {
        System.out.println("======================================= 计算最后三个表的数据 =======================================================");
        //计算参数所在的位置
        String paramsPath = basePath + "\\excel\\材料数据.xlsx";

        //1.excle里获取计算参数
        Map<String, Object>[] caculateParams = GetExcelValue.getCaculateParams(paramsPath);

        //2.子结构框架梁 受弯受剪 验算
        System.out.println("===============================子结构框架梁 受弯受剪 验算 =======================================");
        CaculateTable.caculateTable1(table29, caculateParams[0]);

        //3.子结构框架柱抗剪验算
        System.out.println("=============================== 子结构框架柱抗剪验算 =======================================");
        CaculateTable.caculateTable2(table30, caculateParams[1]);

        //4.悬臂墙配筋验算
        System.out.println("=============================== 悬臂墙配筋验算 =======================================");
        CaculateTable.caculateTable3(table31, caculateParams[2]);
    }

    /**
     * 获取CAD中的编号
     *
     * @param table2
     * @return
     */
    private static String[][] getModelNo(XWPFTable table2) {
        System.out.println();
        System.out.println("========================================================");
        System.out.println("获取模型中的编号");
        List<XWPFTableRow> rows = table2.getRows();
        String[][] returnValue = new String[2][rows.size() - 1];
        for (int i = 1; i < rows.size(); i++) {
            returnValue[0][i - 1] = rows.get(i).getCell(0).getText();
            returnValue[1][i - 1] = rows.get(i).getCell(2).getText();
        }
        for (int i = 0; i < returnValue[0].length; i++) {
            System.out.println(returnValue[0][i]);
        }
        System.out.println();
        for (int i = 0; i < returnValue[1].length; i++) {
            System.out.println(returnValue[1][i]);
        }
        return returnValue;
    }


    private static void init() {
        System.out.println("==========初始化   通过excel获取模型中的编号，层高，累计层高 ===============");
        String path = basePath + "\\excel\\材料数据.xlsx";
        data = GetExcelValue.init(path);
    }

    /**
     * 根据CAD中的编号顺序确定每层所对应的编号位置
     *
     * @param arrayModelNo
     * @return
     */
    private static Map<Integer, List<Integer>> getFloorOnPositionOfModelNO(String[][] arrayModelNo) {
        Map<Integer, List<Integer>> map = new HashMap<>();
        if (arrayModelNo == null || arrayModelNo.length == 0) {
            return map;
        }
        Integer floor;
        for (int i = 0; i < arrayModelNo[0].length; i++) {
            floor = Integer.valueOf(arrayModelNo[0][i].substring(2, 3));
            if (map.containsKey(floor)) {
                map.get(floor).add(i);
            } else {
                List<Integer> position = new ArrayList<>();
                position.add(i);
                map.put(floor, position);
            }
        }
        return map;
    }


    /**
     * 按楼层获取相应位置的数的平均数
     *
     * @return
     */
    private static Map<Integer, Double> getAvgValueGroupByFloorFromTable(Map<Integer, List<Integer>> param, Double[][] valueArray) {
        Map<Integer, Double> map = new HashMap<>();
        Double valueSum = 0d;
        List<Integer> positionList;
        int z = 0;
        for (int i = 1; i <= param.size(); i++) {
            positionList = param.get(i);
            for (Integer posi : positionList) {
                valueSum = 0d;
                for (int k = 1; k < 8; k++) {
                    valueSum += valueArray[posi][k];
                }
                map.put(z++, valueSum / 7.00);
            }
        }
        return map;
    }

    /**
     * 获取CAD中的编号和模型中的编号
     *
     * @param table2
     * @return
     */
    private static String[][][] getModelNo1(XWPFTable table2) {
        System.out.println();
        System.out.println("===============================================");
        System.out.println("获取模型中的编号和CAD中的编号");
        List<XWPFTableRow> rows = table2.getRows();
        String[][][] returnValue = new String[2][rows.size() - 1][2];
        for (int i = 1; i < rows.size(); i++) {
            returnValue[0][i - 1][0] = rows.get(i).getCell(0).getText();
            returnValue[0][i - 1][1] = rows.get(i).getCell(1).getText();

            returnValue[1][i - 1][0] = rows.get(i).getCell(2).getText();
            returnValue[1][i - 1][1] = rows.get(i).getCell(3).getText();
        }
        for (int i = 0; i < returnValue[0].length; i++) {
            System.out.println(returnValue[0][i][0] + " : " + returnValue[0][i][1]);
        }
        System.out.println();
        System.out.println();
        for (int i = 0; i < returnValue[1].length; i++) {
            System.out.println(returnValue[1][i][0] + " : " + returnValue[1][i][1]);
        }
        return returnValue;
    }


    private static Map<Integer, Double[]>[] getDamperFloorAdd(XWPFTable table2, Double[][][] force) {
        String[][][] modelNo = getModelNo1(table2);

        Map<Integer, Double[]>[] returnValue = new Map[2];
        Integer floor;
        String no;
        boolean flag = false;
        int count = 0;
        Double[] data;
        double f;
        for (int h = 0; h < 2; h++) {
            Map<Integer, Double[]> map = new HashMap<>();
            for (int i = 0; i < modelNo[h].length; i++) {
                floor = Integer.valueOf(modelNo[h][i][0].substring(2, 3));
                no = modelNo[h][i][1];
                flag = false;
                for (int j = 0; j < force[h].length; j++) {
                    f = force[h][j][0];
                    if (f == Double.valueOf(no)) {
                        flag = true;
                        count = j;
                        break;
                    }
                }
                if (flag) {
                    if (!map.containsKey(floor)) {
                        data = new Double[7];
                        for (int j = 0; j < 7; j++) {
                            data[j] = force[h][count][j + 1];
                        }
                        map.put(floor, data);
                    } else {
                        data = map.get(floor);
                        for (int j = 0; j < 7; j++) {
                            data[j] = data[j] + force[h][count][j + 1];
                        }
                    }
                }
            }
            map.forEach((k, v) -> {
                System.out.println(k);
                for (int s = 0; s < v.length; s++) {
                    System.out.print(v[s] + ",");
                }
                System.out.println();
            });
            returnValue[h] = map;
        }
        return returnValue;
    }

    private static void dealCellBig(XWPFTableCell cell, String text) {
//        dealCell(cell, text, 14);
        dealCell(cell, text, 10);
    }

    private static void dealCellSM(XWPFTableCell cell, String text) {
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
        cell.setVerticalAlignment(XWPFVertAlign.CENTER);
        pr.setAlignment(ParagraphAlignment.CENTER);
    }

    private static String getName(Integer[] valueCol, int i) {
        String name = null;
        if (valueCol[i] == 1) {
            name = "T1";
        } else if (valueCol[i] == 2) {
            name = "T2";
        } else if (valueCol[i] == 3) {
            name = "T3";
        } else if (valueCol[i] == 4) {
            name = "T4";
        } else if (valueCol[i] == 5) {
            name = "T5";
        } else if (valueCol[i] == 6) {
            name = "R1";
        } else if (valueCol[i] == 7) {
            name = "R2";
        } else {
            System.out.println(" 有效的列无法确定");
        }
        return name;
    }

    private static String getName1(Integer[] valueCol, int i) {
        String name = null;
        if (valueCol[i] == 0) {
            name = "T1";
        } else if (valueCol[i] == 1) {
            name = "T2";
        } else if (valueCol[i] == 2) {
            name = "T3";
        } else if (valueCol[i] == 3) {
            name = "T4";
        } else if (valueCol[i] == 4) {
            name = "T5";
        } else if (valueCol[i] == 5) {
            name = "R1";
        } else if (valueCol[i] == 6) {
            name = "R2";
        } else {
            System.out.println(" 有效的列无法确定");
        }
        return name;
    }


//    public static void main(String args[]) throws IOException {
//        String wordPath = "C:\\Users\\lizhongxiang\\Desktop\\workSpace\\最终\\cccc.docx";
//        FileInputStream in = null;
//        XWPFDocument word = null;
//        FileOutputStream out = null;
//        in = new FileInputStream(wordPath);
//        word = new XWPFDocument(in);
//        out = new FileOutputStream("C:\\Users\\lizhongxiang\\Desktop\\workSpace\\最终\\out"+System.currentTimeMillis()+" .docx");
//        List<XWPFTable> tables = word.getTables();
//        calculateTable(tables.get(29), tables.get(30), tables.get(31));
//        word.write(out);
//        out.flush();
//        out.close();
//        in.close();
//    }
}
