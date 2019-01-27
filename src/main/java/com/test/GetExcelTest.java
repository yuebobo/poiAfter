package com.test;

import com.file.GetExcelValue;
import com.util.Util;

import java.util.ArrayList;
import java.util.List;

public class GetExcelTest {


    public static void  main(String args[]){

        Double[][][] a = new Double[2][2][2];
        Util.setZero(a,0D,Double.class);

//        String basePath = "C:\\Users\\lizhongxiang\\Desktop\\第二套\\源文件\\excel\\材料数据.xlsx";
        String basePath = "C:\\Users\\lizhongxiang\\Desktop\\workSpace\\BRB模板";
        String path1 = basePath + "\\excel\\材料数据.xlsx";
        String path2 = basePath + "\\excel\\参数表.xlsx";
        GetExcelValue.init(path1,path2);

        String p = "C:\\Users\\lizhongxiang\\Desktop\\workSpace\\BRB模板\\excel\\工作簿3.xlsx";
        GetExcelValue.getDamperDisEnergyForceAndDeformation(p);
//        int a = 4%2;
//        System.out.println(a);

//        String path = "C:\\Users\\lizhongxiang\\Desktop\\数据\\工作簿4.xlsx";
//        GetExcelValue.getEarthquakeDamperDisEnergyY(path);

    }
}
