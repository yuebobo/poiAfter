package com.test;

import com.file.GetExcelValue;

import java.util.ArrayList;
import java.util.List;

public class GetExcelTest {


    public static void  main(String args[]){
        String path = "C:\\Users\\lizhongxiang\\Desktop\\第二套\\源文件\\excel\\材料数据.xlsx";
        GetExcelValue.init(path);

        List<String> s = new ArrayList<>();
        int x = s.size();
        s.add("d");
        int y = s.size();

        String str = null;
        String ss = str+"";
        System.out.println(ss.length());
//        int a = 4%2;
//        System.out.println(a);

//        String path = "C:\\Users\\lizhongxiang\\Desktop\\数据\\工作簿4.xlsx";
//        GetExcelValue.getEarthquakeDamperDisEnergyY(path);

    }
}
