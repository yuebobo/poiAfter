package com.insert;

import com.entity.Constants;
import com.util.Util;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.util.Map;

/**
 * 计算最后几个表格
 * 时间 : 2018/10/6.
 */
public class CaculateTable {

    /**
     * 子结构框架梁 受弯受剪 验算
     * @param table
     * @param caculateParams
     */
    public static void caculateTable1(XWPFTable table, Map<String, Object> caculateParams) {

        //1.从Map里获取计算参数
        //混泥土等级
        String concreteGrade = caculateParams.get(Constants.CONCRETE_GRADE).toString();
        String concreteGradeFc = caculateParams.get(Constants.CONCRETE_GRADE_FC).toString();
        String concreteGradeFck = caculateParams.get(Constants.CONCRETE_GRADE_FCK).toString();
        String concreteGradeFt = caculateParams.get(Constants.CONCRETE_GRADE_FT).toString();
        String concreteGradeFtk = caculateParams.get(Constants.CONCRETE_GRADE_FTK).toString();

        //钢筋等级
        String steelGrade = caculateParams.get(Constants.STEEL_GRADE).toString();
        String steelGradeFstk = caculateParams.get(Constants.STEEL_GRADE_FSTK).toString();
        String steelGradeFyk = caculateParams.get(Constants.STEEL_GRADE_FYK).toString();
        String steelGradeFyvk = caculateParams.get(Constants.STEEL_GRADE_FYVK).toString();

        //截面属性(mm)
        String sectionB = caculateParams.get(Constants.SECTION_B).toString();
        String sectionH = caculateParams.get(Constants.SECTION_H).toString();

        //受力状况(kN,KN*m)
        String stressConditionV = caculateParams.get(Constants.STRESS_CONDITION_V).toString();
        String stressConditionM = caculateParams.get(Constants.STRESS_CONDITION_M).toString();

        //箍筋(mm)
        String hoopD = caculateParams.get(Constants.HOOP_D).toString();
        String hoopL = caculateParams.get(Constants.HOOP_L).toString();


        String _afS = caculateParams.get(Constants.PARAM_afS).toString();
        String _af1 = caculateParams.get(Constants.PARAM_af1).toString();
        Double _yre = 0.85;
        Double _afCV = 0.7;

        System.out.println("=============== 计算参数 ===================");
        System.out.println("混泥土等级：" + concreteGrade);
        System.out.println("混泥土等级参数：\nFc= " + concreteGradeFc + "\nFck= " + concreteGradeFck + "\nFt= " + concreteGradeFt + "\nFtk= " + concreteGradeFtk);
        System.out.println("钢筋等级：" + steelGrade);
        System.out.println("钢筋等级：\nFstk= " + steelGradeFstk + "\nFyk= " + steelGradeFyk + "\nFyvk= " + steelGradeFyvk);
        System.out.println("截面属性：\nb= " + sectionB + "\nh= " + sectionH);
        System.out.println("受力状况：\nv= " + stressConditionV + "\nM= " + stressConditionM);
        System.out.println("箍筋：\nd= " + hoopD + "\nl= " + hoopL);
        System.out.println("剩余参数：\n_afS= " + _afS + "\n_af1= " + _af1 + "\n_yre= " + _yre + "\n_afCV= " + _afCV);

        //3.根据公式进行计算
        //1).受弯验算
        Double h0 = Double.valueOf(sectionH) - Double.valueOf(_afS);
        Double _afs_ = Double.valueOf(stressConditionM) * 1000000 / (Double.valueOf(_af1) * Double.valueOf(concreteGradeFck) * Double.valueOf(sectionB) * h0 * h0);
        Double _ypclong_ = 1.00 - Math.pow(Double.valueOf(1 - (2 * _afs_)),0.5);
        Double _As_ = _ypclong_ * Double.valueOf(sectionB) * h0 * Double.valueOf(_af1) * Double.valueOf(concreteGradeFck) / Double.valueOf(steelGradeFstk);
        Double _Ps_ = _As_ / (Double.valueOf(sectionB) * h0);

        //1).受剪验算
        Double _shearForce1_ = 0.2 * Double.valueOf(concreteGradeFck) * Double.valueOf(sectionB) * h0 / (_yre * 1000);
        Double _shearForce2_ = ((0.6 * _afCV * Double.valueOf(concreteGradeFtk) * Double.valueOf(sectionB) * h0)+
                ((Double.valueOf(steelGradeFyvk) * 4 * 0.25 * Math.PI * Math.pow(Double.valueOf(hoopD),2) * h0) / Double.valueOf(hoopL))) / (_yre * 1000);
        System.out.println("=============== 计算结果 ===================");
        System.out.println("受弯验算 ：\n_afs_  = "+ _afs_ + "\n_ypclong_ = " + _ypclong_ + "\n_As_ = "+ _As_ + "\n_Ps_ = "+_Ps_ );
        System.out.println("受剪验算 ：\n_shearForce1_  = "+ _shearForce1_ + "\n_shearForce1_ = " + _shearForce1_ );

        //4.比较数据，得出结论
        String _Ps1_ = table.getRow(9).getCell(6).getText();
        Double _Ps11_ = Double.valueOf(_Ps1_.substring(1,_Ps1_.length() - 1));
        String result ;
        if (_Ps_ < _Ps11_ && _shearForce1_ > Double.valueOf(stressConditionV) && _shearForce2_ > Double.valueOf(stressConditionV)){
            //通过
            result = "通过";
        }else {
            result = "不通过";
        }

        //5.将计算的结果回填到表格中
        Util.insertValueToCell(table.getRow(0).getCell(2),Util.getPrecisionString(sectionB,1));
        Util.insertValueToCell(table.getRow(0).getCell(5),Util.getPrecisionString(sectionH,1));
        Util.insertValueToCell(table.getRow(1).getCell(2),Util.getPrecisionString(stressConditionV,1));
        Util.insertValueToCell(table.getRow(1).getCell(5),Util.getPrecisionString(stressConditionM,1));
        Util.insertValueToCell(table.getRow(2).getCell(2),concreteGrade);
        Util.insertValueToCell(table.getRow(3).getCell(2),Util.getPrecisionString(concreteGradeFck,2));
        Util.insertValueToCell(table.getRow(3).getCell(5),Util.getPrecisionString(_afS,2));
        Util.insertValueToCell(table.getRow(4).getCell(2),Util.getPrecisionString(concreteGradeFc,2));
        Util.insertValueToCell(table.getRow(4).getCell(5),Util.getPrecisionString(concreteGradeFtk,2));
        Util.insertValueToCell(table.getRow(5).getCell(2),steelGrade);
        Util.insertValueToCell(table.getRow(6).getCell(2),Util.getPrecisionString(steelGradeFstk,0));
        Util.insertValueToCell(table.getRow(6).getCell(5),Util.getPrecisionString(steelGradeFyvk,0));
        Util.insertValueToCell(table.getRow(7).getCell(2),Util.getPrecisionString(hoopD,0));
        Util.insertValueToCell(table.getRow(7).getCell(5),Util.getPrecisionString(hoopL,0));
        Util.insertValueToCell(table.getRow(8).getCell(2),Util.getPrecisionString(_afs_,3));
        Util.insertValueToCell(table.getRow(8).getCell(4),Util.getPrecisionString(_ypclong_,3));
        Util.insertValueToCell(table.getRow(9).getCell(2),Util.getPrecisionString(_As_,0));
        Util.insertValueToCell(table.getRow(9).getCell(5),Util.getPrecisionString(_Ps_ * 100,2)+"%");
        Util.insertValueToCell(table.getRow(10).getCell(2),Util.getPrecisionString(_shearForce1_,0));
        Util.insertValueToCell(table.getRow(10).getCell(4),result);
        Util.insertValueToCell(table.getRow(11).getCell(2),Util.getPrecisionString(_shearForce2_,0));
    }

    /**
     * 子结构框架柱抗剪验算
     * @param table
     * @param caculateParams
     */
    public static void caculateTable2(XWPFTable table, Map<String, Object> caculateParams) {
        //1.从Map里获取计算参数
        //混泥土等级
        String concreteGrade = caculateParams.get(Constants.CONCRETE_GRADE).toString();
        String concreteGradeFc = caculateParams.get(Constants.CONCRETE_GRADE_FC).toString();
        String concreteGradeFck = caculateParams.get(Constants.CONCRETE_GRADE_FCK).toString();
        String concreteGradeFt = caculateParams.get(Constants.CONCRETE_GRADE_FT).toString();
        String concreteGradeFtk = caculateParams.get(Constants.CONCRETE_GRADE_FTK).toString();

        //钢筋等级
        String steelGrade = caculateParams.get(Constants.STEEL_GRADE).toString();
        String steelGradeFstk = caculateParams.get(Constants.STEEL_GRADE_FSTK).toString();
        String steelGradeFyk = caculateParams.get(Constants.STEEL_GRADE_FYK).toString();
        String steelGradeFyvk = caculateParams.get(Constants.STEEL_GRADE_FYVK).toString();

        //截面属性(mm)
        String sectionB = caculateParams.get(Constants.SECTION_B).toString();
        String sectionH = caculateParams.get(Constants.SECTION_H).toString();

        //受力状况(kN,KN*m)
        String stressConditionV = caculateParams.get(Constants.STRESS_CONDITION_V).toString();
        String stressConditionM = caculateParams.get(Constants.STRESS_CONDITION_M).toString();
        String stressConditionP = caculateParams.get(Constants.STRESS_CONDITION_P).toString();

        //箍筋(mm)
        String hoopD = caculateParams.get(Constants.HOOP_D).toString();
        String hoopL = caculateParams.get(Constants.HOOP_L).toString();

        //楼层高度
        String floorH = caculateParams.get(Constants.FLOOR_H).toString();

        String _afS = caculateParams.get(Constants.PARAM_afS).toString();
        Double _yre = 0.85;
        Double _afCV = 0.7;

        System.out.println("=============== 计算参数 ===================");
        System.out.println("混泥土等级：" + concreteGrade);
        System.out.println("混泥土等级参数：\nFc= " + concreteGradeFc + "\nFck= " + concreteGradeFck + "\nFt= " + concreteGradeFt + "\nFtk= " + concreteGradeFtk);
        System.out.println("钢筋等级：" + steelGrade);
        System.out.println("钢筋等级：\nFstk= " + steelGradeFstk + "\nFyk= " + steelGradeFyk + "\nFyvk= " + steelGradeFyvk);
        System.out.println("截面属性：\nb= " + sectionB + "\nh= " + sectionH);
        System.out.println("受力状况：\nv= " + stressConditionV + "\nM= " + stressConditionM + "\nP= "+stressConditionP);
        System.out.println("箍筋：\nd= " + hoopD + "\nl= " + hoopL);
        System.out.println("层高： floorH= "+floorH);
        System.out.println("剩余参数：\n_afS= " + _afS + "\n_yre= " + _yre + "\n_afCV= " + _afCV);

        //3.根据公式进行计算
        Double h0 = Double.valueOf(sectionH) - Double.valueOf(_afS);
        Double _result1_ = 0.2 * Double.valueOf(concreteGradeFck) * Double.valueOf(sectionB) * h0 / (_yre * 1000 );
        Double _lamda_ = (Double.valueOf(floorH) - Double.valueOf(_afS)) / (2 * h0);
        Double _result3_ = (((1.05 * Double.valueOf(concreteGradeFtk) * Double.valueOf(sectionB) * h0) / (_lamda_ + 1.00)) +
                            (Double.valueOf(steelGradeFyvk) * 4 * 0.25 * Math.PI * Math.pow(Double.valueOf(hoopD),2) * h0 / Double.valueOf(hoopL)) +
                            (0.056 * Double.valueOf(stressConditionP) * 1000))/ (_yre * 1000);
        System.out.println("=============== 计算结果 ===================");
        System.out.println("受剪验算 ：\n_result1_  = "+ _result1_ + "\n_lamda_ = " + _lamda_ + "\n_result3_ =  "+_result3_ );


        //4.比较数据，得出结论
        String result ;
        if (_result1_ > Double.valueOf(stressConditionV) && _result3_ > Double.valueOf(stressConditionV) ){
            //通过
            result = "通过";
        }else {
            result = "不通过";
        }

        //5.将计算的结果回填到表格中
        Util.insertValueToCell(table.getRow(0).getCell(2),Util.getPrecisionString(sectionB,1));
        Util.insertValueToCell(table.getRow(0).getCell(5),Util.getPrecisionString(sectionH,1));
        Util.insertValueToCell(table.getRow(1).getCell(1),Util.getPrecisionString(floorH,0));
        Util.insertValueToCell(table.getRow(2).getCell(2),Util.getPrecisionString(stressConditionV,1));
        Util.insertValueToCell(table.getRow(2).getCell(5),Util.getPrecisionString(stressConditionM,1));
        Util.insertValueToCell(table.getRow(3).getCell(2),Util.getPrecisionString(stressConditionP,1));
        Util.insertValueToCell(table.getRow(4).getCell(2),concreteGrade);
        Util.insertValueToCell(table.getRow(5).getCell(2),Util.getPrecisionString(concreteGradeFck,2));
        Util.insertValueToCell(table.getRow(5).getCell(5),Util.getPrecisionString(_afS,2));
        Util.insertValueToCell(table.getRow(6).getCell(2),Util.getPrecisionString(concreteGradeFc,2));
        Util.insertValueToCell(table.getRow(6).getCell(5),Util.getPrecisionString(concreteGradeFtk,2));
        Util.insertValueToCell(table.getRow(7).getCell(2),steelGrade);
        Util.insertValueToCell(table.getRow(8).getCell(2),Util.getPrecisionString(steelGradeFstk,0));
        Util.insertValueToCell(table.getRow(8).getCell(5),Util.getPrecisionString(steelGradeFyvk,0));
        Util.insertValueToCell(table.getRow(9).getCell(2),Util.getPrecisionString(hoopD,0));
        Util.insertValueToCell(table.getRow(9).getCell(5),Util.getPrecisionString(hoopL,0));

        Util.insertValueToCell(table.getRow(10).getCell(2),Util.getPrecisionString(_result1_,0));
        Util.insertValueToCell(table.getRow(10).getCell(4),result);
        Util.insertValueToCell(table.getRow(11).getCell(2),Util.getPrecisionString(_lamda_,2));
        Util.insertValueToCell(table.getRow(12).getCell(2),Util.getPrecisionString(_result3_,0));
    }

    /**
     * 悬臂墙配筋验算
     * @param table
     * @param caculateParams
     */
    public static void caculateTable3(XWPFTable table, Map<String, Object> caculateParams) {
//1.从Map里获取计算参数
        //混泥土等级
        String concreteGrade = caculateParams.get(Constants.CONCRETE_GRADE).toString();
        String concreteGradeFc = caculateParams.get(Constants.CONCRETE_GRADE_FC).toString();
        String concreteGradeFck = caculateParams.get(Constants.CONCRETE_GRADE_FCK).toString();
        String concreteGradeFt = caculateParams.get(Constants.CONCRETE_GRADE_FT).toString();
        String concreteGradeFtk = caculateParams.get(Constants.CONCRETE_GRADE_FTK).toString();

        //钢筋等级
        String steelGrade = caculateParams.get(Constants.STEEL_GRADE).toString();
        String steelGradeFstk = caculateParams.get(Constants.STEEL_GRADE_FSTK).toString();
        String steelGradeFyk = caculateParams.get(Constants.STEEL_GRADE_FYK).toString();
        String steelGradeFyvk = caculateParams.get(Constants.STEEL_GRADE_FYVK).toString();

        //截面属性(mm)
        String sectionB = caculateParams.get(Constants.SECTION_B).toString();
        String sectionH = caculateParams.get(Constants.SECTION_H).toString();

        //受力状况(kN,KN*m)
        String stressConditionV = caculateParams.get(Constants.STRESS_CONDITION_V).toString();

        //箍筋(mm)
        String hoopD = caculateParams.get(Constants.HOOP_D).toString();
        String hoopL = caculateParams.get(Constants.HOOP_L).toString();

        //纵筋
        String hoopVerticlD = caculateParams.get(Constants.HOOP_VERTICl_D).toString();
        String hoopVerticlL = caculateParams.get(Constants.HOOP_VERTICl_L).toString();

        //楼层高度
        String floorH = caculateParams.get(Constants.FLOOR_H).toString();
        //阻尼器高度
        String damperH = caculateParams.get(Constants.DAMPER_H).toString();

        String _afS = caculateParams.get(Constants.PARAM_afS).toString();
        Double _yre = 0.85;
        Double _afCV = 0.7;

        System.out.println("=============== 计算参数 ===================");
        System.out.println("混泥土等级：" + concreteGrade);
        System.out.println("混泥土等级参数：\nFc= " + concreteGradeFc + "\nFck= " + concreteGradeFck + "\nFt= " + concreteGradeFt + "\nFtk= " + concreteGradeFtk);
        System.out.println("钢筋等级：" + steelGrade);
        System.out.println("钢筋等级：\nFstk= " + steelGradeFstk + "\nFyk= " + steelGradeFyk + "\nFyvk= " + steelGradeFyvk);
        System.out.println("截面属性：\nb= " + sectionB + "\nh= " + sectionH);
        System.out.println("受力状况：\nv= " + stressConditionV );
        System.out.println("箍筋：\nd= " + hoopD + "\nl= " + hoopL);
        System.out.println("纵筋：\nd= " + hoopVerticlD + "\nl= " + hoopVerticlL);
        System.out.println("层高： floorH= "+floorH);
        System.out.println("阻尼器高度： damperH= "+damperH);
        System.out.println("剩余参数：\n_afS= " + _afS + "\n_yre= " + _yre + "\n_afCV= " + _afCV);

        //3.根据公式进行计算
        Double h0 = Double.valueOf(sectionH) - Double.valueOf(_afS);
        Double _M_ = (Double.valueOf(floorH) - Double.valueOf(damperH)) * Double.valueOf(stressConditionV) * 0.5;
        Double _As_ = _M_  * 1000000/ (Double.valueOf(steelGradeFstk) * (h0 - Double.valueOf(_afS)));
        Double _result_2 = (0.15 * Double.valueOf(concreteGradeFc) * Double.valueOf(sectionB) * h0) / (_yre * 1000);
        Double _result_3 = ((0.6 * _afCV * Double.valueOf(concreteGradeFt) * Double.valueOf(sectionB) * h0) +
                (Double.valueOf(steelGradeFyvk) * 4 * 0.25 * Math.PI * Math.pow(Double.valueOf(hoopD),2) * h0 / Double.valueOf(hoopL))) / (_yre * 1000);

        System.out.println("=============== 计算结果 ===================");
        System.out.println("受剪验算 ：\n_M_  = "+ _M_ + "\n_As_ = " + _As_ + "\n_result_2 =  "+_result_2 + "\n_result_3= "+_result_3 );

        //4.比较数据，得出结论
        Double _As1_ = Double.valueOf(hoopVerticlL) * 0.25 * Math.PI * Math.pow(Double.valueOf(hoopVerticlD),2);

        String result ;
        if (_As1_ > _As_ && _result_2 > Double.valueOf(stressConditionV) && _result_3 > Double.valueOf(stressConditionV)){
            //通过
            result = "通过";
        }else {
            result = "不通过";
        }

        //4.将计算的结果回填到表格中
        Util.insertValueToCell(table.getRow(0).getCell(2),Util.getPrecisionString(sectionB,1));
        Util.insertValueToCell(table.getRow(0).getCell(5),Util.getPrecisionString(sectionH,1));
        Util.insertValueToCell(table.getRow(1).getCell(1),Util.getPrecisionString(floorH,3));
        Util.insertValueToCell(table.getRow(1).getCell(5),Util.getPrecisionString(damperH,3));
        Util.insertValueToCell(table.getRow(2).getCell(2),Util.getPrecisionString(stressConditionV,1));
        Util.insertValueToCell(table.getRow(2).getCell(5),Util.getPrecisionString(_M_.toString(),1));

        Util.insertValueToCell(table.getRow(3).getCell(2),concreteGrade);
        Util.insertValueToCell(table.getRow(4).getCell(2),Util.getPrecisionString(concreteGradeFck,1));
        Util.insertValueToCell(table.getRow(4).getCell(5),Util.getPrecisionString(_afS,1));
        Util.insertValueToCell(table.getRow(5).getCell(2),Util.getPrecisionString(concreteGradeFc,1));
        Util.insertValueToCell(table.getRow(5).getCell(5),Util.getPrecisionString(concreteGradeFtk,1));

        Util.insertValueToCell(table.getRow(6).getCell(2),steelGrade);
        Util.insertValueToCell(table.getRow(7).getCell(2),Util.getPrecisionString(steelGradeFstk,0));
        Util.insertValueToCell(table.getRow(7).getCell(5),Util.getPrecisionString(steelGradeFyvk,0));

        Util.insertValueToCell(table.getRow(8).getCell(2),Util.getPrecisionString(hoopD,0));
        Util.insertValueToCell(table.getRow(8).getCell(5),Util.getPrecisionString(hoopL,0));

        Util.insertValueToCell(table.getRow(9).getCell(2),Util.getPrecisionString(hoopVerticlD,0));
        Util.insertValueToCell(table.getRow(9).getCell(5),Util.getPrecisionString(hoopVerticlL,0));

        Util.insertValueToCell(table.getRow(10).getCell(2),Util.getPrecisionString(_As_,0));
        Util.insertValueToCell(table.getRow(11).getCell(2),Util.getPrecisionString(_result_2,0));
        Util.insertValueToCell(table.getRow(11).getCell(4),result);
        Util.insertValueToCell(table.getRow(12).getCell(2),Util.getPrecisionString(_result_3,0));
    }
}
