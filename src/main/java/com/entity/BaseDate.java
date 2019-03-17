package com.entity;


import java.util.List;
import java.util.Map;

/**
 * 时间 : 2018/10/30.
 */
public class BaseDate {

    //层高
    public  Double[] FLOOR_H;

    //累计层高
    public  Double[] ACCOUNT_FLOOR_H;

    //CAD模型编号  第一维表示楼层，第二维表示编号
    //X方向
    public  String[][] CAD_MODEL_X;
    //Y方向
    public  String[][] CAD_MODEL_Y;

    //SAP编号
    //X方向
    public  Double[]  SAP_NO_X;
    //Y方向
    public  Double[]  SAP_NO_Y;

    //周期
    public String[] CECLE;

    //质量
    public String QUALITY;

    //减震剪力
    public  List[] FX_FY;

    //非减震剪力
    public List[] NOT_FX_FY;

    //层间屈服剪力
    public List[] YIELD_FORCE;

    //梁
    public Map<String,Object> GIRDER_PARAMS ;
    //柱
    public Map<String,Object> PILLAR_PARAMS;
    //悬臂
    public Map<String,Object> CANTILEVER_PARAMS ;
}
