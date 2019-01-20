package com.entity;


/**
 * 最后两列没有获取
 * @author : zyb
 * 时间 : 2019/1/20.
 */
public class Parameter
{

    /**
     * cad编号
     */
    private String cadNumber;

    /**
     * 支撑类型
     */
    private String type;

    /**
     * 芯材牌号
     */
    private String brand;

    /**
     * PK截面_1
     */
    private Double pk_1;
    private Double pk_2;

    /**
     * PK等效面积
     */
    private Double area;

    /**
     * 弹性模量
     */
    private Double elasticModulus;

    /**
     * PK模型轴线长度
      */
    private Double pkAxisLength;

    /**
     * 屈曲约束支撑的刚度
     */
    private Double stiffness;

    public String getCadNumber() {
        return cadNumber;
    }

    public void setCadNumber(String cadNumber) {
        this.cadNumber = cadNumber;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getBrand() {
        return brand;
    }

    public void setBrand(String brand) {
        this.brand = brand;
    }

    public Double getPk_1() {
        return pk_1;
    }

    public void setPk_1(Double pk_1) {
        this.pk_1 = pk_1;
    }

    public Double getPk_2() {
        return pk_2;
    }

    public void setPk_2(Double pk_2) {
        this.pk_2 = pk_2;
    }

    public Double getArea() {
        return area;
    }

    public void setArea(Double area) {
        this.area = area;
    }

    public Double getElasticModulus() {
        return elasticModulus;
    }

    public void setElasticModulus(Double elasticModulus) {
        this.elasticModulus = elasticModulus;
    }

    public Double getPkAxisLength() {
        return pkAxisLength;
    }

    public void setPkAxisLength(Double pkAxisLength) {
        this.pkAxisLength = pkAxisLength;
    }

    public Double getStiffness() {
        return stiffness;
    }

    public void setStiffness(Double stiffness) {
        this.stiffness = stiffness;
    }

    @Override
    public String toString() {
        return "Parameter{" +
                "     " +  cadNumber +
                "     " +  type +
                "     " +  brand +
                "     " +  pk_1 +
                "     "  + pk_2 +
                "     " + area +
                "     " + elasticModulus +
                "     " + pkAxisLength +
                "     " + stiffness +
                '}';
    }
}
