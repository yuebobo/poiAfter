package com.entity;

/**
 *
 * @author : zyb
 * 时间 : 2019/1/20.
 */
public class FloorParameter {

    /**
     * 层高
     */
    private Double floorH;

    /**
     * 累计层高
     */
    private Double addUpFloorH;

    /**
     * 支撑编号
     */
    private String number;

    /**
     * 支撑类型
     */
    private String type;

    /**
     * 芯材牌号
     */
    private String brand;

    /**
     * 屈服力
     */
    private Double force;

    /**
     * 屈服位移
     */
    private Double displacement;

    /**
     * 屈服后刚度
     */
    private Double stiffness;

    /**
     * 外观形状
     */
    private String shape;

    /**
     * 根数
     */
    private Integer count;

    public Double getFloorH() {
        return floorH;
    }

    public void setFloorH(Double floorH) {
        this.floorH = floorH;
    }

    public Double getAddUpFloorH() {
        return addUpFloorH;
    }

    public void setAddUpFloorH(Double addUpFloorH) {
        this.addUpFloorH = addUpFloorH;
    }

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
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

    public Double getForce() {
        return force;
    }

    public void setForce(Double force) {
        this.force = force;
    }

    public Double getDisplacement() {
        return displacement;
    }

    public void setDisplacement(Double displacement) {
        this.displacement = displacement;
    }

    public Double getStiffness() {
        return stiffness;
    }

    public void setStiffness(Double stiffness) {
        this.stiffness = stiffness;
    }

    public String getShape() {
        return shape;
    }

    public void setShape(String shape) {
        this.shape = shape;
    }

    public Integer getCount() {
        return count;
    }

    public void setCount(Integer count) {
        this.count = count;
    }

    @Override
    public String toString() {
        return "FloorParameter{" + floorH +
                "     " + addUpFloorH +
                "     " +  number +
                "     " +  type +
                "     " +  brand +
                "     " +  force +
                "     " +  displacement +
                "     " +  stiffness +
                "     " +  shape +
                "     " +  count +
                '}';
    }
}
