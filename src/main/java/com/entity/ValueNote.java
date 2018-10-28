package com.entity;

public class ValueNote {

	private String floor;
	
	private double t1x = 0;
	private double t2x = 0;
	private double t3x = 0;
	private double t4x = 0;
	private double t5x = 0;
	private double r1x = 0;
	private double r2x = 0;
	
	private double t1y = 0;
	private double t2y = 0;
	private double t3y = 0;
	private double t4y = 0;
	private double t5y = 0;
	private double r1y = 0;
	private double r2y = 0;
	
	public ValueNote(String floor){
		this.floor = floor;
	}
	
	public String getFloor() {
		return floor;
	}
	public void setFloor(String floor) {
		this.floor = floor;
	}
	public double getT1x() {
		return t1x;
	}
	public void setT1x(double t1x) {
		this.t1x = t1x;
	}
	public double getT2x() {
		return t2x;
	}
	public void setT2x(double t2x) {
		this.t2x = t2x;
	}
	public double getT3x() {
		return t3x;
	}
	public void setT3x(double t3x) {
		this.t3x = t3x;
	}
	public double getT4x() {
		return t4x;
	}
	public void setT4x(double t4x) {
		this.t4x = t4x;
	}
	public double getT5x() {
		return t5x;
	}
	public void setT5x(double t5x) {
		this.t5x = t5x;
	}
	public double getR1x() {
		return r1x;
	}
	public void setR1x(double r1x) {
		this.r1x = r1x;
	}
	public double getR2x() {
		return r2x;
	}
	public void setR2x(double r2x) {
		this.r2x = r2x;
	}
	public double getT1y() {
		return t1y;
	}
	public void setT1y(double t1y) {
		this.t1y = t1y;
	}
	public double getT2y() {
		return t2y;
	}
	public void setT2y(double t2y) {
		this.t2y = t2y;
	}
	public double getT3y() {
		return t3y;
	}
	public void setT3y(double t3y) {
		this.t3y = t3y;
	}
	public double getT4y() {
		return t4y;
	}
	public void setT4y(double t4y) {
		this.t4y = t4y;
	}
	public double getT5y() {
		return t5y;
	}
	public void setT5y(double t5y) {
		this.t5y = t5y;
	}
	public double getR1y() {
		return r1y;
	}
	public void setR1y(double r1y) {
		this.r1y = r1y;
	}
	public double getR2y() {
		return r2y;
	}
	public void setR2y(double r2y) {
		this.r2y = r2y;
	}

	@Override
	public String toString() {
		return "floor=" + floor + "\n t1x=" + t1x + ", t2x=" + t2x + ", t3x=" + t3x + ", t4x=" + t4x
				+ ", t5x=" + t5x + ", r1x=" + r1x + ", r2x=" + r2x + ", \nt1y=" + t1y + ", t2y=" + t2y + ", t3y=" + t3y
				+ ", t4y=" + t4y + ", t5y=" + t5y + ", r1y=" + r1y + ", r2y=" + r2y;
	}
	
}
