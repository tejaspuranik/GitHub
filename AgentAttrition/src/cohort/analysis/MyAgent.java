package cohort.analysis;

import java.util.ArrayList;
import java.util.Date;

public class MyAgent {
	private String empNo;
	private Date hireDate;
	private String partner;
	private String siteDesc;
	private String location;
	private int attritionCount;
	private int attritionFlag;
	private String PF;
	private int Jan31Diff;
	private int Feb28Diff;
	private int Mar31Diff;
	private int Apr30Diff;
	private int May31Diff;
	private int Jun30Diff;
	private int Jul31Diff;
	private int Aug31Diff;
	private int Sep30Diff;
	private int Oct31Diff;
	private int Nov30Diff;
	private int Dec31Diff;
	private int Jan31Diff18;
	private int Feb28Diff18;
	private int Mar31Diff18;
	public int getJan31Diff18() {
		return Jan31Diff18;
	}
	public void setJan31Diff18(int jan31Diff18) {
		Jan31Diff18 = jan31Diff18;
	}
	public int getFeb28Diff18() {
		return Feb28Diff18;
	}
	public void setFeb28Diff18(int feb28Diff18) {
		Feb28Diff18 = feb28Diff18;
	}
	public int getMar31Diff18() {
		return Mar31Diff18;
	}
	public void setMar31Diff18(int mar31Diff18) {
		Mar31Diff18 = mar31Diff18;
	}
	private int May01Diff;
	private String lob;
	private ArrayList<Double> abtMonths;
	private ArrayList<Float> segMonths;
	private ArrayList<Float> segProjectedMonths;
	private ArrayList<Float> segAttritionMonths;
	public ArrayList<Float> getSegAttritionMonths() {
		return segAttritionMonths;
	}
	public void setSegAttritionMonths(ArrayList<Float> segAttritionMonths) {
		this.segAttritionMonths = segAttritionMonths;
	}
	public ArrayList<Float> getSegProjectedMonths() {
		return segProjectedMonths;
	}
	public void setSegProjectedMonths(ArrayList<Float> segProjectedMonths) {
		this.segProjectedMonths = segProjectedMonths;
	}
	private ArrayList<Double> fcrMonths;
	private ArrayList<Double> sv60Months;
	private ArrayList<String> segNameMonths;
	private ArrayList<String> queueNameMonths;
	
	public ArrayList<String> getSegNameMonths() {
		return segNameMonths;
	}
	public void setSegNameMonths(ArrayList<String> segNameMonths) {
		this.segNameMonths = segNameMonths;
	}
	public ArrayList<String> getQueueNameMonths() {
		return queueNameMonths;
	}
	public void setQueueNameMonths(ArrayList<String> queueNameMonths) {
		this.queueNameMonths = queueNameMonths;
	}
	
	public ArrayList<Double> getFcrMonths() {
		return fcrMonths;
	}
	public void setFcrMonths(ArrayList<Double> fcrMonths) {
		this.fcrMonths = fcrMonths;
	}
	public ArrayList<Double> getSv60Months() {
		return sv60Months;
	}
	public void setSv60Months(ArrayList<Double> sv60Months) {
		this.sv60Months = sv60Months;
	}
	
	public ArrayList<Float> getSegMonths() {
		return segMonths;
	}
	public void setSegMonths(ArrayList<Float> segMonths) {
		this.segMonths = segMonths;
	}
	public String getLob() {
		return lob;
	}
	public void setLob(String lob) {
		this.lob = lob;
	}
	public String getResult() {
		return result;
	}
	public void setResult(String result) {
		this.result = result;
	}
	private String result;
	
	public Date getHireDate() {
		return hireDate;
	}
	public void setHireDate(Date hireDate) {
		this.hireDate = hireDate;
	}
	public String getPartner() {
		return partner;
	}
	public void setPartner(String partner) {
		this.partner = partner;
	}
	public String getSiteDesc() {
		return siteDesc;
	}
	public void setSiteDesc(String siteDesc) {
		this.siteDesc = siteDesc;
	}
	public String getLocation() {
		return location;
	}
	public void setLocation(String location) {
		this.location = location;
	}
	public int getAttritionCount() {
		return attritionCount;
	}
	public void setAttritionCount(int attritionCount) {
		this.attritionCount = attritionCount;
	}
	public int getAttritionFlag() {
		return attritionFlag;
	}
	public void setAttritionFlag(int attritionFlag) {
		this.attritionFlag = attritionFlag;
	}
	public int getJan31Diff() {
		return Jan31Diff;
	}
	public void setJan31Diff(int jan31Diff) {
		Jan31Diff = jan31Diff;
	}
	public int getFeb28Diff() {
		return Feb28Diff;
	}
	public void setFeb28Diff(int feb28Diff) {
		Feb28Diff = feb28Diff;
	}
	public int getMar31Diff() {
		return Mar31Diff;
	}
	public void setMar31Diff(int mar31Diff) {
		Mar31Diff = mar31Diff;
	}
	public int getApr30Diff() {
		return Apr30Diff;
	}
	public void setApr30Diff(int apr30Diff) {
		Apr30Diff = apr30Diff;
	}
	public int getMay31Diff() {
		return May31Diff;
	}
	public void setMay31Diff(int may31Diff) {
		May31Diff = may31Diff;
	}
	public int getJun30Diff() {
		return Jun30Diff;
	}
	public void setJun30Diff(int jun30Diff) {
		Jun30Diff = jun30Diff;
	}
	public int getJul31Diff() {
		return Jul31Diff;
	}
	public void setJul31Diff(int jul31Diff) {
		Jul31Diff = jul31Diff;
	}
	public int getAug31Diff() {
		return Aug31Diff;
	}
	public void setAug31Diff(int aug31Diff) {
		Aug31Diff = aug31Diff;
	}
	public int getSep30Diff() {
		return Sep30Diff;
	}
	public void setSep30Diff(int sep30Diff) {
		Sep30Diff = sep30Diff;
	}
	public int getOct31Diff() {
		return Oct31Diff;
	}
	public void setOct31Diff(int oct31Diff) {
		Oct31Diff = oct31Diff;
	}
	public int getNov30Diff() {
		return Nov30Diff;
	}
	public void setNov30Diff(int nov30Diff) {
		Nov30Diff = nov30Diff;
	}
	public int getDec31Diff() {
		return Dec31Diff;
	}
	public void setDec31Diff(int dec31Diff) {
		Dec31Diff = dec31Diff;
	}
	public int getMay01Diff() {
		return May01Diff;
	}
	public void setMay01Diff(int may01Diff) {
		May01Diff = may01Diff;
	}
	private float s1;
	private float s2;
	private float s3;
	private float s4;
	private float s5;
	private float s6;
	private float s7;
	private float s8;
	private float s9;
	private float s10;
	private float s11;
	private float s12;
	private float s13;
	private float s14;
	private float s15;
	
	
	public String getEmpNo() {
		return empNo;
	}
	public void setEmpNo(String empNo) {
		this.empNo = empNo;
	}
	public float getS1() {
		return s1;
	}
	public void setS1(float s1) {
		this.s1 = s1;
	}
	public float getS2() {
		return s2;
	}
	public void setS2(float s2) {
		this.s2 = s2;
	}
	public float getS3() {
		return s3;
	}
	public void setS3(float s3) {
		this.s3 = s3;
	}
	public float getS4() {
		return s4;
	}
	public void setS4(float s4) {
		this.s4 = s4;
	}
	public float getS5() {
		return s5;
	}
	public void setS5(float s5) {
		this.s5 = s5;
	}
	public float getS6() {
		return s6;
	}
	public void setS6(float s6) {
		this.s6 = s6;
	}
	public float getS7() {
		return s7;
	}
	public void setS7(float s7) {
		this.s7 = s7;
	}
	public float getS8() {
		return s8;
	}
	public void setS8(float s8) {
		this.s8 = s8;
	}
	public float getS9() {
		return s9;
	}
	public void setS9(float s9) {
		this.s9 = s9;
	}
	public float getS10() {
		return s10;
	}
	public void setS10(float s10) {
		this.s10 = s10;
	}
	public float getS11() {
		return s11;
	}
	public void setS11(float s11) {
		this.s11 = s11;
	}
	public float getS12() {
		return s12;
	}
	public void setS12(float s12) {
		this.s12 = s12;
	}
	public float getS13() {
		return s13;
	}
	public void setS13(float s13) {
		this.s13 = s13;
	}
	public float getS14() {
		return s14;
	}
	public void setS14(float s14) {
		this.s14 = s14;
	}
	public float getS15() {
		return s15;
	}
	public void setS15(float s15) {
		this.s15 = s15;
	}
	public String getPF() {
		return PF;
	}
	public void setPF(String pF) {
		PF = pF;
	}
	public ArrayList<Double> getAbtMonths() {
		return abtMonths;
	}
	public void setAbtMonths(ArrayList<Double> abtMonths) {
		this.abtMonths = abtMonths;
	}

	
}
