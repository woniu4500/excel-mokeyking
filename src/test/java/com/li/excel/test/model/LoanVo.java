package com.li.excel.test.model;

import java.math.BigDecimal;
import java.util.Date;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.li.excel.annotation.ExcelField;
import com.li.excel.annotation.ExcelSheet;

/**
 * 
 * @Title: LoanVo.java 
 * @Package com.li.excel.test.model 
 * @Description: demo data  
 * @author leevan
 * @date 2018年11月14日 下午2:59:04
 * @version 1.0.0
 */
@ExcelSheet(name = "贷款列表", headColor = HSSFColor.HSSFColorPredefined.PALE_BLUE)
public class LoanVo {
    @ExcelField(name = "借据编号",width = 20)
	private String billNo;
	
    @ExcelField(name = "姓名",width = 10)
	private String name;
	
    @ExcelField(name = "手机号码", order = 10, tags = {2, 3}, width = 20)
	private String mobile;
	
    @ExcelField(name = "证件号码", tags = {1},width = 50)
	private String certCode;
	
    @ExcelField(name = "生日", format = "yyyy-MM-dd", width = 10)
	private Date birthday;
	
    @ExcelField(name = "性别", defaultValue = "保密")
	private String sex;
	
    @ExcelField(name = "贷款起始日期", format = "yyyy-MM-dd", width = 10)
	private Date loanStartDate;
    
    @ExcelField(name = "贷款结束日期", format = "yyyy-MM-dd", width = 10)
	private Date loanEndDate;
    
    @ExcelField(name = "贷款金额",format = "0.00", align = HorizontalAlignment.RIGHT, width = 10)
	private BigDecimal loanAmt;
	
    @ExcelField(name = "贷款期数",string = "共{{value}}期")
	private int loanTerm;
    
	public LoanVo(String billNo, String name, String mobile, String certCode, Date birthday, String sex,
			Date loanStartDate, Date loanEndDate, BigDecimal loanAmt, int loanTerm) {
		super();
		this.billNo = billNo;
		this.name = name;
		this.mobile = mobile;
		this.certCode = certCode;
		this.birthday = birthday;
		this.sex = sex;
		this.loanStartDate = loanStartDate;
		this.loanEndDate = loanEndDate;
		this.loanAmt = loanAmt;
		this.loanTerm = loanTerm;
	}

	public String getBillNo() {
		return billNo;
	}

	public String getName() {
		return name;
	}

	public String getMobile() {
		return mobile;
	}

	public String getCertCode() {
		return certCode;
	}

	public Date getBirthday() {
		return birthday;
	}

	public String getSex() {
		return sex;
	}

	public Date getLoanStartDate() {
		return loanStartDate;
	}

	public Date getLoanEndDate() {
		return loanEndDate;
	}

	public BigDecimal getLoanAmt() {
		return loanAmt;
	}

	public int getLoanTerm() {
		return loanTerm;
	}

	public void setBillNo(String billNo) {
		this.billNo = billNo;
	}

	public void setName(String name) {
		this.name = name;
	}

	public void setMobile(String mobile) {
		this.mobile = mobile;
	}

	public void setCertCode(String certCode) {
		this.certCode = certCode;
	}

	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}

	public void setSex(String sex) {
		this.sex = sex;
	}

	public void setLoanStartDate(Date loanStartDate) {
		this.loanStartDate = loanStartDate;
	}

	public void setLoanEndDate(Date loanEndDate) {
		this.loanEndDate = loanEndDate;
	}

	public void setLoanAmt(BigDecimal loanAmt) {
		this.loanAmt = loanAmt;
	}

	public void setLoanTerm(int loanTerm) {
		this.loanTerm = loanTerm;
	}


}
