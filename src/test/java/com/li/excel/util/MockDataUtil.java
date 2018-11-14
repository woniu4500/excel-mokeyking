package com.li.excel.util;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.li.excel.test.model.LoanVo;

/**
 * 
 * @Title: MockDataUtil.java 
 * @Package com.li.excel.util 
 * @Description: mock demo data 
 * @author leevan
 * @date 2018年11月14日 下午2:46:54
 * @version 1.0.0
 */
public class MockDataUtil {
    
	public static List<LoanVo> getMockData(int n) {
		//mock数据
        List<LoanVo> loanVoList = new ArrayList<LoanVo>();
        for (int i = 0; i < n; i++) {
            LoanVo loanVo = new LoanVo("20181114a"+leftFill6Zero(i), "李大大"+i, "18682"+leftFill6Zero(i), 
            		"412723198409"+leftFill6Zero(i), new Date(), "", new Date(), new Date(), new BigDecimal(10000.00), 6);
            loanVoList.add(loanVo);
        }
		return loanVoList;
	}
    
	public static String leftFill6Zero(int i) {
		String zero6 = "000000";
		return zero6.substring(0,6-Integer.toString(i).length())+Integer.toString(i);
	}
	
	public static void main(String [] args) {
		System.out.println(leftFill6Zero(1));
	}
}
