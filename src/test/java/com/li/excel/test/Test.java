package com.li.excel.test;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.li.excel.ExcelUtil;
import com.li.excel.test.model.LoanVo;
import com.li.excel.util.MockDataUtil;

/**
 * 
 * @Title: Test.java 
 * @Package com.li.excel.test 
 * @Description: excel-mokeyking demo 
 * @author leevan
 * @date 2018年11月14日 下午3:01:20
 * @version 1.0.0
 */
public class Test {
    private static Logger logger = LoggerFactory.getLogger(Test.class);

    public static void exportSingleSheetExcel() {
        List<LoanVo> shopDTOList = MockDataUtil.getMockData(35000);
        String filePath = "D:\\demo-sheet.xls";
        ExcelUtil.exportToFile(filePath, shopDTOList, LoanVo.class, 0);
    }

    
    public static void exportMutiSheetExcel() {
        List<LoanVo> shopDTOList = MockDataUtil.getMockData(35000);
        List<LoanVo> nShopDTOList = MockDataUtil.getMockData(5000);
        String mFilePath = "D:\\demo-mutisheet.xls";
        Collection<?>[] collectionArr = new Collection<?>[]{shopDTOList,nShopDTOList};
        int [] intArr = new int[] {0,0};
        ExcelUtil.exportMutiToFile(mFilePath, collectionArr, intArr);
    }
    
    public static void importFromFile() {
        String filePath = "D:\\demo-sheet.xlsx";
        
        List<LoanVo> list = null;
		try {
			list = ExcelUtil.importFromFile(new File(filePath), LoanVo.class);
		} catch (IllegalAccessException | InvalidFormatException | InstantiationException | InvocationTargetException
				| IntrospectionException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        for(Object vo :list) {
            System.out.println(new Gson().toJson(vo));
        }
    }
    
    public static void importFromFile2Map() {
        String filePath = "D:\\demo-sheet.xlsx";
        
        List<Map<String,String>> list = null;
		try {
			list = ExcelUtil.importFromFile(new File(filePath));
		} catch (InvalidFormatException | IOException e) {
			e.printStackTrace();
		}
        for(Object vo :list) {
            System.out.println(new Gson().toJson(vo));
        }
    }
    
    public static void main(String[] args) {
    	//test single sheet excel
    	exportSingleSheetExcel();
    	
    	//test muti sheet excel
    	//exportMutiSheetExcel();

        //test single sheet excel 2 object list
    	//importFromFile();
    	
    	//test single sheet excel 2 map list
    	//importFromFile2Map();

    }

}
