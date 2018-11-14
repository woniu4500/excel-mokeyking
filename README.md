# excel-mokeyking
	基于注解导入导出excel（可以基于注解定制sheet名称，表头以及对数据日期格式的更加灵活的控制，比如数字金额右靠，文本左靠。）

## Example

### annotation
```java
    @ExcelField(name = "借据编号",width = 20)
	private String billNo;
	
    @ExcelField(name = "姓名",width = 10)
	private String name;
	
    @ExcelField(name = "手机号码", order = 10, group = {2, 3}, width = 20)
	private String mobile;
	
    @ExcelField(name = "证件号码", group = {1},width = 50)
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
    
    注解说明：
    name表示列名,format表示格式化,defaultValue表示设置替代默认值,string拼接替代字符串，
    order表示列的排列顺序，group表示分组导出（比如你这个类需要根据在不同的需求里导出的字段不同，
    那么可以根据group分组导出）
    algin调整列的居左居中居右位置，width设置好列宽美化展示
    ExcelField为字段列名注解 ExcelSheet为sheet表格注解
```
### 导出excel

    /** 导出单sheet的excel文件
     * @param filePath
     * @param collection
     * @param clazz
     * @param group
     */
    public static void exportToFile(String filePath, Collection<?> collection, Class<?> clazz, int group) {
    	ExcelExportUtil.exportSingleSheetToFile(filePath,collection,group);
    } 

    /** 导出多个sheet的excel文件
     * @param filePath
     * @param collectionArr
     * @param clazzArr
     * @param group
     */
    public static void exportMutiToFile(String filePath, Collection<?>[] collectionArr, int[] group) {
    	ExcelExportUtil.exportMutiSheetToFile(filePath,collectionArr,group);
    } 
### 从excel读数据对象

    /** 读取excel文件为Map类型的列表
     * @param file
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
	
    public static List<Map<String, String>> importFromFile(File file) throws IOException, InvalidFormatException {
        return ExcelImportUtil.importFromFile(file);
    }

    /**
     * @param file
     * @param clazz
     * @return 读取excel文件为泛型类型的列表
     * @throws IllegalAccessException
     * @throws IntrospectionException
     * @throws InvalidFormatException
     * @throws IOException
     * @throws InstantiationException
     * @throws InvocationTargetException
     */
    public static <T> List<T> importFromFile(File file, Class<T> clazz) throws IllegalAccessException,
            IntrospectionException, InvalidFormatException, IOException, InstantiationException,
            InvocationTargetException {
        return ExcelImportUtil.importFromFile(file, clazz);
    }
    
### 好用请给STAR
