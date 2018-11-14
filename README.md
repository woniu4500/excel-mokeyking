# excel-mokeyking
	基于注解导入导出excel（可以基于注解定制sheet名称，表头以及对数据日期格式的更加灵活的控制，比如数字金额右靠，文本左靠。）

## Example

### annotation
```java
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
    
```
### 导出excel

    /** 导出单sheet的excel文件
     * @param filePath
     * @param collection
     * @param clazz
     * @param tag
     */
    public static void exportToFile(String filePath, Collection<?> collection, Class<?> clazz, int tag) {
    	ExcelExportUtil.exportSingleSheetToFile(filePath,collection,tag);
    } 

    /** 导出多个sheet的excel文件
     * @param filePath
     * @param collectionArr
     * @param clazzArr
     * @param tag
     */
    public static void exportMutiToFile(String filePath, Collection<?>[] collectionArr, int[] tag) {
    	ExcelExportUtil.exportMutiSheetToFile(filePath,collectionArr,tag);
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
    
