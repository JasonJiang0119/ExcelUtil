# ExcelUtil
ExcelUtil

1.本Excel基于apache poi 4.1.2框架(最新版本的POI框架)
分别引入两个框架Jar包
(1).poi包 
groupId->org.apache.poi
artifactId->poi
version->4.1.2
(2).poi-ooxml
groupId->org.apache.poi
artifactId->poi-ooxml
version->4.1.2

2.ExcelUtil使用规则
(1).本ExcelUtil导出Excel 方法为export，采用SXSSFWorkbook，支持大批量数据导出，入参参数为
filenNamePrefix：文件名前缀
rowName：Excel导出文件表头
data：泛型实体对象
sheetName：Excel Sheet名
response：HttpServletResponse用于输出文件流
说明:该export方法适用于list平铺模板，如data对象中某个字段类型为List，该Export方法会将此List对象与其他字段类型平铺成一行数据

(2).导入读取数据，方法名为：readExcel为重载，区别在于其中一个方法入参为MultipartFile，一个为File，两者都需传入需要转换的pojo class，转换原理为：根据传入pojo class ，通过getDeclaredFields()方法获取该对象的所有字段，根据poi框架遍历数据，将该pojo class字段名作为Key值，poi框架获取的表格值为Value存入Map中，利用阿里fastJson反序列化对象
注意：由于是按序读取，反序列化对象定义顺序需与导入Excel模板保持一致
