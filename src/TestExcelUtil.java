import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 继承基础导出工具修改导出格式
 * @author:  Asami
 * @ClassName:  TestExcelUtil
 * @date:  2018年12月4日 下午2:22:42
 */
public class TestExcelUtil<T> extends BaseExcelUtil<T>{

	public TestExcelUtil(List<T> list, String excelName, String excelTitle, String[] headers, String outputUrl, String pattern) {
		super(list, excelName, excelTitle, headers, outputUrl, pattern);
	}

	@Override
	protected HSSFCellStyle setTitleStyle() {
		// 生成一个样式
        HSSFCellStyle titleStyle = workbook.createCellStyle();
        
        Font fontStyle = workbook.createFont(); // 字体样式
		fontStyle.setBold(true); // 加粗
		fontStyle.setFontName("黑体"); // 字体
		fontStyle.setFontHeightInPoints((short) 13); // 大小
		titleStyle.setFont(fontStyle);
        return titleStyle;
	}
	
	@Override
	protected void setTitle(HSSFCellStyle style) {
		HSSFRow titleRow = sheet.createRow(index);
        // 下标从0开始 起始行号，终止行号， 起始列号，终止列号
        CellRangeAddress region = new CellRangeAddress(0, 1, 0, 3);
        sheet.addMergedRegion(region);
        HSSFCell cell = titleRow.createCell(0);
        if(style!= null) {
        	cell.setCellStyle(style);
        }
        cell.setCellValue(excelTitle);
        index = index + 2 ;
	}
	
	@Override
	protected void setDatas(HSSFCellStyle style, String trueFalse) {
		// 遍历集合数据，产生数据行
        Iterator<Book> it = (Iterator<Book>) data.iterator();
        HSSFRow row;
        index -- ;
        while (it.hasNext()) {
        	index ++;
        	row = sheet.createRow(index);
        	Book book = it.next();
        	row.createCell(0).setCellValue(book.getAuthor());
        	row.createCell(1).setCellValue(book.getName());
        	row.createCell(2).setCellValue(book.getType());
        	row.createCell(3).setCellValue(book.getPrice());
        	
        }
	}
	
	public static void main(String[] args) {
		String[] headers =
    		{ "学号", "姓名", "年龄", "性别"};
    	List<Book> dataset = new ArrayList<Book>();
    	Book bk1 = new Book();
		bk1.setAuthor("张三");
		bk1.setDate(new Date());
		bk1.setName("儒林外史");
		bk1.setPrice(50);
		bk1.setType("文学");
		bk1.setUsed(true);
		dataset.add(bk1);
		Book bk2 = new Book();
		bk2.setAuthor("李四");
		bk2.setDate(new Date());
		bk2.setName("儒林外史");
		bk2.setPrice(50);
		bk2.setType("文学");
		bk2.setUsed(true);
		dataset.add(bk2);
		Book bk3 = new Book();
		bk3.setAuthor("王五");
		bk3.setDate(new Date());
		bk3.setName("儒林外史");
		bk3.setPrice(50);
		bk3.setType("文学");
		bk3.setUsed(true);
		dataset.add(bk3);
        // 测试学生
		TestExcelUtil<Book> util = new TestExcelUtil<Book>(dataset, "测试学生表", "学生标题", headers, "F://", null);
    	util.excelTitle = "修改学生表";
        util.export(null);
	}
}
