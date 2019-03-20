import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;


/**
 * 利用POI导出Excel文件
 * @author:  Asami
 * @ClassName:  ExportExcel
 * @date:  2018年12月4日 上午10:36:30
 */
public class BaseExcelUtil<T> {
	
	protected String excelTitle;//excel抬头
	
	protected List<T> data;//数据
	
	protected String outputUrl;//输出路径
	
	protected HSSFWorkbook workbook;//工作簿
	
	protected HSSFSheet sheet;//表格
	
	protected String pattern;//导出时间格式
	
	protected String[] headers;
	
	protected HSSFPatriarch patriarch;
	
	protected int index = 0 ;//当前生成的行数
	
	
	protected HSSFCellStyle setTitleStyle() {
		// 生成一个样式
        HSSFCellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        return titleStyle;
	}
	

	protected HSSFCellStyle setDataStyle() {
		// 生成并设置另一个样式
        HSSFCellStyle style2 = workbook.createCellStyle();
        style2.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        
        return style2;
	}
	
	
	protected HSSFCellStyle setHeadStyle() {
		// 生成一个样式
        HSSFCellStyle HeadStyle = workbook.createCellStyle();
        // 设置这些样式
        HeadStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
        HeadStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        HeadStyle.setBorderBottom(BorderStyle.THIN);
        HeadStyle.setBorderLeft(BorderStyle.THIN);
        HeadStyle.setBorderRight(BorderStyle.THIN);
        HeadStyle.setBorderTop(BorderStyle.THIN);
        HeadStyle.setAlignment(HorizontalAlignment.CENTER);
        HeadStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        
        // 生成一个字体
 		Font fontStyle = workbook.createFont(); // 字体样式
		fontStyle.setBold(true); // 加粗
		fontStyle.setFontName("黑体"); // 字体
		fontStyle.setFontHeightInPoints((short) 11); // 大小
		// 将字体样式添加到单元格样式中 
		HeadStyle.setFont(fontStyle);
		return HeadStyle;
	}

	
	public BaseExcelUtil(List<T> list,String excelName,String excelTitle,String[] headers,String outputUrl,String pattern) {
		if(pattern == null) {
			this.pattern = "yyyy-MM-dd";
		}
		this.data = list;
		this.headers = headers;
		this.excelTitle = excelTitle;
		this.outputUrl = outputUrl + excelName + ".xls";
		// 声明一个工作薄
        workbook = new HSSFWorkbook();
        // 生成一个表格
        sheet = workbook.createSheet(excelTitle);
        // 设置表格默认列宽度为15个字节
        sheet.setDefaultColumnWidth((short) 15);
        // 声明一个画图的顶级管理器
        patriarch= sheet.createDrawingPatriarch();
	}
	
	
	/**
	 * 执行导出任务,传参为A,B形式 A为真值显示内容B为假值显示内容,如不传默认真假
	 * @param trueFalse
	 */
	public void export(String trueFalse) {
		
			
		setTitle(setTitleStyle());
		
        setHeagers(setHeadStyle());
        
        setDatas(setDataStyle(),trueFalse);
        
        wirteExcel();
		
	}
	/**
	 * 输出Excel文件
	 */
	protected void wirteExcel() {
		//输出
        try {
        	OutputStream out = new FileOutputStream(outputUrl);
			workbook.write(out);
			out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
	}

	/**
	 * 设置标题行
	 */
	protected void setTitle(HSSFCellStyle style) {
		HSSFRow titleRow = sheet.createRow(index);
        // 下标从0开始 起始行号，终止行号， 起始列号，终止列号
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, headers.length - 1);
        sheet.addMergedRegion(region);
        HSSFCell cell = titleRow.createCell(0);
        if(style!= null) {
        	cell.setCellStyle(style);
        }
        cell.setCellValue(excelTitle);
        index ++ ;
	}
	
	/**
	 * 设置表格列名
	 */
	protected void setHeagers(HSSFCellStyle style) {
		// 产生表格标题行
        HSSFRow row = sheet.createRow(index);
        for (short i = 0; i < headers.length ; i++)
        {
            HSSFCell cell = row.createCell(i);
            if(style != null) {
            	cell.setCellStyle(style);
            }
            HSSFRichTextString text = new HSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
	}
	
	/**
	 * 添加注释
	 * dx1 dy1 起始单元格中的x,y坐标.
	 * dx2 dy2 结束单元格中的x,y坐标
     * col1,row1 指定起始的单元格，下标从0开始
	 * col2,row2 指定结束的单元格 ，下标从0开始
	 */
	protected void addComment(String author, String text,int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2) {
		// 定义注释的大小和位置,详见文档
	    HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(dx1,dy1, dx2, dy2,col1, row1,col2, row2));    
	    // 设置注释内容
	    comment.setString(new HSSFRichTextString(text));
	    // 设置注释作者，当鼠标移动到单元格上是可以在状态栏中看到该内容.
	    comment.setAuthor(author);
	}
	
	
	
	/**
	 * 设置数据行
	 * @param style
	 */
	protected void setDatas(HSSFCellStyle style,String trueFalse) {
		// 遍历集合数据，产生数据行
        Iterator<T> it = data.iterator();
        String[] trueFalses = new String[]{"是","否"};
        if(trueFalse!= null) {
        	trueFalses = trueFalse.split(",");
        }
        HSSFRow row;
        while (it.hasNext()) {
            index ++ ;
            row = sheet.createRow(index);
            T t = (T) it.next();
            // 利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
            Field[] fields = t.getClass().getDeclaredFields();
            try {
				for (short i = 0; i < fields.length; i++) {
				    HSSFCell cell = row.createCell(i);
				    if(style!= null) {
				    	cell.setCellStyle(style);
				    }
				    Field field = fields[i];
				    String fieldName = field.getName();
				    String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
				    Class tCls = t.getClass();
				    Method getMethod = tCls.getMethod(getMethodName,new Class[]{});
				    Object value = getMethod.invoke(t, new Object[]{});
				    // 判断值的类型后进行强制类型转换
				    String textValue = null;
				    if (value instanceof Boolean){
				        boolean bValue = (Boolean) value;
				        textValue = trueFalses[0];
				        if (!bValue)
				        {
				            textValue = trueFalses[1];
				        }
				    } else if (value instanceof Date) {
				        Date date = (Date) value;
				        SimpleDateFormat sdf = new SimpleDateFormat(pattern);
				        textValue = sdf.format(date);
				    } else if (value instanceof byte[]) {
				        // 有图片时，设置行高为60px;
				        row.setHeightInPoints(60);
				        // 设置图片所在列宽度为80px,注意这里单位的一个换算
				        sheet.setColumnWidth(i, (short) (35.7 * 80));
				        // sheet.autoSizeColumn(i);
				        byte[] bsValue = (byte[]) value;
				        HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0,
				                1023, 255, (short) 6, index, (short) 6, index);
				        anchor.setAnchorType(ClientAnchor.AnchorType.byId(2));
				        patriarch.createPicture(anchor, workbook.addPicture(
				                bsValue, HSSFWorkbook.PICTURE_TYPE_JPEG));
				    } else {
				        // 其它数据类型都当作字符串简单处理
				        textValue = value.toString();
				    }
				    // 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
				    if (textValue != null) {
				        Pattern p = Pattern.compile("^//d+(//.//d+)?$");
				        Matcher matcher = p.matcher(textValue);
				        if (matcher.matches()) {// 是数字当作double处理
				            cell.setCellValue(Double.parseDouble(textValue));
				        }
				        else {
				            HSSFRichTextString richString = new HSSFRichTextString(textValue);
				            HSSFFont font3 = workbook.createFont();
				            font3.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
				            richString.applyFont(font3);
				            cell.setCellValue(richString);
				        }
				    }
				}
			} catch (NumberFormatException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (NoSuchMethodException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SecurityException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (InvocationTargetException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
        }
	}
	
}