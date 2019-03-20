import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Testexport {
	
	public static void main(String[] args) {
		List<Book> mybook = new ArrayList<Book>();
		Book bk1 = new Book();
		bk1.setAuthor("张三");
		bk1.setDate(new Date());
		bk1.setName("儒林外史");
		bk1.setPrice(50);
		bk1.setType("文学");
		bk1.setUsed(true);
		mybook.add(bk1);
		Book bk2 = new Book();
		bk2.setAuthor("李四");
		bk2.setDate(new Date());
		bk2.setName("儒林外史");
		bk2.setPrice(50);
		bk2.setType("文学");
		bk2.setUsed(true);
		mybook.add(bk2);
		Book bk3 = new Book();
		bk3.setAuthor("王五");
		bk3.setDate(new Date());
		bk3.setName("儒林外史");
		bk3.setPrice(50);
		bk3.setType("文学");
		bk3.setUsed(true);
		mybook.add(bk3);
		
		String[] headers = new String[]{"书名","作者","价格","类型","日期","是否使用"};
		BaseExcelUtil<Book> util = new BaseExcelUtil<Book>(mybook, "测试一下", "测试标题", headers, "E://", null);
		util.export("使用中,未使用");
	}
}
