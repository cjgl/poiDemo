package cn.cjgl;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Calendar;
import java.util.Iterator;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

public class Hssf {
	public void createCells() throws IOException {
		HSSFWorkbook wb = new HSSFWorkbook(); // 建立新HSSFWorkbook对象

		HSSFCellStyle cellstyle = wb.createCellStyle();
		cellstyle.setAlignment(HSSFCellStyle.ALIGN_CENTER_SELECTION);
		HSSFCreationHelper createHelper = wb.getCreationHelper();

		HSSFSheet sheet = wb.createSheet("new sheet"); // 建立新的sheet对象
		HSSFRow row = sheet.createRow(0);
		// 在sheet里创建一行，参数为行号（第一行，此处可想象成数组）
		HSSFCell cell = row.createCell(0);
		cell.setCellStyle(cellstyle);
		// 在row里建立新cell（单元格），参数为列号（第一列）
		cell.setCellValue(1); // 设置cell的整数类型的值
		row.createCell(1).setCellValue(1.2); // 设置cell浮点类型的值
		row.createCell(2).setCellValue("test"); // 设置cell字符类型的值
		row.createCell(3).setCellValue(true); // 设置cell布尔类型的值
		HSSFCellStyle cellStyle = wb.createCellStyle(); // 建立新的cell样式
		// cellStyle.setDataFormat(HSSFDataFormat.
		// getBuiltinFormat("m/d/yy h:mm"));
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
				"yyyy-m-d h:mm"));
		// 设置cell样式为定制的日期格式
		HSSFCell dCell = row.createCell(4);
		// dCell.setCellValue(new Date()); //设置cell为日期类型的值
		dCell.setCellValue(Calendar.getInstance());
		dCell.setCellStyle(cellStyle); // 设置该cell日期的显示格式
		HSSFCell csCell = row.createCell(5);
		// csCell.setEncoding(HSSFCell.ENCODING_UTF_16);
		// 设置cell编码解决中文高位字节截断
		csCell.setCellValue("中文测试_Chinese Words Test"); // 设置中西文结合字符串
		row.createCell(6).setCellType(HSSFCell.CELL_TYPE_ERROR);
		// 建立错误cell

		sheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 4));

		sheet.createFreezePane(0, 2);

		FileOutputStream fileOut = new FileOutputStream("f:/workbook.xls");
		wb.write(fileOut);
		fileOut.close();
	}
	
	public void dataValidation() throws IOException {
		HSSFWorkbook wb = new HSSFWorkbook(); 
		Sheet sheet = wb.createSheet("Data Validation"); 
		CellRangeAddressList addressList = new CellRangeAddressList(0, 1, 0, 1); 
		DVConstraint dvConstraint = DVConstraint.createExplicitListConstraint( 
		      new String[]{"10", "20", "30"}); 
		HSSFDataValidation dataValidation = new HSSFDataValidation(addressList, dvConstraint); 
		dataValidation.setSuppressDropDownArrow(false);//false下拉 
		//sheet.addValidationData(dataValidation); 
		  
		dataValidation.setErrorStyle(HSSFDataValidation.ErrorStyle.STOP); 
		dataValidation.createErrorBox("Box Title", "Message Text");
		dataValidation.createPromptBox("Title", "Message Text"); 
		dataValidation.setShowPromptBox(true);
		 
		sheet.addValidationData(dataValidation);
		
		FileOutputStream out = new FileOutputStream("f://workbook.xls"); 
		wb.write(out); 
		out.close();
	}
	
	public void forEach() throws IOException{
		InputStream inp = new FileInputStream("f://workbook.xls");
		HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
		Sheet sheet1 = wb.getSheetAt(0);
		//for (HSSFRow row : sheet1) { 
		for (Iterator<Row> rit = sheet1.rowIterator(); rit.hasNext();){
			Row row = rit.next();
		  //for (HSSFCell cell : row) {
		    for (Iterator<Cell> cit = row.cellIterator(); cit.hasNext(); ) {
		    	Cell cell = cit.next();
		    	CellReference cellRef = new 
		    		CellReference(row.getRowNum(), cell.getColumnIndex()); 
		    	System.out.print(cellRef.formatAsString()); 
		    	System.out.print(" - "); 
		     
		    	switch(cell.getCellType()) {
		    		case Cell.CELL_TYPE_STRING:
		    			System.out.println(cell.getRichStringCellValue().getString());
		    			break;
		    		case Cell.CELL_TYPE_NUMERIC:
		    			if(DateUtil.isCellDateFormatted(cell)) {
		    				System.out.println(cell.getDateCellValue()); 
		    			} else {
		    				System.out.println(cell.getNumericCellValue());
		    			} 
		    			break; 
		    		case Cell.CELL_TYPE_BOOLEAN:
		    			System.out.println(cell.getBooleanCellValue());
		    			break;
		    		case Cell.CELL_TYPE_FORMULA:
		    			System.out.println(cell.getCellFormula());
		    			break;
		    		default:
		    			System.out.println();
		    	}
		    }
		}
	}
	
	public void getText() throws IOException{
		InputStream inp = new FileInputStream("f://workbook.xls"); 
		HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
		ExcelExtractor extractor = new ExcelExtractor(wb); 
		extractor.setFormulasNotResults(true); 
		extractor.setIncludeSheetNames(false); 
		String text = extractor.getText(); 
		System.out.println(text);
	}
	
	public void groupRow() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook(); 
		Sheet sheet1 = wb.createSheet("new sheet"); 
		 
		sheet1.groupRow( 5, 14 ); 
		//sheet1.groupRow( 7, 14 ); 
		//sheet1.groupRow( 16, 19 ); 
		/*
		sheet1.groupColumn( (short)4, (short)7 ); 
		sheet1.groupColumn( (short)9, (short)12 ); 
		sheet1.groupColumn( (short)10, (short)11 );
		*/
		sheet1.setRowGroupCollapsed( 7, true );
		 
		FileOutputStream fileOut = new FileOutputStream("f://workbook.xls");
		wb.write(fileOut); 
		fileOut.close();
	}
	
	public void mergedRegion() throws IOException{
		 Workbook wb = new HSSFWorkbook(); 
		 Sheet sheet = wb.createSheet("new sheet"); 
		 
		 Row row = sheet.createRow(1); 
		 Cell cell = row.createCell(1); 
		 cell.setCellValue("This is a test of merging"); 
		 
		 sheet.addMergedRegion(new CellRangeAddress( 
		            1, //first row (0-based) 
		            2, //last row  (0-based) 
		            1, //first column (0-based) 
		            2  //last column  (0-based) 
		 )); 
		 
		//将输出流写入一个文件 
		FileOutputStream fileOut = new FileOutputStream("f://workbook.xls"); 
		wb.write(fileOut); 
		fileOut.close();
	}
	
	public void setBackgroundColor() throws IOException{
		Workbook wb = new HSSFWorkbook(); 
		Sheet sheet = wb.createSheet("new sheet"); 
		//创建一列，在其中加入多个单元格，列索引号从0 开始，单元格的索引号也是从 0 
		//开始 
		Row row = sheet.createRow(1); 
		// 浅绿色背景色 
		CellStyle style = wb.createCellStyle(); 
		style.setFillBackgroundColor(IndexedColors.AQUA.getIndex()); 
		style.setFillPattern(CellStyle.BIG_SPOTS); 
		Cell cell = row.createCell(1); 
		cell.setCellValue("X"); 
		cell.setCellStyle(style);
		//橙色前景色。前景色：前景色是指正在使用的填充颜色，而非字体颜色 
		style = wb.createCellStyle(); 
		style.setFillForegroundColor(IndexedColors.ORANGE.getIndex()); 
		style.setFillPattern(CellStyle.SOLID_FOREGROUND); 
		cell = row.createCell(2); 
		cell.setCellValue("X");
		cell.setCellStyle(style); 
		 
		//将输出流写入一个文件 
		FileOutputStream fileOut = new FileOutputStream("f://workbook.xls"); 
		wb.write(fileOut); 
		fileOut.close();
	}
	
	public void setComment() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook(); //or new HSSFWorkbook(); 
		 
	    CreationHelper factory = wb.getCreationHelper(); 
	 
	    Sheet sheet = wb.createSheet(); 
	     
	    Cell cell = sheet.createRow(3).createCell(5); 
	    cell.setCellValue("F4"); 
	     
	    Drawing drawing = sheet.createDrawingPatriarch(); 
	 
	    ClientAnchor anchor = factory.createClientAnchor(); 
	    Comment comment = drawing.createCellComment(anchor); 
	    RichTextString str = factory.createRichTextString("Hello, World!"); 
	    comment.setString(str); 
	    comment.setAuthor("Apache POI");
	    //将注释注册到单元格 
	    cell.setCellComment(comment); 
	    String fname = "f://workbook.xls"; 
	    //if(wb instanceof XSSFWorkbook) fname += "x"; 
	    FileOutputStream out = new FileOutputStream(fname); 
	    wb.write(out); 
	    out.close();
	}
	
	public void setFooter() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook(); 
		Sheet sheet = wb.createSheet("format sheet"); 
		Footer footer = sheet.getFooter(); 
		 
		//footer.setRight( "Page " + HSSFFooter.page() + " of " + 
		//HSSFFooter.numPages() );
		footer.setCenter("Page " + HSSFFooter.page() + " of " + HSSFFooter.numPages());
		 
		 
		//为电子表格创建多行多列 
		for(int i = 0; i < 128; i++){
			Row row = sheet.createRow(i);
			Cell cell = row.createCell(1);
			cell.setCellValue("row:"+i);
		}

		sheet.setZoom(150, 100);
		sheet.createFreezePane(0, 1);
		wb.setRepeatingRowsAndColumns(0, -1, -1, 0, 0);
		FileOutputStream fileOut = new FileOutputStream("f://workbook.xls"); 
		wb.write(fileOut); 
		fileOut.close();
	}
	
	public void testFont() throws IOException{
		Workbook wb = new HSSFWorkbook(); 
		Sheet sheet = wb.createSheet("new sheet"); 
		 
		//创建一列，在其中加入多个单元格，列索引号从0 开始，单元格的索引号也是从 0 
		//开始 
		//Row row = sheet.createRow(1); 
		 
		// 创建一个字体并修改它的样式.
		/*
		Font font = wb.createFont(); 
		font.setFontHeightInPoints((short)10); 
		//font.setFontName("Courier New");
		font.setFontName("仿宋"); 
		font.setItalic(true); 
		font.setStrikeout(true);
		*/
		//font.setColor(IndexedColors.ROSE.getIndex());
		//System.out.println(IndexedColors.ROSE.getIndex());
		//字体在样式（style）中载入才能使用，因此创建一个style来使用该字体 
		//CellStyle style = wb.createCellStyle(); 
		//style.setFont(font); 
		//创建一个单元格，并在其中加入内容.
		for(int i = 0; i < 50; i++){
			Font font = wb.createFont(); 
			font.setFontHeightInPoints((short)10); 
			//font.setFontName("Courier New");
			font.setFontName("仿宋"); 
			font.setItalic(true); 
			font.setStrikeout(true);
			font.setColor((short)i);
			CellStyle style = wb.createCellStyle();
			style.setFont(font);
			Row row = sheet.createRow(i); 
			Cell cell = row.createCell(1); 
			cell.setCellValue("This is a \ntest of fonts"); 
			cell.setCellStyle(style);
		}
		 sheet.autoSizeColumn((short)1); 
		//将输出流写入一个文件 
		FileOutputStream fileOut = new FileOutputStream("f://workbook.xls"); 
		wb.write(fileOut); 
		fileOut.close();
	}
	
	public void testHyperlink() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook(); //or new HSSFWorkbook(); 
		CreationHelper createHelper = wb.getCreationHelper(); 
		 
		//设置单元格格式为超级链接 
		//默认情况下超级链接为下划线、蓝色字体 
		CellStyle hlink_style = wb.createCellStyle(); 
		Font hlink_font = wb.createFont(); 
		hlink_font.setUnderline(Font.U_SINGLE); 
		hlink_font.setColor(IndexedColors.BLUE.getIndex()); 
		hlink_style.setFont(hlink_font); 
		 
		Cell cell; 
		Sheet sheet = wb.createSheet("Hyperlinks"); 
		//URL 
		cell = sheet.createRow(0).createCell((short)0); 
		cell.setCellValue("URL Link"); 
		 
		Hyperlink link = createHelper.createHyperlink(Hyperlink.LINK_URL); 
		link.setAddress("http://www.baidu.com"); 
		cell.setHyperlink(link); 
		cell.setCellStyle(hlink_style); 
		 
		//关联到当前目录的文件 
		cell = sheet.createRow(1).createCell((short)0); 
		cell.setCellValue("File Link"); 
		link = createHelper.createHyperlink(Hyperlink.LINK_FILE); 
		link.setAddress("fund.xls"); 
		cell.setHyperlink(link); 
		cell.setCellStyle(hlink_style); 
		    
		//e-mail 关联 
		cell = sheet.createRow(2).createCell((short)0); 
		cell.setCellValue("Email Link"); 
		link = createHelper.createHyperlink(Hyperlink.LINK_EMAIL); 
		//注意：如果邮件主题中存在空格，请保证是按照URL 格式书写的 
		link.setAddress("mailto:poi@apache.org?subject=Hyperlinks"); 
		cell.setHyperlink(link); 
		cell.setCellStyle(hlink_style); 
		    
		//关联到工作簿中的位置 
		    
		//创建一个目标sheet页和目标单元格 
		Sheet sheet2 = wb.createSheet("Target Sheet"); 
		sheet2.createRow(0).createCell((short)0).setCellValue("Target Cell"); 
		 
		cell = sheet.createRow(3).createCell((short)0); 
		cell.setCellValue("Worksheet Link"); 
		Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT); 
		link2.setAddress("'Target Sheet'!A1"); 
		cell.setHyperlink(link2); 
		cell.setCellStyle(hlink_style); 
		 
		FileOutputStream out = new FileOutputStream("f://workbook.xls"); 
		wb.write(out); 
		out.close();
	}
	
	public void hssfAlign() throws IOException{
		Workbook wb = new HSSFWorkbook(); 
		Sheet sheet = wb.createSheet("new sheet"); 
			 
		//创建一列，在其中加入多个单元格，列索引号从0 开始，单元格的索引号也是从 0 
		//开始. 
		Row row = sheet.createRow(1); 
		//创建一个单元格，并在其中加入内容. 
		Cell cell = row.createCell(1); 
		cell.setCellValue(100); 
				 
		//设置单元格边框为四周环绕. 
		CellStyle style = wb.createCellStyle(); 
		style.setBorderBottom(CellStyle.BORDER_THICK); 
		style.setBottomBorderColor(IndexedColors.RED.getIndex()); 
		style.setBorderLeft(CellStyle.BORDER_THIN); 
		style.setLeftBorderColor(IndexedColors.GREEN.getIndex()); 
		style.setBorderRight(CellStyle.BORDER_THIN); 
		style.setRightBorderColor(IndexedColors.BLUE.getIndex()); 
		style.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED); 
		style.setTopBorderColor(IndexedColors.BLACK.getIndex()); 
		cell.setCellStyle(style); 
				 
		//将输出流写入一个文件 
		FileOutputStream fileOut = new FileOutputStream("f://workbook.xls"); 
		wb.write(fileOut); 
		fileOut.close();
	}
}
