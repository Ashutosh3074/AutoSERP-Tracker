package SEOKeywordPresencce.Repository;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

public class excelSheetUtility {

	static int rowcount = 1;
	static String sheetName = "IDFC";
	static HSSFWorkbook wb = new HSSFWorkbook();
	static HSSFSheet sheet = wb.createSheet(sheetName);
	static String writecode = "geo0.xls";

	// code is to write url value in excel sheet..
	public static void failcode(String clent_name, String keyword,  String microsite, int nullorganic, String localities, String localitiesPos1,String localitiesPos2,String localitiesPos3,String localitiesPos4,String status) throws IOException {
		HSSFRow row = sheet.createRow(rowcount);
		// iterating c number of columns
		int cellUrl = 0;
		HSSFCell cellurl = row.createCell(cellUrl);
		cellurl.setCellValue(clent_name);
		int cellUrlone = 1;
		HSSFCell cellurlone = row.createCell(cellUrlone);
		cellurlone.setCellValue(keyword);
		int cellUrltwo = 2;
		HSSFCell cellurlthree = row.createCell(cellUrltwo);
		cellurlthree.setCellValue(microsite);
		int cellUrlfour = 3;
		HSSFCell cellurlfive = row.createCell(cellUrlfour);
		cellurlfive.setCellValue(nullorganic);
		int cellUrlsix = 4;
		HSSFCell cellurlsix= row.createCell(cellUrlsix);
		cellurlsix.setCellValue(localities);
		int cellUrl7 =5 ;
		HSSFCell cellurl7= row.createCell(cellUrl7);
		cellurl7.setCellValue(localitiesPos1);
		int cellUrl8 =6 ;
		HSSFCell cellurl8= row.createCell(cellUrl8);
		cellurl8.setCellValue(localitiesPos2);
		int cellUrl9 =7 ;
		HSSFCell cellurl9= row.createCell(cellUrl9);
		cellurl9.setCellValue(localitiesPos3);
		
		int cellUrl10 =8 ;
		HSSFCell cellurl10= row.createCell(cellUrl10);
		cellurl10.setCellValue(localitiesPos4);
		int cellUrl11 =9 ;
		HSSFCell cellurl11= row.createCell(cellUrl11);
		cellurl11.setCellValue(status);
		
		
		
		FileOutputStream fileOut1 = new FileOutputStream(writecode);
		wb.write(fileOut1);
		rowcount++;
	}

	// code to write book a test drive..
	public static void passcode(String clent_name, String keyword,  String microsite, int nullorganic, String localities, String localitiesPos1,String localitiesPos2,String localitiesPos3,String localitiesPos4,String status) throws IOException {
		HSSFRow row = sheet.createRow(rowcount);
		// iterating c number of columns
		int cellUrl = 0;
		HSSFCell cellurl = row.createCell(cellUrl);
		cellurl.setCellValue(clent_name);
		int cellUrlone = 1;
		HSSFCell cellurlone = row.createCell(cellUrlone);
		cellurlone.setCellValue(keyword);
		int cellUrltwo = 2;
		HSSFCell cellurlthree = row.createCell(cellUrltwo);
		cellurlthree.setCellValue(microsite);
		int cellUrlfour = 3;
		HSSFCell cellurlfive = row.createCell(cellUrlfour);
		cellurlfive.setCellValue(nullorganic);
		int cellUrlsix = 4;
		HSSFCell cellurlsix= row.createCell(cellUrlsix);
		cellurlsix.setCellValue(localities);
		int cellUrl7 =5 ;
		HSSFCell cellurl7= row.createCell(cellUrl7);
		cellurl7.setCellValue(localitiesPos1);
		int cellUrl8 =6 ;
		HSSFCell cellurl8= row.createCell(cellUrl8);
		cellurl8.setCellValue(localitiesPos2);
		int cellUrl9 =7 ;
		HSSFCell cellurl9= row.createCell(cellUrl9);
		cellurl9.setCellValue(localitiesPos3);
		
		int cellUrl10 =8 ;
		HSSFCell cellurl10= row.createCell(cellUrl10);
		cellurl10.setCellValue(localitiesPos4);
		int cellUrl11 =9 ;
		HSSFCell cellurl11= row.createCell(cellUrl11);
		cellurl11.setCellValue(status);
		FileOutputStream fileOut1 = new FileOutputStream(writecode);
		wb.write(fileOut1);
		rowcount++;
	}

	// This code is to write a row header values...
	public static void headerValues() throws IOException {
		int newrow = 0;
		HSSFRow row1 = sheet.createRow(newrow);
//		excelSheetUtility2.Passcode(client_Name, finalykeyword, microsite, organicpasscount,nearbybuffer.toString() , position1nearbybuffer.toString(),position2nearbybuffer.toString(),position3nearbybuffer.toString(),position4nearbybuffer.toString());

		ArrayList<String> names = new ArrayList<String>(Arrays.asList("client Name","Search Keyword","Website Url On Organic Serach","Organic URL Position","NearBy Localities in Website","Postion at 1 - locality Website","Postion at 2 - locality Website","Postion at 3 - locality Website","Postion More than 3 - locality Website","Status"));
		int c = 0;
		for (String cellName : names) {
			HSSFCell cell = row1.createCell(c++);
			cell.setCellValue(cellName);
		}
		FileOutputStream fileOut3 = new FileOutputStream(writecode);
		wb.write(fileOut3);
	}

	
}


