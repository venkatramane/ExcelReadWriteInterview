package stepDefinition;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cucumber.api.java.en.Given;






public class Step_Definition {
	
	@Given("^Excel duplicatefinder timezone difference$")
	public void excel_duplicatefinder_timezone_difference() throws Exception {

		Set<String> duplicate_finder=new HashSet<String>();

		
		
		
		File file=new File("C:\\Users\\VENKATRAMAN\\Desktop\\LexisNexis\\sample_excelfile.xlsx");
		File modified_file=new File("C:\\Users\\VENKATRAMAN\\Desktop\\LexisNexis\\Output_excelfile.xlsx");
		
		FileInputStream fis=new FileInputStream(file);    //Src InputStream
		FileInputStream fis1=new FileInputStream(modified_file);  // TargetFile InputStream
		
		XSSFWorkbook wb=new XSSFWorkbook(fis);      // Workbook
		XSSFWorkbook new_wb=new XSSFWorkbook(fis1);
		
		XSSFSheet sheet=wb.getSheet("Sheet1");    // Src file sheet get Workbook
		XSSFSheet modified_sheet=new_wb.createSheet(); // target file Create sheet
		
		CellStyle origStyle = wb.getCellStyleAt(2); //source file header style copy
		CellStyle newStyle = new_wb.createCellStyle(); //create style
		
		CellStyle origStyle1 = wb.getCellStyleAt(13); // Just values assign in non style
		CellStyle newStyle1 = new_wb.createCellStyle();
		
	//header
		
		for(int i=0;i<sheet.getRow(0).getLastCellNum();i++) {  // sheet -source file sheet
			Row row=sheet.getRow(0);  // 0 constant 1row (header Values)
			Cell cell=row.getCell(i);
			String value=cell.getStringCellValue();
			Row row1;
			if(i==0) row1=modified_sheet.createRow(0);  // modifed shet -- target file sheet cell 0
			else row1=modified_sheet.getRow(0);  // getrow() -- for another cell
			Cell cell1=row1.createCell(i);
			Cell new_cell=modified_sheet.getRow(0).createCell(12);  //  extra column
			new_cell.setCellValue("Time Taken (days)");  // headding
			
			cell1.setCellValue(value);  //
			if(i==9) {
				cell1.setCellValue("Time Ingested (AEST Time)" );  // changed header name
			}
			newStyle.cloneStyleFrom(origStyle);  // source file header style clone
			cell1.setCellStyle(newStyle);    // aditional method while calling it will assign style form old to new change cell style 
			new_cell.setCellStyle(newStyle);  // time taken cell style
			
			
		}
		
		
			
		
		
		
		
	//values
		
		CellStyle cellStyle = new_wb.createCellStyle();     
		CreationHelper createHelper = new_wb.getCreationHelper();  // dateFormater copy - fetch purpose
		cellStyle.setDataFormat(
		    createHelper.createDataFormat().getFormat("dd/MM/yyyy")); // to set date format double to date format
	//	cell = row.createCell(1);
	//	cell.setCellValue(new Date());
	//	cell.setCellStyle(cellStyle);
		
		
		for(int i=1;i<sheet.getLastRowNum();i++) {  //row
			
			Row row1=modified_sheet.createRow(i);   // create row// i=0 or 1 mean create row else getrow
				String date1=null;        // date 1 receive
				String date2=null;        // date 2 send
			
			for(int j=0;j<sheet.getRow(i).getLastCellNum();j++) {  //col
				
			// total number of column cell type string and numeric	
				
				if(j==1||j==7) {            //1 & 7 date  // 2 & 9 Time mention
					Row row=sheet.getRow(i);
					Cell cell=row.getCell(j); 
					CellType type=cell.getCellType();
					double date=cell.getNumericCellValue();  // getting numeric value return type double
					Date dateformat=DateUtil.getJavaDate(date);  // convert double to date
					String value=new SimpleDateFormat("dd-MM-yyyy").format(dateformat);
					if(j==1) {
						date1=value;    // store a if 1 col value here
					}
					else {
						date2=value;   // else 7 store a value in date2
					}
					System.out.println(value);
					row1=modified_sheet.getRow(i);
					Cell cell1=row1.createCell(j);
					cell1.setCellValue(value);
					cell1.setCellStyle(cellStyle);
					
					//find different 
					
					
					
					

				}
	// taking time			
				else if(j==2) {
					Row row=sheet.getRow(i);
					Cell cell=row.getCell(j);
					CellType type=cell.getCellType();
					double date=cell.getNumericCellValue(); //numeric value - time
					Date dateformat=DateUtil.getJavaDate(date);
					String value=new SimpleDateFormat("hh:mm a").format(dateformat); //a- am||Pm , HH-01,02, hh-1
					System.out.println(value);
					row1=modified_sheet.getRow(i);
					Cell cell1=row1.createCell(j);   // fetch the value in string type
					cell1.setCellValue(value);
					cell1.setCellStyle(cellStyle);
//completed third col value
				}
			// add time by 2	
				else if(j==9) {              // mnl to aest
					Row row=sheet.getRow(i);
					Cell cell=row.getCell(j);
					CellType type=cell.getCellType();
					double time=cell.getNumericCellValue();  // get numeric value
					Date time_format=DateUtil.getJavaDate(time);
					String sdf=new SimpleDateFormat("hh:mm:ss a").format(time_format);
					
					SimpleDateFormat format1=new SimpleDateFormat("hh:mm:ss a");
					
					Date time1=format1.parse(sdf);
					long update_time=time1.getTime()+2*1000*60*60; // Australian Eastern Standard Time is 2 hours ahead of Manila, Metro Manila, Philippines
					String outputtime=new SimpleDateFormat("h:mm a").format(update_time);
			
					row1=modified_sheet.getRow(i);
					Cell cell1=row1.createCell(j);
					cell1.setCellValue(outputtime);  //Value feched
					cell1.setCellStyle(cellStyle);

				}
				
				else {
					
					if(j==11||j==12) continue;
					else {
					Row row=sheet.getRow(i);
					Cell cell=row.getCell(j);
					String value=cell.getStringCellValue();
					row1=modified_sheet.getRow(i);   //get String value fetch value
					Cell cell1=row1.createCell(j);
					cell1.setCellValue(value);
					
					}
				}
				
				
				
				
			}
		
	// different finder and adding to that value minus value		
			SimpleDateFormat format=new SimpleDateFormat("dd-MM-yyyy");

			Date d1=format.parse(date1);
			Date d2=format.parse(date2);  // 
			long diff=d1.getTime()-d2.getTime();
			long days_diff=diff/(1000*60*60*24);
			modified_sheet.getRow(i).createCell(12).setCellValue(days_diff);
		}
		
		
		
		
		
		
		
		
		
		
		
		//duplicate remover
		//VCI col remove duplicate value using set collection //2
		for(int i=0;i<modified_sheet.getLastRowNum();i++) {
			if(!duplicate_finder.add(modified_sheet.getRow(i).getCell(0).getStringCellValue())) {
				modified_sheet.removeRow(modified_sheet.getRow(i));
			}
			
			
			FileOutputStream fos=new FileOutputStream(modified_file);
			new_wb.write(fos);
		
	/*	File modified_file_re=new File("E:\\venkat project\\sample_excelfile10.xlsx");
		FileInputStream fis1_re=new FileInputStream(modified_file_re);
		XSSFWorkbook new_wb_re=new XSSFWorkbook(fis1_re);
		XSSFSheet modified_sheet_re=new_wb_re.getSheet("LN");
		
		for(int i=0;i<modified_sheet_re.getLastRowNum();i++) {
			if(!duplicate_finder.add(modified_sheet_re.getRow(i).getCell(0).getStringCellValue())) {
				modified_sheet_re.removeRow(modified_sheet_re.getRow(i));
			}*/
		}
		


	}
}