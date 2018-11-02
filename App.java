import java.io.File; 
import java.io.FileInputStream; 
import java.io.FileNotFoundException; 
import java.io.FileOutputStream; 
import java.io.IOException; 
import java.sql.Date;
import java.util.ArrayList;
import java.util.HashMap; 
import java.util.Iterator;
import java.util.List;
import java.util.Map; 
import java.util.Set; 
import org.apache.poi.ss.usermodel.Cell; 
import org.apache.poi.ss.usermodel.Row; 
import org.apache.poi.xssf.usermodel.XSSFSheet; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 
 
public class App 
{ 
	public static void main(String[] args) 
	{ 
		try 
		{ 
		File excelIn = new File("C:\\Users\\Balaji\\Desktop\\Clearing house Help\\OldFormat.xlsx");
		File excelOut = new File("C:\\Users\\Balaji\\Desktop\\Clearing house Help\\NewFormat.xlsx");
		FileInputStream fis = new FileInputStream(excelIn); 
		XSSFWorkbook book = new XSSFWorkbook(fis); 
		XSSFSheet sheet = book.getSheetAt(0); 
		Iterator<Row> itr = sheet.iterator();
		Map<Integer, ArrayList<String>> sheetData =  new HashMap<Integer, ArrayList<String>>(); 
		ArrayList<String[]> tableData = new ArrayList<String[]>();
		boolean tab1Flag= false;
		int tab1Data=0;
		boolean tab2Flag = false;
		int tab2Data=0;
		boolean tab3Flag = false;
		int tab3Data=0;
		boolean tab4Flag = false;
		int tab4Data=0;
		boolean tab5Flag = false;
		int tab5Data=0;
		boolean tab6Flag = false;
		int tab6Data=0;
		
		int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();
		for(int i=0; i<rowCount+1; i++){
			ArrayList<String> tableRow = new ArrayList<String>();
			Row row = sheet.getRow(i);
			if(row != null) {
			for(int j=0; j<row.getLastCellNum(); j++){
	              Cell cell = row.getCell(j);
	              tableRow.add(cell.getStringCellValue());
	            }
			sheetData.put(i, tableRow);
			
			}
		}
		
		
		// writing data into XLSX file 
		XSSFWorkbook outBook = new XSSFWorkbook();
		XSSFSheet tab1Sheet = null;
		XSSFSheet tab2Sheet = null;
		XSSFSheet tab3Sheet = null;
		XSSFSheet tab4Sheet = null;
		XSSFSheet tab5Sheet = null;
		XSSFSheet tab6Sheet = null;
		Set<Integer> newRows = sheetData.keySet(); 
		int rownum = 0;
		int cellnum = 0;
		Row tab1SheetRow = null;
		Row tab2SheetRow = null;
		Row tab3SheetRow = null;
		Row tab4SheetRow = null;
		Row tab5SheetRow = null;
		Row tab6SheetRow = null;
		for (int key : newRows) {
			ArrayList<String> objArr = sheetData.get(key);
			for (String obj : objArr) {
				if(obj.equalsIgnoreCase("Tab1")) {
					tab1Sheet = outBook.createSheet(obj);
					rownum = tab1Sheet.getLastRowNum();
					tab1SheetRow = tab1Sheet.createRow(rownum++);
					System.out.println("Tab 1 created");
					System.out.println("Row 1 created");
					 tab1Flag =true;
					tab1Data=3;
					break;
					
				}
				if(obj.equalsIgnoreCase("Tab2")) {
					tab2Sheet = outBook.createSheet(obj);
					rownum = tab2Sheet.getLastRowNum();
					tab2SheetRow = tab2Sheet.createRow(rownum++);
					System.out.println("Tab 2 created");
					System.out.println("Row 1 created");
					 tab2Flag =true;
					tab2Data=3;
					break;
					
				}
				if(obj.equalsIgnoreCase("Tab3")) {
					tab3Sheet = outBook.createSheet(obj);
					rownum = tab3Sheet.getLastRowNum();
					tab3SheetRow = tab3Sheet.createRow(rownum++);
					System.out.println("Tab 3 created");
					System.out.println("Row 1 created");
					 tab3Flag =true;
					tab3Data=3;
					break;
					
				}
				if(obj.equalsIgnoreCase("Tab4")) {
					tab4Sheet = outBook.createSheet(obj);
					rownum = tab4Sheet.getLastRowNum();
					tab4SheetRow = tab4Sheet.createRow(rownum++);
					System.out.println("Tab 4 created");
					System.out.println("Row 1 created");
					 tab4Flag =true;
					tab4Data=3;
					break;
					
				}
				if(obj.equalsIgnoreCase("Tab5")) {
					tab5Sheet = outBook.createSheet(obj);
					rownum = tab5Sheet.getLastRowNum();
					tab5SheetRow = tab5Sheet.createRow(rownum++);
					System.out.println("Tab 5 created");
					System.out.println("Row 1 created");
					 tab5Flag =true;
					tab5Data=3;
					break;
					
				}
				if(obj.equalsIgnoreCase("Tab6")) {
					tab6Sheet = outBook.createSheet(obj);
					rownum = tab6Sheet.getLastRowNum();
					tab6SheetRow = tab6Sheet.createRow(rownum++);
					System.out.println("Tab 6 created");
					System.out.println("Row 1 created");
					 tab6Flag =true;
					tab6Data=3;
					break;
					
				}
				if(tab1Flag && tab1Data>0) {
					Cell cell = tab1SheetRow.createCell(cellnum++); 
					cell.setCellValue((obj));
					System.out.println(obj);
					
				}
				if(tab2Flag && tab2Data>0) {
					Cell cell = tab2SheetRow.createCell(cellnum++); 
					cell.setCellValue((obj));
					System.out.println(obj);
					
				}
				if(tab3Flag && tab3Data>0) {
					Cell cell = tab3SheetRow.createCell(cellnum++); 
					cell.setCellValue((obj));
					System.out.println(obj);
					
				}
				if(tab4Flag && tab4Data>0) {
					Cell cell = tab4SheetRow.createCell(cellnum++); 
					cell.setCellValue((obj));
					System.out.println(obj);
					
				}
				if(tab5Flag && tab5Data>0) {
					Cell cell = tab5SheetRow.createCell(cellnum++); 
					cell.setCellValue((obj));
					System.out.println(obj);
					
				}
				if(tab6Flag && tab6Data>0) {
					Cell cell = tab6SheetRow.createCell(cellnum++); 
					cell.setCellValue((obj));
					System.out.println(obj);
					
				}
		}
			if(tab1Flag) {
				tab1Data--;	
				cellnum=0;
			}
			if(tab2Flag) {
				tab2Data--;	
				cellnum=0;
			}
			if(tab3Flag) {
				tab3Data--;	
				cellnum=0;
			}
			if(tab4Flag) {
				tab4Data--;	
				cellnum=0;
			}
			if(tab5Flag) {
				tab5Data--;	
				cellnum=0;
			}
			if(tab6Flag) {
				tab6Data--;	
				cellnum=0;
			}
			if(tab1Data==1) {
				tab1SheetRow = tab1Sheet.createRow(rownum++);
				System.out.println("Row 2 created");
				
			}
			if(tab2Data==1) {
				tab2SheetRow = tab2Sheet.createRow(rownum++);
				System.out.println("Row 2 created");
				cellnum=0;
			}
			if(tab3Data==1) {
				tab3SheetRow = tab3Sheet.createRow(rownum++);
				System.out.println("Row 2 created");
				cellnum=0;
			}
			if(tab4Data==1) {
				tab4SheetRow = tab4Sheet.createRow(rownum++);
				System.out.println("Row 2 created");
				cellnum=0;
			}
			if(tab5Data==1) {
				tab5SheetRow = tab5Sheet.createRow(rownum++);
				System.out.println("Row 2 created");
				cellnum=0;
			}
			if(tab6Data==1) {
				tab6SheetRow = tab6Sheet.createRow(rownum++);
				System.out.println("Row 2 created");
				cellnum=0;
			}
			
		}

		FileOutputStream os = new FileOutputStream(excelOut); 
		outBook.write(os);
//		System.out.println("Writing on Excel file Finished ...");
		os.close(); 
		outBook.close(); 
		book.close();
		fis.close(); 
		} catch (FileNotFoundException fe) 
		{ 
			fe.printStackTrace(); 
			} catch (IOException ie)
				{ 
				ie.printStackTrace(); 
				} }
	
		}














//while (itr.hasNext()) 
//{ 
//	Row row = itr.next(); 
//	Iterator<Cell> cellIterator = row.cellIterator(); 
//while (cellIterator.hasNext()) 
//{ 
//	Cell cell = cellIterator.next(); 
//	if(cell.getStringCellValue().equals("Tab1")) {
//		tab1Flag = true;
//		break;
//	}
//	if(tab1Flag) {
//		tableHeader.add(cell.getStringCellValue());
//		System.out.print(cell.getStringCellValue() + "\t");
////		tab1Flag = false;
////		tab1DataFlag = true;
////		break;
//	}
////	if(tab1DataFlag) {
////		tableData.add(cell.getStringCellValue());
////		System.out.print(cell.getStringCellValue() + "\t");
////		tab1DataFlag = true;
////		break;
////	}
//	if(cell.getStringCellValue().equals("Tab2")) {
//		tab1Flag = false;
//	}
//	
//} 
//
//System.out.println(""); 
//} 
//XSSFSheet tab1Sheet = outBook.createSheet(); 	
////
////
//
//
//{
// {
//Row row = tab1Sheet.createRow(rownum++);
//ArrayList<String> objArr = sheetData.get(key); 
//int cellnum = 0; 
//for (String obj : objArr)
//{ 
//	Cell cell = row.createCell(cellnum++); 
//	cell.setCellValue((obj));
//}
//}
//
//}
