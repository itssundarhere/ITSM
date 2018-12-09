package com.java;

import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.Map.Entry;

public class XlsReadWrite {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		Map<Integer, String> positionOfHeaderMap = new HashMap<Integer, String>(); // Create map
		List<String> columnName = new ArrayList<String>(); // Column name List
		List<String> columnsToBeWrittenFile1 = new ArrayList<String>();
		List<String> columnsToBeWrittenFile2 = new ArrayList<String>();
		List<String> columnsToBeWrittenFile3 = new ArrayList<String>();

		FileInputStream file = new FileInputStream(new File("./ITSM_Entities.xls"));
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		HSSFSheet sheet = workbook.getSheetAt(0);
		HSSFRow row = sheet.getRow(0); // Get first row
		// following is boilerplate from the java doc

		short minColIx = row.getFirstCellNum(); // get the first column index for a row
		short maxColIx = row.getLastCellNum(); // get the last column index for a row
		for (short colIx = minColIx; colIx < maxColIx - 1; colIx++) { // loop from first to last index
			HSSFCell cell = row.getCell(colIx); // get the cell
			positionOfHeaderMap.put(cell.getColumnIndex(), cell.getStringCellValue()); // add the cell contents (name of
																						// column) and
			// cell index to the map
		}
		positionOfHeaderMap.forEach((k, v) -> columnName.add(v)); // printing Column names
//------System.out.println(columnName);

		Iterator hmIterator = positionOfHeaderMap.entrySet().iterator();
		while (hmIterator.hasNext()) {
			Map.Entry mapElement = (Map.Entry) hmIterator.next();
			if (((String) mapElement.getValue()).contains("Project Name")) {
				columnsToBeWrittenFile1.add((String) mapElement.getValue());
				columnsToBeWrittenFile2.add((String) mapElement.getValue());
				columnsToBeWrittenFile3.add((String) mapElement.getValue());
			} else if (((String) mapElement.getValue()).contains("Entity Name")) {
				columnsToBeWrittenFile1.add((String) mapElement.getValue());
				columnsToBeWrittenFile2.add((String) mapElement.getValue());
				columnsToBeWrittenFile3.add((String) mapElement.getValue());
			} else if (((String) mapElement.getValue()).contains("Entity Attributes")) {
				columnsToBeWrittenFile2.add("Attribute Name");
			} else if (((String) mapElement.getValue()).contains("[rel Name] --> Rel Entity []")) {
				columnsToBeWrittenFile3.add("Releated Entity");
				columnsToBeWrittenFile3.add("Releationship Name");
			}
		}

		/*System.out.println(columnsToBeWrittenFile1);
		System.out.println(columnsToBeWrittenFile2);
		System.out.println(columnsToBeWrittenFile3);
		System.out.println(positionOfHeaderMap);*/

		/*------------------------------------------------------------------------------------------------------------------------------*/
		List<Map> dataProcessingList = new ArrayList<Map>();
		/*------------------------------------------------------------------------------------------------------------------------------*/
		/*Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			int index=0;
			row = (HSSFRow)rowIterator.next();
			Iterator < Cell >  cellIterator = row.cellIterator();
			while(cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				for(String colName:columnName) {
					if(colName.equalsIgnoreCase(cell.getStringCellValue())) {
						System.out.println(cell.getStringCellValue());
					}
				}
			}
		}*/
		
		
		
		FileOutputStream fileOut = new FileOutputStream("./poi-test.xls");
		HSSFWorkbook workbook1 = new HSSFWorkbook();
		HSSFSheet worksheet1 = workbook1.createSheet("POI Worksheet");
		Row header = worksheet1.createRow(0);
	    for(int i=0;i<columnsToBeWrittenFile1.size();i++) {						// Header writting
	    	header.createCell(i).setCellValue(columnsToBeWrittenFile1.get(i));
	    }
	    int i_r=0;
		for (Row myrow : sheet) {
			i_r++;
			if(i_r>1) {
				Row cr_row = worksheet1.createRow(i_r-1);
				int i_c=0;
			    for (Cell mycell : myrow) {
			    	i_c++;
			    	if(i_c>1&&i_c<=4) {
			    		if(mycell.getCellTypeEnum()==CellType.STRING&&mycell.getStringCellValue()!=""
			    				+ "") {
				    		cr_row.createCell(i_c-2).setCellValue(mycell.getStringCellValue());
				    		System.out.println("row  : "+i_r+" Col  : "+i_c+" Value : "+mycell.getStringCellValue());
			    		}else {
			    			cr_row.createCell(i_c-2).setCellValue("");
			    		}
			    	}
			    }
			    i_c=0;
			}   
		}
	    workbook1.write(fileOut);
	    workbook1.close();
	    fileOut.close();
	    workbook.close();
		file.close();
	}
}