package main;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AdjDivCalculation {
	static Logger logger = Logger.getLogger(AdjDivCalculation.class);
	public HSSFSheet shareNumber = null;
	public HSSFSheet CDividend = null;

	public AdjDivCalculation(HSSFSheet s1, HSSFSheet s2) {
		this.shareNumber = s1;
		this.CDividend = s2;
	}

	public void calculateAdjDiv() {
		File f = new File("AdjDiv.xls");
		int rowCount = 0;
		rowCount = this.shareNumber.getPhysicalNumberOfRows();
		HSSFRow r = this.shareNumber.getRow(0);
		int col = r.getPhysicalNumberOfCells();
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet adjDiv = wb.createSheet("AdjDiv");

		try {
			sector ss=new sector();
			int value=1;
			ArrayList<String> sector1= ss.sectorValues();
			for (int i = 0; i < rowCount; i++) {
				HSSFRow row = this.shareNumber.getRow(i);
				if (row == null)
					break;

				System.out.println("priting first two column");
				Row rowAd = adjDiv.createRow(i);

				for (int j = 0; j < col; j++) {
					Cell c = rowAd.createCell(j);
					
					if (i == 0 && row.getCell(j) != null) {
						String v = row.getCell(j).toString();
						System.out.println(" " + v);
						c.setCellValue(v);
					} else {
						if (j == 2 || row.getCell(j)==null)
						{
							break;
						}
							
						else if(j==1){
							c.setCellValue(sector1.get(value));
							System.out.println(sector1.get(value));
							value++;
						}
						else{
							String v = row.getCell(j).toString();
							System.out.println(" " + v);
							c.setCellValue(v);
						}
					}

				}
			}

			for (int i = 1; i < rowCount; i++) {

				HSSFRow row = this.shareNumber.getRow(i);
				if (row == null)
					break;
				HSSFRow rowCdividend = this.CDividend.getRow(i);
				int colCount = row.getPhysicalNumberOfCells();
				// System.out.println("colCount: "+colCount);
				Row rowAd = adjDiv.getRow(i);
				for (int j = 2; j < colCount; j++) {
					if (row.getCell(j) == null)
						break;
					Cell c = rowAd.createCell(j);
					Cell c1 = row.getCell(j);
					Cell c2 = rowCdividend.getCell(j);
					Cell C3 = null;
					if (c1.equals(null) || c2.equals(null)) {
						// Cell c =sheet.getRow(i).getCell(j);
						c.setCellValue("-");
					} else {
						double value1 = c1.getNumericCellValue();
						double value2 = c2.getNumericCellValue();
						double result = value1 * value2;
						// Cell c =sheet.getRow(i).getCell(j);
						c.setCellValue(result);
					}
				}
			}
			System.out.println();
			
			FileOutputStream outFile = new FileOutputStream(f);
			wb.write(outFile);
			outFile.close();
			wb.close();

		} catch (FileNotFoundException e) {
			logger.fatal("exception : " + e);
		} catch (IOException e) {
			logger.fatal("exception : " + e);
		}

	}
}
