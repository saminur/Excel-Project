package main;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class AdjRightCalculation {
	static Logger logger = Logger.getLogger(AdjRightCalculation.class);
	public HSSFSheet shareNumber = null;
	public HSSFSheet multiplier = null;
	public HSSFSheet rightP = null;

	public AdjRightCalculation(HSSFSheet s1, HSSFSheet s2, HSSFSheet s3) {
		this.shareNumber = s1;
		this.multiplier = s2;
		this.rightP = s3;
	}

	public void calculateAdjRight() {
		// TODO Auto-generated method stub
		File f = new File("AdjRight.xls");
		int rowCount = 0;
		rowCount = this.shareNumber.getPhysicalNumberOfRows();
		HSSFRow r = this.shareNumber.getRow(0);
		int col = r.getPhysicalNumberOfCells();
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet adjDiv = wb.createSheet("AdjRight");
		try {
			
			sector ss=new sector();
			int value=1;
			ArrayList<String> sector1= ss.sectorValues();
			
			for (int i = 0; i < rowCount; i++) {
				HSSFRow row = this.shareNumber.getRow(i);
				if (row == null)
					break;

				System.out.println("priting first two column and First row");
				Row rowAd = adjDiv.createRow(i);

				for (int j = 0; j < col; j++) {
					Cell c = rowAd.createCell(j);

					if (i == 0 && row.getCell(j) != null) {
						String v = row.getCell(j).toString();
						System.out.println(" " + v);
						c.setCellValue(v);
					} else {
						if (j == 2)
							break;
						else if(j==1){
							c.setCellValue(sector1.get(value));
							value++;
						}
						else if (row.getCell(j) != null) {
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
				HSSFRow rowMultiplier = this.multiplier.getRow(i);
				HSSFRow rowRightP = this.multiplier.getRow(i);
				int colCount = row.getPhysicalNumberOfCells();
				// System.out.println("colCount: "+colCount);
				Row rowAd = adjDiv.getRow(i);
				for (int j = 2; j < colCount; j++) {
					if (row.getCell(j) == null)
						break;
					Cell c = rowAd.createCell(j);
					Cell c1 = row.getCell(j);
					Cell c2 = rowMultiplier.getCell(j);
					Cell c3 = rowRightP.getCell(j);
					//Cell C4 = null;
					if (c1==null || c2==null || c3==null) {
						// Cell c =sheet.getRow(i).getCell(j);
						c.setCellValue("-");
					} else {
						double value1 = c1.getNumericCellValue();
						double value2 = c2.getNumericCellValue();
						double value3=c3.getNumericCellValue();
						double result = value1 * value2*value3;
						// Cell c =sheet.getRow(i).getCell(j);
						c.setCellValue(result);
					}
				}
			}
			//System.out.println();
			
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
