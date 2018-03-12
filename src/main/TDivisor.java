package main;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TDivisor {
	static Logger logger = Logger.getLogger(TDivisor.class);
	public HSSFSheet price = null;

	public TDivisor(HSSFSheet s0) {
		this.price = s0;
	}

	public void calculateTDivisor() throws IOException, NullPointerException {
		
		File tFIle = new File("TDivisor.xls");
		
		//for reading data from different file
		File fadjDiv = new File("AdjDiv.xls");
		File fadjRight = new File("AdjRight.xls");
		File fmCap = new File("MCap.xls");
		System.out.println("file name 1: " + fadjDiv.getAbsolutePath());
		System.out.println("file name 2: " + fadjRight.getAbsolutePath());
		System.out.println("file name 3: " + fmCap.getAbsolutePath());

		try {
			Workbook wb1 = WorkbookFactory.create(fadjDiv);
			// XSSFWorkbook wb1 = new XSSFWorkbook(fadjDiv);

			Workbook wb2 = WorkbookFactory.create(fadjRight);

			Workbook wb3 = WorkbookFactory.create(fmCap);

			Sheet adj = wb1.getSheetAt(0);
			Sheet adjR = wb2.getSheetAt(0);
			Sheet mP = wb3.getSheetAt(0);

			logger.debug("file name TDivisor: " + tFIle.getAbsolutePath());
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet TDivisor = wb.createSheet("TDivisor");

			int rowCount = adj.getPhysicalNumberOfRows();
			//int col=adj.get(
			sector ss=new sector();
			int value=1;
			ArrayList<String> sector1= ss.sectorValues();
			int col1=adj.getRow(0).getPhysicalNumberOfCells();
			for (int i = 0; i < rowCount; i++) {
				Row row = adj.getRow(i);
				if (row == null)
					break;

				System.out.println("priting first two column");
				Row rowTD = TDivisor.createRow(i);

				for (int j = 0; j < col1; j++) {
					Cell c = rowTD.createCell(j);
					
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

				Row rowTD = TDivisor.getRow(i);
				Cell pricerow = this.price.getRow(i).getCell(2);
				Row rowAdjDiv = adj.getRow(i);
				if (rowAdjDiv == null)
					break;
				Row rowadjRight = adjR.getRow(i);
				Row rowmCap = mP.getRow(i);
				Cell Tt = rowTD.createCell(2);
				
				if(pricerow==null){
					Tt.setCellValue("-");
				}
				else if (pricerow.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					Tt.setCellValue(pricerow.getNumericCellValue());
				} else {
					Tt.setCellValue("-");
				}
				int colCount = rowAdjDiv.getPhysicalNumberOfCells();
				for (int j = 3; j < colCount; j++) {
					if (rowAdjDiv.getCell(j) == null)
						break;
					Cell cellTD = rowTD.createCell(j);
					Cell celladjDIv = rowAdjDiv.getCell(j);
					Cell celladjRight = rowadjRight.getCell(j);
					Cell cellmCap = rowmCap.getCell(j);

					Cell prev = rowTD.getCell(j - 1);
					if(prev==null || prev.getCellType()==Cell.CELL_TYPE_STRING){
						cellTD.setCellValue("-");
						continue;
					}

					if (celladjDIv.getCellType() == Cell.CELL_TYPE_NUMERIC
							&& celladjRight.getCellType() == Cell.CELL_TYPE_NUMERIC
							&& cellmCap.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						double value1 = celladjDIv.getNumericCellValue();
						double value2 = celladjRight.getNumericCellValue();
						double value3 = cellmCap.getNumericCellValue();
						System.out.println("Value1: " + value1 + " Value2: " + value2 + " Value3: " + value3);
						double vv = -1;
						if (value3 == 0) {
							cellTD.setCellValue(0);
							continue;
						}
						vv = (value3 - value1 + value2) / value3;
						double result = vv * prev.getNumericCellValue();
						cellTD.setCellValue(result);
					}
				}
			}

			FileOutputStream outFile2 = new FileOutputStream(tFIle);
			wb.write(outFile2);
			outFile2.close();
			wb.close();
		} catch (Exception e) {
			logger.fatal("exception at reading", e);
		}
	}

}
