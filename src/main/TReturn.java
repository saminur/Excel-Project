package main;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class TReturn {
	
	static Logger logger = Logger.getLogger(TReturn.class);
	
	public TReturn() {
		
	}

	public void calculateTReturn() throws IOException, NullPointerException {
		
		File tFIle = new File("TReturn.xls");
		
		//for reading data from different file
		File fmCap = new File("MCap.xls");
		File ftDiv = new File("TDivisor.xls");
		System.out.println("file name 1: " + fmCap.getAbsolutePath());
		System.out.println("file name 2: " +ftDiv.getAbsolutePath() );

		try {
			
			
			Workbook wb1 = WorkbookFactory.create(fmCap);
			Workbook wb2 = WorkbookFactory.create(ftDiv);
			// XSSFWorkbook wb1 = new XSSFWorkbook(fadjDiv);
			
			Sheet mP = wb1.getSheetAt(0);
			Sheet tDivisor = wb2.getSheetAt(0);
			

			logger.debug("file name TReturn: " + tFIle.getAbsolutePath());
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet TReturn = wb.createSheet("TReturn");

			int rowCount = tDivisor.getPhysicalNumberOfRows();
			//int col=adj.get(
			sector ss=new sector();
			int value=1;
			ArrayList<String> sector1= ss.sectorValues();
			
			int col1=tDivisor.getRow(0).getPhysicalNumberOfCells();
			for (int i = 0; i < rowCount; i++) {
				Row row = tDivisor.getRow(i);
				if (row == null)
					break;

				System.out.println("priting first two column of Treturn");
				Row rowTR = TReturn.createRow(i);

				for (int j = 0; j < col1; j++) {
					Cell c = rowTR.createCell(j);
					
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

				Row rowTR = TReturn.getRow(i);
				//Cell pricerow = this.price.getRow(i).getCell(2);
				Row rTD = tDivisor.getRow(i);
				if (rTD == null)
					break;
				Row rowmCap = mP.getRow(i);

				int colCount = rTD.getPhysicalNumberOfCells();
				for (int j = 2; j < colCount; j++) {
					if (rTD.getCell(j) == null)
						break;
					Cell cellTR = rowTR.createCell(j);
					Cell celltD = rTD.getCell(j);					
					Cell cellmCap = rowmCap.getCell(j);
					
					if (celltD.getCellType() == Cell.CELL_TYPE_NUMERIC
							&& cellmCap.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						double value1 = celltD.getNumericCellValue();
						double value3 = cellmCap.getNumericCellValue();
						System.out.println("Value1: " + value1  + " Value3: " + value3);
						double vv = -1;
						if (value1 == 0) {
							cellTR.setCellValue(0);
							continue;
						}
						vv = value3/ value1;
						cellTR.setCellValue(vv);
					}
					else{
						cellTR.setCellValue("-");
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
