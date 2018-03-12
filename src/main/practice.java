package main;

import java.io.File;
import java.io.FileInputStream;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class practice {
	static Logger logger = Logger.getLogger(practice.class);

	public static void main(String[] args) throws Exception {
		org.apache.log4j.PropertyConfigurator.configure("log4j.properties");

		
		String[] sh = { "Price", "SharedNumber", "CDividend", "Multilier",
				"RightP","Index&Vol" };

		
		try {
			File f = new File("CompiledIndex.xls");
			System.out.println("file name: " + f.getAbsolutePath());
			FileInputStream file=new FileInputStream(f);
			// Workbook wb=Workbook.getWorkbook(f);
			HSSFWorkbook wb = new HSSFWorkbook(file);
			
			
			HSSFSheet s1 = wb.getSheetAt(1);
			HSSFSheet s2 = wb.getSheetAt(2);
			HSSFSheet s3 = wb.getSheetAt(3);
			HSSFSheet s0=wb.getSheetAt(0);
			//calculating AdjDiv
			AdjDivCalculation adjDiv = new AdjDivCalculation(s1, s2);
			adjDiv.calculateAdjDiv();
			
			//calculate AdjRIght 
			AdjRightCalculation adjRight=new AdjRightCalculation(s1, s2, s3);
			adjRight.calculateAdjRight();
			
			//Calculate MCap
			MCap claculateMCap=new MCap(wb.getSheetAt(0),s1);
			claculateMCap.calculateMCap();
			
			//calculate TDivisor
			TDivisor tDivisor=new TDivisor(s0);
			tDivisor.calculateTDivisor();
			
			//calculate TReturn
			TReturn tReturn=new TReturn();
			tReturn.calculateTReturn();
			
			//calculate FinalCalculation
			FinalCalculation fc=new FinalCalculation(wb.getSheetAt(5));
			fc.calculateBySectors();
			
			System.out.println("got sheet");
			for (int i = 0; i < sh.length; i++) {

				// System.out.println("-----------------Loop"+" "+i+" For Sheet "
				// +sh[i]+" ----------------");
				HSSFSheet s = wb.getSheet(sh[i]);
				int rowCount = 0;
				if (s != null)
					rowCount = s.getPhysicalNumberOfRows();
				else
					continue;
				// System.out.println("rowCount: "+rowCount);
				for (int i1 = 0; i1 < rowCount; i1++) {
					HSSFRow row = s.getRow(i1);
					int colCount = row.getPhysicalNumberOfCells();
					// System.out.println("colCount: "+colCount);
					for (int j = 0; j < colCount; j++) {
						if (j % 16 == 0)
							System.out.println();
						Cell c = row.getCell(j);
						// if(c!=null)System.out.print(" c.getStringCellValue(): "+c.toString());
					}
					// System.out.println("\n");

				}
			}
			// wb.close();
			file.close();
			wb.close();
		} catch (Exception e) {
			logger.fatal("exception at reading", e);
		} finally {
			try {
				// wb.close();
			} catch (Exception e) {
				logger.fatal("exception e", e);
			}
		}

		System.out.println("completed");
	}


}
