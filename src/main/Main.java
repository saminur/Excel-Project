package main;

import java.io.File;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
//	static Logger logger = Logger.getLogger(Main.class);
//
//	public static void main(String[] args) throws Exception {
//		org.apache.log4j.PropertyConfigurator.configure("log4j.properties");
//
//		File f = new File("CompiledIndex.xlsm");
//
//		String[] sh = { "Price", "SharedNumber", "CDividend", "Multilier",
//				"RightP" };
//
//		System.out.println("file name: " + f.getAbsolutePath());
//		try {
//			// Workbook wb=Workbook.getWorkbook(f);
//			XSSFWorkbook wb = new XSSFWorkbook(f);
//			
//			
//			XSSFSheet s1 = wb.getSheetAt(1);
//			XSSFSheet s2 = wb.getSheetAt(2);
//			XSSFSheet s3 = wb.getSheetAt(3);
//			//calculating AdjDiv
//			AdjDivCalculation adjDiv = new AdjDivCalculation(s1, s2);
//			adjDiv.calculateAdjDiv();
//			
//			//calculate AdjRIght 
//			AdjRightCalculation adjRight=new AdjRightCalculation(s1, s2, s3);
//			adjRight.calculateAdjRight();
//			
//			//Calculate MCap
//			MCap claculateMCap=new MCap(wb.getSheetAt(0),s1);
//			claculateMCap.calculateMCap();
//			
//			//calculate TDivisor
//			TDivisor tDivisor=new TDivisor();
//			tDivisor.calculateTDivisor();
//			
//			System.out.println("got sheet");
//			for (int i = 0; i < sh.length; i++) {
//
//				// System.out.println("-----------------Loop"+" "+i+" For Sheet "
//				// +sh[i]+" ----------------");
//				XSSFSheet s = wb.getSheet(sh[i]);
//				int rowCount = 0;
//				if (s != null)
//					rowCount = s.getPhysicalNumberOfRows();
//				else
//					continue;
//				// System.out.println("rowCount: "+rowCount);
//				for (int i1 = 0; i1 < rowCount; i1++) {
//					XSSFRow row = s.getRow(i1);
//					int colCount = row.getPhysicalNumberOfCells();
//					// System.out.println("colCount: "+colCount);
//					for (int j = 0; j < colCount; j++) {
//						if (j % 16 == 0)
//							System.out.println();
//						Cell c = row.getCell(j);
//						// if(c!=null)System.out.print(" c.getStringCellValue(): "+c.toString());
//					}
//					// System.out.println("\n");
//
//				}
//			}
//			// wb.close();
//		} catch (Exception e) {
//			logger.fatal("exception at reading", e);
//		} finally {
//			try {
//				// wb.close();
//			} catch (Exception e) {
//				logger.fatal("exception e", e);
//			}
//		}
//
//		System.out.println("completed");
//	}

}
