package main;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class FinalCalculation {
	static Logger logger = Logger.getLogger(FinalCalculation.class);

	public HSSFSheet IVF=null;
	public FinalCalculation(HSSFSheet s) {
		this.IVF=s;
	}

	public void calculateBySectors() {

		File tFIle = new File("IndexVol.xls");

		// for reading data from different file
		File fadjDiv = new File("AdjDiv.xls");
		File fadjRight = new File("AdjRight.xls");
		File fmCap = new File("MCap.xls");
		File ftDivisor = new File("TDivisor.xls");

		try {

			// workbooks for different calculated result
			Workbook wbAdjDiv = WorkbookFactory.create(fadjDiv);
			Workbook wbAdjright = WorkbookFactory.create(fadjRight);
			Workbook wbMCap = WorkbookFactory.create(fmCap);
			Workbook wbTDivisor = WorkbookFactory.create(ftDivisor);

			// get sheets of those calculation
			Sheet adj = wbAdjDiv.getSheetAt(0);
			Sheet adjR = wbAdjright.getSheetAt(0);
			Sheet mP = wbMCap.getSheetAt(0);
			Sheet tDivisor = wbTDivisor.getSheetAt(0);

			logger.debug("file name TDivisor: " + tFIle.getAbsolutePath());
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet TDivisor = wb.createSheet("Index&Vol");

			Map<Integer, String> kMap = new HashMap<Integer, String>();

			sectorMapping SP = new sectorMapping();
			kMap = SP.uniqSectorMapping();

			// Final row count
			int rowForFInalSheet = kMap.size() * 8 + 1;
			// for callculating Market Cap, Entry , Cash Dividend , right share,
			// adjustment
			int preCalRow = adj.getPhysicalNumberOfRows();
			int preCalCol = adj.getRow(0).getPhysicalNumberOfCells();
			int indexOfFinal = 1;
			for (Map.Entry<Integer, String> entry : kMap.entrySet()) {
				String singleSectorValue = entry.getValue().trim();
				// Map<String,Map<String,Double>> singleDayFCal=new
				// HashMap<String,Map<String,Double>>();

				ArrayList<IndexVolClass> vv = new ArrayList<IndexVolClass>();
				// int index =0;
				for (int i = 2; i < preCalCol; i++) {
					double adjrightSum = 0;
					double adjDIvSum = 0;
					double mCapSum = 0;
					double divisor = 165;
					double Adjustment = 0;
					double Entry = 0;
					double total = 0;

					for (int j = 1; j < preCalRow; j++) {

						if(adj.getRow(j).getCell(1)==null) break;
						String Cellvalue=adj.getRow(j).getCell(1).getStringCellValue().trim();
						if (Cellvalue.equals(singleSectorValue)) {
							Row adjRow = adj.getRow(j);
							Cell adjCell = adjRow.getCell(i);
							if (adjCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
								adjDIvSum += adjCell.getNumericCellValue();
							}

							Row rowAdjR = adjR.getRow(j);
							Cell adjRCell = rowAdjR.getCell(i);
							if (adjRCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
								adjrightSum += adjRCell.getNumericCellValue();
							}

							Row mCap = mP.getRow(j);
							Cell mCell = mCap.getCell(i);
							if (mCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
								mCapSum += mCell.getNumericCellValue();
							}
						}
					}
					IndexVolClass vol = new IndexVolClass();
					vol.MarketCap = mCapSum;
					vol.Divisor = 165;
					vol.Entry = 0;
					vol.CashDividend = adjDIvSum;
					vol.RightShare = adjrightSum;
					vol.Adjustments = Entry + adjrightSum - adjDIvSum;
					vol.total = mCapSum / divisor;
					vv.add(vol);
					// index++
				}

				// First row set
				Row frow=TDivisor.createRow(0);
				for(int j=0;j<preCalCol;j++){
					HSSFCell ivfCell=this.IVF.getRow(0).getCell(j);
					String charVal=ivfCell.toString();
					Cell fcell=frow.createCell(j);
					fcell.setCellValue(charVal);
					
				}
				Row rowMarketCap = TDivisor.createRow(indexOfFinal++);
				Row rowDivisor = TDivisor.createRow(indexOfFinal++);
				Row rowEntry = TDivisor.createRow(indexOfFinal++);
				Row rowCashDividend = TDivisor.createRow(indexOfFinal++);
				Row rowRightShare = TDivisor.createRow(indexOfFinal++);
				Row rowAdjustments = TDivisor.createRow(indexOfFinal++);
				Row rowTotal = TDivisor.createRow(indexOfFinal++);

				int arrayIndex = 0;
				// int colCount = rowAdjDiv.getPhysicalNumberOfCells();
				//Assigning 1st column values 
				Cell cellrowMarketCap1 = rowMarketCap.createCell(0);
				Cell cellDivisor1 = rowDivisor.createCell(0);
				Cell cellEntry1 = rowEntry.createCell(0);
				Cell cellCashDividend1 = rowCashDividend.createCell(0);
				Cell cellRIghtShare1 = rowRightShare.createCell(0);
				Cell cellAdjustment1 = rowAdjustments.createCell(0);
				Cell celltotal1 = rowTotal.createCell(0);
				
				String val=singleSectorValue+" Market Cap";
				cellrowMarketCap1.setCellValue(val);
				cellDivisor1.setCellValue("Divisor");
				cellEntry1.setCellValue("Entry");
				cellCashDividend1.setCellValue("Cash Dividend");
				cellRIghtShare1.setCellValue("Right Share");
				cellAdjustment1.setCellValue("Adjustments");
				celltotal1.setCellValue(singleSectorValue);
				
				for (int j = 2; j < preCalCol; j++) {

					Cell cellrowMarketCap = rowMarketCap.createCell(j);
					Cell cellDivisor = rowDivisor.createCell(j);
					Cell cellEntry = rowEntry.createCell(j);
					Cell cellCashDividend = rowCashDividend.createCell(j);
					Cell cellRIghtShare = rowRightShare.createCell(j);
					Cell cellAdjustment = rowAdjustments.createCell(j);
					Cell celltotal = rowTotal.createCell(j);

					cellrowMarketCap.setCellValue(vv.get(arrayIndex).MarketCap);
					cellDivisor.setCellValue(vv.get(arrayIndex).Divisor);
					cellEntry.setCellValue(vv.get(arrayIndex).Entry);
					cellCashDividend.setCellValue(vv.get(arrayIndex).CashDividend);
					cellRIghtShare.setCellValue(vv.get(arrayIndex).RightShare);
					cellAdjustment.setCellValue(vv.get(arrayIndex).Adjustments);
					celltotal.setCellValue(vv.get(arrayIndex).total);
					arrayIndex++;

				}
				//provide a blank row for facilitate the calculation 
				Row blankRor=TDivisor.createRow(indexOfFinal++);
				for(int j=0;j<preCalCol;j++){
					
					Cell fcell=blankRor.createCell(j);
					fcell.setCellValue("");
					
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
