package com.krishagni.tkids.merger.impl;

import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.krishagni.tkids.merger.Processor;

public class XLSProcessor implements Processor {
	// private Map<String, Integer> finalOutputHeader = new HashMap<>();
	
	private List<String> finalHeader;
	
//	private Map<Integer, String> currentHeader;
	
	private String[] printBufferArray;
	
	private DataFormatter dataFormatter = new DataFormatter();
	
	public XLSProcessor(List<String> finalHeader) {
		this.finalHeader = finalHeader;
		
		// Make a printBufferArray
//		createPrintBufferArray();
	}
	
	@Override
	public void process(Workbook workbook, CSVWriter csvReader) throws IOException {
		for (Sheet sheet : workbook) {
			// Read Header
			// addHeadersToFinalOutput(sheet);
			// Add current header
			// addToCurrentHeader(sheet);
			
			// Process the sheet
			for (Row row : sheet) {
				// Skip header row
				if (row.getRowNum() == 0) {
					continue;
				}
				// create printBufferArray
				initPrintBufferArray();
				int currCellIdx = 0;
				for (Cell cell : row) {
					// Get current val
					String val = dataFormatter.formatCellValue(cell);
					if (val.isEmpty()) {
						currCellIdx++;
						continue;
					}
					
					// Get col header name for current cell
					String currCellHead = dataFormatter.formatCellValue(sheet.getRow(0).getCell(currCellIdx));
					
					// Get idx of curr cell header in final header list
					// Write to the cell[idx] of current row of output CSV
					addToPrintBufferArray(finalHeader.indexOf(currCellHead), val);
					currCellIdx++;
				}
				// CSV moves to next row : Write to CSV
				csvReader.printToCSV(printBufferArray);
			};
		};
		
		// close the workbook.
		workbook.close();
	}
	
	private void initPrintBufferArray() {
		this.printBufferArray = new String[finalHeader.size()];
	}

	private void addToPrintBufferArray(int idx, String val) {
		printBufferArray[idx] = val;
	}

//	private void addHeadersToFinalOutput(Sheet sheet) {
//		sheet.getRow(0).forEach(cell -> {
//			int size = finalOutputHeader.size();
//			final Integer idx = Integer.valueOf(++size); // workaround
//			finalOutputHeader.computeIfAbsent(dataFormatter.formatCellValue(cell), k -> idx);
//		});
//	}
	
//	private void addToCurrentHeader(Sheet sheet) {
//		currentHeader = new HashMap<>();
//		sheet.getRow(0).forEach(cell -> {
//			int size = currentHeader.size();
//			final Integer idx = Integer.valueOf(++size); // workaround
//			currentHeader.computeIfAbsent(idx, k -> dataFormatter.formatCellValue(cell));
//		});
//	}
	
//	public Map<Integer, String> getCurrentHeader() {
//		return currentHeader;
//	}
//
//	public void setCurrentHeader(Map<Integer, String> currentHeader) {
//		this.currentHeader = currentHeader;
//	}
}
