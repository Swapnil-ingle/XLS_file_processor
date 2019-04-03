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
	private List<String> finalHeader;
	
	private String[] printBufferArray;
	
	private DataFormatter dataFormatter = new DataFormatter();
	
	public XLSProcessor(List<String> finalHeader) {
		this.finalHeader = finalHeader;
	}
	
	@Override
	public void process(Workbook workbook, CSVWriter csvReader) throws IOException {
		for (Sheet sheet : workbook) {
			for (Row row : sheet) {
				if (row.getRowNum() == 0) {
					continue;
				}

				initPrintBufferArray();
				int currCellIdx = 0;

				for (Cell cell : row) {
					String val = dataFormatter.formatCellValue(cell);
					if (val.isEmpty()) {
						currCellIdx++;
						continue;
					}
					
					String currCellHead = dataFormatter.formatCellValue(sheet.getRow(0).getCell(currCellIdx));
					
					addToPrintBufferArray(finalHeader.indexOf(currCellHead), val);
					currCellIdx++;
				}
				csvReader.printToCSV(printBufferArray);
			};
		};
		
		workbook.close();
	}
	
	private void initPrintBufferArray() {
		this.printBufferArray = new String[finalHeader.size()];
	}

	private void addToPrintBufferArray(int idx, String val) {
		printBufferArray[idx] = val;
	}
}
