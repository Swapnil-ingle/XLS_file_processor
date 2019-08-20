package com.krishagni.tkids.merger.impl;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.krishagni.tkids.merger.Processor;

public class XLSProcessor implements Processor {
	private List<String> finalHeaders;
	
	private String[] printBufferArray;
	
	private DataFormatter dataFormatter = new DataFormatter();
	
	// Map<headerName, Map<FileName, List<Sheets>>>;
	private Map<String, Map<String, List<String>>> headersOriginMap = new HashMap<String, Map<String,List<String>>>();

	public XLSProcessor(List<String> finalHeaders) {
		this.finalHeaders = finalHeaders;
	}

	public Map<String, Map<String, List<String>>> getHeadersOriginMap() {
		return headersOriginMap;
	}

	public void setHeadersOriginMap(Map<String, Map<String, List<String>>> headersOriginMap) {
		this.headersOriginMap = headersOriginMap;
	}
	
	@Override
	public void process(File file, Workbook workbook, CSVWriter csvWriter) throws IOException {
		for (Sheet sheet : workbook) {
			for (Row row : sheet) {
				if (row.getRowNum() == 0) {
					populateHeadersOriginMap(file.getName(), sheet.getSheetName(), row);
					continue;
				}

				initPrintBufferArray();

				for (Cell cell : row) {
					String val = dataFormatter.formatCellValue(cell);
					if (val.isEmpty()) {
						continue;
					}
					
					String currCellHead = getCurrCellHeader(sheet, cell);
					
					if (currCellHead.isEmpty()) {
						// Ignore stray cells which don't belong under any column header
						continue;
					}

					addToPrintBufferArray(finalHeaders.indexOf(currCellHead), val);
				}
				csvWriter.printToCSV(printBufferArray);
			};
		};
		
		workbook.close();
	}
	
	private void populateHeadersOriginMap(String fileName, String sheetName, Row row) {
		for (Cell cell : row) {
			String header = dataFormatter.formatCellValue(cell);
			if (header.isEmpty()) {
				continue;
			}

			header = header.toLowerCase();

			Map<String, List<String>> origins = null;
			List<String> sheets = null;

			if (headersOriginMap.get(header) == null) {
				// If current header does not exist in the origin map
				origins = new HashMap<>();

				sheets = new ArrayList<String>();
				sheets.add(sheetName);

				origins.put(fileName, sheets);

				// Mark the new header origins <header --> <fileName,[sheet(s)Name]>>
				headersOriginMap.put(header, origins);
			} else {
				sheets = headersOriginMap.get(header).get(fileName);

				if (sheets != null) {
					// If current fileName is present under the current header
					// Add current sheetName under the currFile
					sheets.add(sheetName);
				} else {
					// If current fileName is not added under the current header
					// Add new origins under the current fileName
					Map<String, List<String>> filesUnderCurrentHeaders = headersOriginMap.get(header);

					sheets = new ArrayList<String>();
					sheets.add(sheetName);

					filesUnderCurrentHeaders.put(fileName, sheets);
				}
			}
		}
	}

	@Override
	public void writeHeaderOriginToCsv(CSVWriter csvWriter) throws IOException {
		List<String> row = new ArrayList<>();

		for (Entry<String, Map<String, List<String>>> entry : headersOriginMap.entrySet()) {
			Map<String, List<String>> value = entry.getValue();
			String header = entry.getKey();

			for (Entry<String, List<String>> entry1 : value.entrySet()) {
				String fileName = entry1.getKey();
				List<String> sheets = entry1.getValue();

				row.add(header);
				row.add(fileName);
				row.addAll(sheets);

				csvWriter.printToCSV(row);
				row.clear();
			}

			csvWriter.flushPrinter();
		}
	}

	private String getCurrCellHeader(Sheet sheet, Cell curCell) {
		int currCellIdx = curCell.getAddress().getColumn();
		Cell headerRowCell = sheet.getRow(0).getCell(currCellIdx);
		String cellVal = dataFormatter.formatCellValue(headerRowCell);
		return cellVal.trim().toLowerCase();
	}

	private void initPrintBufferArray() {
		this.printBufferArray = new String[finalHeaders.size()];
	}

	private void addToPrintBufferArray(int idx, String val) {
		printBufferArray[idx] = val;
	}
}
