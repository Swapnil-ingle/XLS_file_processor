package com.krishagni.tkids.merger;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

import com.krishagni.tkids.merger.impl.CSVWriter;

public interface Processor {
//	public void process(Workbook workbook);
	
//	public Map<String, Integer> getFinalOutputHeader();

	void process(Workbook workbook, CSVWriter csvReader) throws IOException;
}
