package com.krishagni.tkids.merger;

import java.io.File;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

import com.krishagni.tkids.merger.impl.CSVWriter;

public interface Processor {
	public void process(File file, Workbook workbook, CSVWriter csvWriter) throws IOException;

	public void writeHeaderOriginToCsv(CSVWriter csvWriter) throws IOException;
}
