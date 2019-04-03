package com.krishagni.tkids.merger.impl;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

public class CSVWriter {
	private FileWriter writer;
	
	private CSVPrinter csvPrinter;
	
	public CSVWriter(File file) throws IOException {
		this.writer = new FileWriter(new File(file, "output.csv"), true);
		this.csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT);
	}
	
	public CSVWriter(String folder, String file) throws IOException {
		this.writer = new FileWriter(new File(folder, file), true);
		this.csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT);
	}
	
	public void writeHeader(List<String> headers) throws IOException {
		printToCSV(headers);
	}
	
	public void printToCSV(String[] data) throws IOException {
		csvPrinter.printRecord(Arrays.asList(data));
	}
	
	public void printToCSV(Collection<String> data) throws IOException {
		csvPrinter.printRecord(data);
	}
	
	public void closeCsvWriter() throws IOException {
		writer.close();
		csvPrinter.close();
	}
	
	public void flushPrinter() throws IOException {
		csvPrinter.flush();
	}

	public FileWriter getWriter() {
		return writer;
	}

	public void setWriter(FileWriter writer) {
		this.writer = writer;
	}

	public CSVPrinter getCsvPrinter() {
		return csvPrinter;
	}

	public void setCsvPrinter(CSVPrinter csvPrinter) {
		this.csvPrinter = csvPrinter;
	}
}
