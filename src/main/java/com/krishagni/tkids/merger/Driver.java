package com.krishagni.tkids.merger;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;

import com.krishagni.tkids.merger.impl.CSVWriter;
import com.krishagni.tkids.merger.impl.XLSProcessor;
import com.krishagni.tkids.merger.impl.XLSReader;

public class Driver {
	public static final String SOURCE_PATH = "/Users/swapnil/links/tkids-merge-xls-script/input-files";
	
	public static final String OUTPUT_FILE_NAME = "output.csv";

	public static void main(String[] args) {
		try {
			Reader reader = new XLSReader(SOURCE_PATH);
			List<String> outputHeader = readAllHeaders(reader, reader.getFileList());
			
			Processor processor = new XLSProcessor(outputHeader);
			CSVWriter csvWriter = new CSVWriter(SOURCE_PATH, OUTPUT_FILE_NAME);

			csvWriter.printToCSV(outputHeader);
			
			for (File file : reader.getFileList()) {
				processor.process(reader.getReader(file.getAbsolutePath()), csvWriter);
				csvWriter.flushPrinter();
			}
			
			csvWriter.closeCsvWriter();
		} catch (Exception e) {
			System.out.println("Error while processing: ");
			e.printStackTrace();
		} finally {
			System.out.println("Done processing!!");
		}
	}

	private static List<String> readAllHeaders(Reader reader, List<File> fileList) 
			throws EncryptedDocumentException, IOException {
		List<String> headers = new ArrayList<>();
		
		for (File file : fileList) {
			reader.getHeadersForFile(file).forEach(header -> {
				if (!headers.contains(header)) {
					headers.add(header);
				}
			});
		}
		return headers;
	}
}
