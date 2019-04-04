package com.krishagni.tkids.merger;

import java.io.File;
import java.io.IOException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;

import com.krishagni.tkids.merger.impl.CSVWriter;
import com.krishagni.tkids.merger.impl.XLSProcessor;
import com.krishagni.tkids.merger.impl.XLSReader;

public class Driver {
	private static final String OUTPUT_FILE_NAME = "output";

	private static final String OUTPUT_DIR_NAME = "merge-outputs";

	private static final String OUTPUT_CSV_NAME_DATETIME_FORMAT = "yyyy_MM_dd-HH_mm_ss";

	public static void main(String[] args) {
		try {
			Reader reader = new XLSReader(getSourceFilePath(args));
			List<String> outputHeader = readAllHeaders(reader, reader.getFileList());
			
			Processor processor = new XLSProcessor(outputHeader);
			CSVWriter csvWriter = new CSVWriter(getOutputDir(args[0]), getOutputFileName(args[0]));

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

	private static String getOutputDir(String inputFilesPath) {
		File outputDir = new File(inputFilesPath + File.separatorChar + OUTPUT_DIR_NAME);
		if (!outputDir.exists()) {
			outputDir.mkdir();
		}

		return outputDir.getAbsolutePath();
	}

	private static String getOutputFileName(String outputFile) {
		return OUTPUT_FILE_NAME + "_" + getCurrentTimeStamp() + ".csv";
	}

	private static String getCurrentTimeStamp() {
		Timestamp ts = new Timestamp(System.currentTimeMillis());
		return new SimpleDateFormat(OUTPUT_CSV_NAME_DATETIME_FORMAT).format(ts);
	}

	private static String getSourceFilePath(String[] args) throws ArrayIndexOutOfBoundsException {
		if (args.length <= 0) {
			throw new ArrayIndexOutOfBoundsException("Please provide the path for source files as runtime arguments");
		}

		return args[0];
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
