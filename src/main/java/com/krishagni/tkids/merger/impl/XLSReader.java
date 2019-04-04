package com.krishagni.tkids.merger.impl;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.krishagni.tkids.merger.Reader;

public class XLSReader implements Reader {
	
	private List<File> files;
	
	private Workbook workbook;

	private static final String INPUT_FILE_EXTNS = ".xlsx";

	public XLSReader(String sourcePath) {
		files = new ArrayList<>();
		File folder = new File(sourcePath);
		files.addAll(Arrays.asList(folder.listFiles(applyFilter())));
	}

	private FilenameFilter applyFilter() {
		return new FilenameFilter() {
			@Override
			public boolean accept(File file, String name) {
				return name.endsWith(INPUT_FILE_EXTNS);
			}
		};
	}

	@Override
	public List<File> getFileList() {
		return files;
	}

	@Override
	public Workbook getReader(String file) throws EncryptedDocumentException, IOException {
		workbook = WorkbookFactory.create(new File(file));
		return workbook;
	}
	
	@Override
	public List<String> getHeadersForFile(File file) throws EncryptedDocumentException, IOException {
		List<String> headers = new ArrayList<>();
		DataFormatter dataFormatter = new DataFormatter();
		
		getReader(file.getAbsolutePath()).forEach(sheet -> {
			sheet.getRow(0).forEach(cell -> {
				String val = dataFormatter.formatCellValue(cell);
				if (!val.isEmpty() && !headers.contains(val)) {
					headers.add(val.toLowerCase().trim());
				}
			});
		});
		
		closeXLSReader();
		return headers;
	}

	@Override
	public void closeXLSReader() throws IOException {
		workbook.close();
	}
}
