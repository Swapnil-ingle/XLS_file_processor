package com.krishagni.tkids.merger.impl;

import java.io.File;
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
	
	public XLSReader(String sourcePath) {
		files = new ArrayList<>();
		File folder = new File(sourcePath);
		files.addAll(Arrays.asList(folder.listFiles()));
	}

	@Override
	public List<File> getFileList() {
		return getFiles();
	}

	@Override
	public Workbook getReader(String file) throws EncryptedDocumentException, IOException {
		return WorkbookFactory.create(new File(file));
	}
	
	@Override
	public List<String> getHeadersForFile(File file) throws EncryptedDocumentException, IOException {
		List<String> headers = new ArrayList<>();
		DataFormatter dataFormatter = new DataFormatter();
		
		
		getReader(file.getAbsolutePath()).forEach(sheet -> {
			sheet.getRow(0).forEach(cell -> {
				String val = dataFormatter.formatCellValue(cell);
				if (!headers.contains(val)) {
					headers.add(val);
				}
			});
		});
		
		return headers;
	}

	public List<File> getFiles() {
		return files;
	}

	public void setFiles(List<File> files) {
		this.files = files;
	}

}
