package com.krishagni.tkids.merger;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;

public interface Reader {
	public List<File> getFileList();
	
	public Workbook getReader(String filePath) throws EncryptedDocumentException, IOException;

	List<String> getHeadersForFile(File file) throws EncryptedDocumentException, IOException;

	void closeXLSReader() throws IOException;
}
