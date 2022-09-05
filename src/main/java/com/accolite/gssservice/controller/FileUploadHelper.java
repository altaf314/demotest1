package com.accolite.gssservice.controller;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

@Component
public class FileUploadHelper {

//	public final String UPLOAD_DIR = "D:\\Learning\\GoogleSheetProject\\gss-backend\\UploadFile\\src\\main\\resources\\static\\excelfile";

	public final String UPLOAD_DIR = new ClassPathResource("static/excelfile/").getFile().getAbsolutePath();

	public FileUploadHelper() throws IOException {
		super();
	}

	public boolean uploadFile(MultipartFile multipartFile) {
		boolean f = false;

		try {

			Files.copy(multipartFile.getInputStream(),
					Paths.get(UPLOAD_DIR + File.separator + multipartFile.getOriginalFilename()),
					StandardCopyOption.REPLACE_EXISTING);

			f = true;

		} catch (Exception e) {
			e.printStackTrace();
		}

		return f;
	}
	public void getFile()
	{
		
	}
}
