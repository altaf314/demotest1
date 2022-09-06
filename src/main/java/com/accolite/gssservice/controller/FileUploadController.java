package com.accolite.gssservice.controller;

import java.io.ByteArrayOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.accolite.gssservice.service.CompareSheetUtil;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@CrossOrigin("*")
@RestController
public class FileUploadController {

	@PostMapping("/upload-file")
	public ResponseEntity<?> uploadFile(@RequestParam("file1") MultipartFile file,
			@RequestParam("file2") MultipartFile file2) {

		try {
			log.info("content-type :: {}, {}", file.getContentType(), file.getContentType());

			if (file.isEmpty()) {
				return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("File is Empty");
			}

			if (file2.isEmpty()) {
				return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("File 2 is Empty");
			}

			if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
					&& !file2.getContentType()
							.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
				return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Only Excel File are allowed");
			}

			Workbook workbook = CompareSheetUtil.compareFile(file, file2);

			ByteArrayOutputStream bos = new ByteArrayOutputStream();
			try {
				workbook.write(bos);
			} finally {
				bos.close();
			}
			byte[] bytes = bos.toByteArray();

			HttpHeaders header = new HttpHeaders();
			header.setContentType(
					new MediaType("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
			header.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=result.xls");
			header.setContentLength(bytes.length);

			return ResponseEntity.ok().headers(header).body(bytes);
		} catch (Exception e) {
			log.error("Exception occurred :: ", e);
			e.printStackTrace();
			@SuppressWarnings("deprecation")
			String errorMsg = !StringUtils.isEmpty(e.getMessage()) ? e.getMessage()
					: "Some Thing went wrong! Try after Sometime";
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(errorMsg);
		}
	}

}
