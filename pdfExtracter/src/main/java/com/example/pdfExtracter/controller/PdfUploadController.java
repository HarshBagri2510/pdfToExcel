package com.example.pdfExtracter.controller;

import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.DirectoryNotEmptyException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;



import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;

@RestController
@CrossOrigin
@RequestMapping(value="/api/upload")
public class PdfUploadController {

	@PostMapping(value="/pdfUpload",produces="application/json")
	/* TODO:Remove all the hard path to properties file */
	public String fileUploadForPack(HttpServletRequest request, HttpServletResponse response) throws Exception {
		//Response responseD = new Response("200", "Success");
		MultipartHttpServletRequest mRequest;
		String filename = "output.pdf";
		String pdfValues=null;
		//ProcessorderProd orders =null;
		String productionOrder=null;
		try {
			mRequest = (MultipartHttpServletRequest) request;
			mRequest.getParameterMap();

			Iterator itr = mRequest.getFileNames();

			while (itr.hasNext()) {
				List<MultipartFile> mFiles = mRequest.getFiles((String) itr.next());
				for(MultipartFile mFile: mFiles) {
				String fileName = mFile.getOriginalFilename();
				//System.out.println(fileName);

				java.nio.file.Path path = Paths.get("D:/CMS_HOSKOTE/ProcessOrders/temp/" + filename);
				Files.deleteIfExists(path);
				InputStream in1 = mFile.getInputStream();
				Files.copy(in1, path);
			    pdfValues = readPdf();
			    //System.out.println("pdfValues "+pdfValues);
			    try (// Create a Word document
				XWPFDocument wordDocument = new XWPFDocument()) {
					XWPFParagraph paragraph = wordDocument.createParagraph();
					paragraph.createRun().setText(pdfValues);

					// Save the Word document
					File outputFile = new File("D:/CMS_HOSKOTE/ProcessOrders/temp/" + "output.pdf");
					try (FileOutputStream out = new FileOutputStream(outputFile)) {
					    wordDocument.write(out);
					}
				}
	            boolean value = new File("D:/CMS_HOSKOTE/ProcessOrders/temp/output/package/" + "/").mkdirs();
				java.nio.file.Path pathForOutput = Paths
						.get("D:/CMS_HOSKOTE/ProcessOrders/temp/output/package/" + "/" + productionOrder + ".pdf");
				InputStream in2 = mFile.getInputStream();
				Files.deleteIfExists(pathForOutput);
				//System.out.println("File is existing and successfully deleted");								
				Files.copy(in2, pathForOutput);
				
			}
			}
		} catch (Exception e) {
//			responseD.setCode("400");
//			responseD.setData(productionOrder);
			e.printStackTrace();
		}
		//responseD.setData(productionOrder);
		return pdfValues;
	}

	public static String readPdf() throws IOException {
		//System.out.println("Main Method Started");
		File file = new File("D:/CMS_HOSKOTE/ProcessOrders/temp/output.pdf");
		PDDocument document = PDDocument.load(file);
		PDFTextStripper pdfStripper = new PDFTextStripper();
		String text = pdfStripper.getText(document);
		//System.out.println(text);
		text = text.trim();
		text = text.replaceAll(" +", " ");
		text = text.replaceAll("(?m)^[ \t]*\r?\n", "");
		//System.out.println("text "+text);
		deleteIfExist();
		writeToFile(text);
		//PdfValuesForPack infos = readData();
		document.close();
		//System.out.println("Main Method Ended");
		return text;

	}

	public static void deleteIfExist() {
		try {
			Files.deleteIfExists(Paths.get("D:\\CMS_HOSKOTE\\ProcessOrders\\txt\\temp.txt"));
		} catch (NoSuchFileException e) {
			//System.out.println("No such file/directory exists");
		} catch (DirectoryNotEmptyException e) {
			//System.out.println("Directory is not empty.");
		} catch (IOException e) {
			//System.out.println("Invalid permissions.");
		}

		//System.out.println("Deletion successful.");
	}

	public static void writeToFile(String content) {
		BufferedWriter bw = null;
		FileWriter fw = null;
		try {
			fw = new FileWriter("D:\\CMS_HOSKOTE\\ProcessOrders\\txt\\temp.txt");
			bw = new BufferedWriter(fw);
			bw.write(content);
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			try {
				if (bw != null)
					bw.close();

				if (fw != null)
					fw.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
	}
	
	@PostMapping(value="/convert", produces="application/json")
	public ResponseEntity<byte[]> convertPdfToExcel(@RequestParam("file") MultipartFile file) {
	    try {
	        // Load PDF document
	        PDDocument document = PDDocument.load(file.getInputStream());
	        PDFTextStripper stripper = new PDFTextStripper();
	        String text = stripper.getText(document);
	        document.close();

	        // Create Excel workbook
	        XSSFWorkbook workbook = new XSSFWorkbook();
	        Sheet sheet = workbook.createSheet("PDF Data");

	        // Create a bold font and apply it to cell style
	        XSSFFont boldFont = workbook.createFont();
	        boldFont.setBold(true);
	        XSSFCellStyle boldStyle = workbook.createCellStyle();
	        boldStyle.setFont(boldFont);

	        // Create header row
	        String[] headers = {"SERVICE NAME", "SERVICE RESOURCE", "SPEND"};
	        Row headerRow = sheet.createRow(0);
	        for (int i = 0; i < headers.length; i++) {
	            Cell cell = headerRow.createCell(i);
	            cell.setCellValue(headers[i]);
	            cell.setCellStyle(boldStyle); // Apply bold style to header cells
	            sheet.autoSizeColumn(i);
	        }

	        // Split the PDF text by lines
	        String[] lines = text.split("\n");
	        int rowNum = 1;

	        // Regular expression to match lines with three parts: "SERVICE NAME", "SERVICE RESOURCE", "SPEND"
	        Pattern pattern = Pattern.compile("^(.*?)(\\s+)(.*?)(\\s+)(\\$[0-9,.]+)$");

	        for (String line : lines) {
	            Matcher matcher = pattern.matcher(line.trim());
	            if (matcher.find()) {
	                Row row = sheet.createRow(rowNum++);
	                Cell serviceNameCell = row.createCell(0);
	                Cell serviceResourceCell = row.createCell(1);
	                Cell spendCell = row.createCell(2);

	                serviceNameCell.setCellValue(matcher.group(1).trim());
	                serviceResourceCell.setCellValue(matcher.group(3).trim());
	                spendCell.setCellValue(matcher.group(5).trim());

	                // Adjust columns width
	                sheet.autoSizeColumn(0);
	                sheet.autoSizeColumn(1);
	                sheet.autoSizeColumn(2);
	            }
	        }

	        // Write to ByteArrayOutputStream
	        ByteArrayOutputStream out = new ByteArrayOutputStream();
	        workbook.write(out);
	        workbook.close();

	        // Return the Excel file as a byte array
	        return ResponseEntity.ok()
	                .header("Content-Disposition", "attachment; filename=\"converted.xlsx\"")
	                .body(out.toByteArray());

	    } catch (IOException e) {
	        e.printStackTrace();
	        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
	    }
	}


//	@PostMapping(value="/convert", produces="application/json")
//	public ResponseEntity<byte[]> convertPdfToExcel(@RequestParam("file") MultipartFile file) {
//	    try {
//	        // Load PDF document
//	        PDDocument document = PDDocument.load(file.getInputStream());
//	        PDFTextStripper stripper = new PDFTextStripper();
//	        String text = stripper.getText(document);
//	        document.close();
//
//	        // Create Excel workbook
//	        XSSFWorkbook workbook = new XSSFWorkbook();
//	        Sheet sheet = workbook.createSheet("PDF Data");
//
//	        // Create additional data row
//	        Row additionalDataRow = sheet.createRow(0);
//	        Cell sponsorshipCell = additionalDataRow.createCell(0);
//	        Cell subscriptionCostCell = additionalDataRow.createCell(1);
//
//	        // Extract additional data (Assuming it is the first part of the text)
//	        String[] lines = text.split("\n");
//	        String sponsorship = lines[0].trim();
//	        String subscriptionCost = lines.length > 1 ? lines[1].trim() : "";
//
//	        sponsorshipCell.setCellValue(sponsorship);
//	        subscriptionCostCell.setCellValue(subscriptionCost);
//
//	        // Create header row
//	        String[] headers = {"SERVICE NAME", "SERVICE RESOURCE", "SPEND"};
//	        Row headerRow = sheet.createRow(1); // Adjusted to row 1 to accommodate additional data
//	        for (int i = 0; i < headers.length; i++) {
//	            Cell cell = headerRow.createCell(i);
//	            cell.setCellValue(headers[i]);
//	            sheet.autoSizeColumn(i);
//	        }
//
//	        // Regular expression to match lines with three parts: "SERVICE NAME", "SERVICE RESOURCE", "SPEND"
//	        Pattern pattern = Pattern.compile("^(.*?)(\\s+)(.*?)(\\s+)(\\$[0-9,.]+)$");
//
//	        int rowNum = 2; // Start at row 2 to skip the additional data and header row
//	        for (String line : lines) {
//	            Matcher matcher = pattern.matcher(line.trim());
//	            if (matcher.find()) {
//	                Row row = sheet.createRow(rowNum++);
//	                Cell serviceNameCell = row.createCell(0);
//	                Cell serviceResourceCell = row.createCell(1);
//	                Cell spendCell = row.createCell(2);
//
//	                serviceNameCell.setCellValue(matcher.group(1).trim());
//	                serviceResourceCell.setCellValue(matcher.group(3).trim());
//	                spendCell.setCellValue(matcher.group(5).trim());
//
//	                // Adjust columns width
//	                sheet.autoSizeColumn(0);
//	                sheet.autoSizeColumn(1);
//	                sheet.autoSizeColumn(2);
//	            }
//	        }
//
//	        // Write to ByteArrayOutputStream
//	        ByteArrayOutputStream out = new ByteArrayOutputStream();
//	        workbook.write(out);
//	        workbook.close();
//
//	        // Return the Excel file as a byte array
//	        return ResponseEntity.ok()
//	                .header("Content-Disposition", "attachment; filename=\"converted.xlsx\"")
//	                .body(out.toByteArray());
//
//	    } catch (IOException e) {
//	        e.printStackTrace();
//	        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
//	    }
//	}

}


