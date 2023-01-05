package com.losom.excelfilepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.losom.apachepoiconcept.Student;

public class ExcelCommon_POI {
	private static String[] columns = { "Id", "Name", "Student Code", "Gmail", "Score", "Address" };
	private static List<Student> students = new ArrayList<Student>();

	public static void createExcelFile() {
		students.add(new Student(1L, "Som", "HCl1", "somhcl@gmail.com", 9, "Vietnam"));
		students.add(new Student(2L, "Anee", "HCl1", "aneehcl@gmail.com", 10, "India"));
		students.add(new Student(3L, "James", "HCl1", "jampshcl@gmail.com", 9, "America"));
		students.add(new Student(4L, "Victory", "HCl1", "victoryhcl@gmail.com", 9, "Vietnam"));
		students.add(new Student(5L, "Sandeep", "HCl1", "sandeephcl@gmail.com", 9, "India"));
		students.add(new Student(6L, "Yen", "HCl1", "yenhcl@gmail.com", 8, "Vietnam"));
		students.add(new Student(7L, "Duc", "HCl1", "duchcl@gmail.com", 9, "Laos"));
		students.add(new Student(8L, "Nam", "HCl1", "namhcl@gmail.com", 8, "Vietnam"));
		students.add(new Student(9L, "John", "HCl1", "Johnhcl@gmail.com", 10, "Vietnam"));
		students.add(new Student(10L, "Quynh", "HCl1", "quynhhcl@gmail.com", 9, "China"));

		try (XSSFWorkbook newWorkBook = new XSSFWorkbook()) {
			XSSFSheet newSheet = newWorkBook.createSheet("Student-File");
			XSSFRow headerRow = newSheet.createRow(0);
			for (int i = 0; i < columns.length; i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(columns[i]);
				newSheet.autoSizeColumn(i);
			}
			XSSFFont fontHeader = newWorkBook.createFont();
			fontHeader.setBold(true);
			fontHeader.setFontName("Courier New");
			fontHeader.setColor(IndexedColors.DARK_RED.index);

			int rowNumber = 1;
			for (Student student : students) {
				Row row = newSheet.createRow(rowNumber++);
				row.createCell(0, CellType.NUMERIC).setCellValue(student.getId());
				row.createCell(1, CellType.STRING).setCellValue(student.getName());
				row.createCell(2, CellType.STRING).setCellValue(student.getStudentCode());
				row.createCell(3, CellType.STRING).setCellValue(student.getEmail());
				row.createCell(4, CellType.NUMERIC).setCellValue(student.getScore());
				row.createCell(5, CellType.STRING).setCellValue(student.getAddress());
			}
			FileOutputStream fileOut = null;
			try {
				fileOut = new FileOutputStream("D:\\OJTJavaHCL/Students.xlsx");
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}
			try {
				newWorkBook.write(fileOut);
			} catch (IOException e) {
				e.printStackTrace();
			}
			try {
				fileOut.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void readAndWriteFile() {
		File fileStudent = new File("D:\\OJTJavaHCL/Students.xlsx");
		FileInputStream newFileIn = null;
		try {
			newFileIn = new FileInputStream(fileStudent);
			System.out.println("Excel file written successfully");
		} catch (FileNotFoundException e) {
			System.out.println("Fild not found!");
		}

		// Creating a workbook from an excel file .xlsx;
		XSSFWorkbook newWorkBook = null;
		try {
			newWorkBook = new XSSFWorkbook(newFileIn);
		} catch (IOException e1) {
			System.out.println("Created a workbook failed");
		}

		// Getting the sheet at index zero
		XSSFSheet newSheet = newWorkBook.getSheetAt(0);
		Iterator<Row> rowIterator = newSheet.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = (Cell) cellIterator.next();
				switch (cell.getCellType()) {
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue() + "\t");
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue() + "\t");
					break;
				case STRING:
					System.out.print(cell.getStringCellValue() + "\t");
					break;
				case BLANK:
					System.out.print("");
					break;
				case ERROR:
					System.out.print(cell.getErrorCellValue() + "\t");
				case FORMULA:
					System.out.print(cell.getCellFormula() + "\t");
				default:
					System.out.println("");
					break;
				}
			}
			System.out.println();

		}
		try {
			newFileIn.close();
		} catch (IOException e) {
			System.out.println("Closed excel file failed!");
		}
		// Write old excel file to new excel file
		try {
			FileOutputStream fileOut = new FileOutputStream("D:\\OJTJavaHCL/StudentCopy.xlsx");
			System.out.println("Retrieved excel file successfullly!");
			try {
				fileOut.close();
			} catch (IOException e) {
				System.out.println("Closed excel file failed");
			}
		} catch (FileNotFoundException e) {
			System.out.println("Retrieved file excel failed");
		}

	}

}
