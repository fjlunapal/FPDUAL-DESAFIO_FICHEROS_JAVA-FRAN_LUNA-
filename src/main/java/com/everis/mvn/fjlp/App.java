package com.everis.mvn.fjlp;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RichTextString;

public class App {
	public static void main(String[] args) {
		escribirExcel();

		leerExcel();

		System.out.println("Ejemplo Finalizado.");
	}

	public static void escribirExcel() {
		try {
			// Se crea el libro Excel
			HSSFWorkbook wb = new HSSFWorkbook();

			// Se crea una nueva hoja dentro del libro
			HSSFSheet sheet = wb.createSheet("HojaEjemplo");

			// Se crea una fila dentro de la hoja
			HSSFRow row = sheet.createRow((short) 0);
			HSSFRow row2 = sheet.createRow((short) 1);
			HSSFRow row3 = sheet.createRow((short) 2);
			HSSFRow row4 = sheet.createRow((short) 3);
			HSSFRow row5 = sheet.createRow((short) 4);
			HSSFRow row6 = sheet.createRow((short) 5);
			HSSFRow row7 = sheet.createRow((short) 6);
			HSSFRow row8 = sheet.createRow((short) 7);
			HSSFRow row9 = sheet.createRow((short) 8);
			HSSFRow row10 = sheet.createRow((short) 9);

			// Creamos celdas de varios tipos
			int a = 1;
			int b = 1;
			for (int i = 1; i <= 100; i++) {
				int mult = a * b;

				if (b == 1)
					row.createCell((short) a - 1).setCellValue(a + "*" + b + "=" + mult);
				else if (b == 2)
					row2.createCell((short) a - 1).setCellValue(a + "*" + b + "=" + mult);
				else if (b == 3)
					row3.createCell((short) a - 1).setCellValue(a + "*" + b + "=" + mult);
				else if (b == 4)
					row4.createCell((short) a - 1).setCellValue(a + "*" + b + "=" + mult);
				else if (b == 5)
					row5.createCell((short) a - 1).setCellValue(a + "*" + b + "=" + mult);
				else if (b == 6)
					row6.createCell((short) a - 1).setCellValue(a + "*" + b + "=" + mult);
				else if (b == 7)
					row7.createCell((short) a - 1).setCellValue(a + "*" + b + "=" + mult);
				else if (b == 8)
					row8.createCell((short) a - 1).setCellValue(a + "*" + b + "=" + mult);
				else if (b == 9)
					row9.createCell((short) a - 1).setCellValue(a + "*" + b + "=" + mult);
				else if (b == 10)
					row10.createCell((short) a - 1).setCellValue(a + "*" + b + "=" + mult);
				a++;
				if (a == 11) {
					a = 1;
					b += 1;
				}

			}
			// Escribimos los resultados a un fichero Excel
			FileOutputStream fileOut = new FileOutputStream("Tablas_Multiplicar.xls");

			wb.write(fileOut);
			fileOut.close();
		} catch (IOException e) {
			System.out.println("Error al escribir el fichero.");
		}
	}

	public static void leerExcel() {
		try {
			// Se abre el fichero Excel
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("Tablas_Multiplicar.xls"));

			// Se obtiene el libro Excel
			HSSFWorkbook wb = new HSSFWorkbook(fs);

			// Se obtiene la primera hoja
			HSSFSheet sheet = wb.getSheetAt(0);

			// Se leen las tablas de multiplicar
			for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
				// Se obtiene la primera fila de la hoja
				HSSFRow row = sheet.getRow(rowNum);
				for (int j = 0; j <= row.getLastCellNum(); j++) {
					// Se obtiene la celda i-esima
					HSSFCell cell = row.getCell((short) j);

					// Si la celda leida no está vacía
					if (cell != null) {
						// Se imprime en pantalla la celda según su tipo
						switch (cell.getCellType()) {
						case NUMERIC:
							System.out.println("Número: " + cell.getNumericCellValue());
							break;
						case STRING:
							System.out.println("String: " + cell.getStringCellValue());
							break;
						case BOOLEAN:
							System.out.println("Boolean: " + cell.getBooleanCellValue());
							break;
						default:
							System.out.println("Default: " + cell.getDateCellValue());
							break;
						}
					}
				}

			}
		} catch (IOException ex) {
			System.out.println("Error al leer el fichero.");
		}
	}
}