package net.slg.homelinux.office.annotations;

/**
 * @author VZMUELLC
 * 
 * 
 * MapExcelExporter
 * Nimmt eine List aus Maps entgegen und generiert daraus 
 * eine .xls Datei im Tmp-Verzeichnis des Betriebssystems.
 * 
 * Mapaufbau siehe ExcelExporterTest.java
 * 
 * created:  17.12.2011
 * 
 * Ablauf:
 * createExcelFile(createXLSMapData(list, findColumns(list)));
 * 
 */

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Date;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeMap;
import java.util.Vector;

import net.slg.homelinux.office.annotations.ExportUtil;
import net.slg.homelinux.office.annotations.ui.MapToExcelProcessBarTask;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.eclipse.swt.widgets.Display;

public class MapExcelExporter {
	
	private static Logger logger = Logger.getLogger(MapExcelExporter.class);
	private String workBookTitle = "Table";
	private String[] columnTitles = null;
	private int mode;

	public MapExcelExporter() {
		this.mode = ExportUtil.MODE_HEADLESS;
	}
	
	public MapExcelExporter(int mode) {
		this.mode = mode;
	}
	
	public MapExcelExporter(int mode, List<Map<String, Object>> list) {
		this.mode = mode;
		startExport(list);
	}

	public void startExport(List<Map<String, Object>> list) {
		logger.info("Excel Export gestartet");

		if (list.size() > 0) {
			createExcelFile(createXLSMapData(list, findColumns(list)));
		} else {
			logger.error("Map leer!");
		}
	}

	public File startExport(List<Map<String, Object>> list, String path){
		
		if (list.size() > 0) {
			TreeMap<Integer, String> columnMap = this.findColumns(list);
			HSSFWorkbook mapData = this.createXLSMapData(list, columnMap);
			File excelFile = this.createExcelFile(mapData);
			return excelFile;			
		} else {
			logger.error("Map leer!");
			return null;
		}
	}
	
	public TreeMap<Integer, String> findColumns(List<Map<String, Object>> entitaetList) {
		// fill columnNameMap
		int i = 0;
		Map<String, Object> satz = entitaetList.get(0);
		TreeMap<Integer, String> columnNameMap = new TreeMap<Integer, String>();

		for (Map.Entry<String, Object> entry : satz.entrySet()){
			columnNameMap.put(i, "" + entry.getKey());
			i++;
		}
		
		return columnNameMap;
	}

	public HSSFWorkbook createXLSMapData(List<Map<String, Object>> entitaetList, TreeMap<Integer, String> columnNameMap) {

		Map<Integer, Object> dataMap = new TreeMap<Integer, Object>();
		 
		HSSFWorkbook workBook = new HSSFWorkbook();
		Vector<HSSFSheet> sheets = new Vector<HSSFSheet>();
		HSSFSheet defaultsheet = workBook.createSheet("Worksheet");

		sheets.add(defaultsheet);

		// erste Zeile erstellen
		HSSFRow row = sheets.get(0).createRow((int) 0);

		// Create a new font and alter it.
		HSSFFont font = workBook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Arial");
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

		// Fonts are set into a style so create a new one to use.
		HSSFCellStyle style = workBook.createCellStyle();
		style.setFont(font);

		// Spaltennamen in der 1. Zeile einfuegen
		int spalte = 0;
		for (Entry<Integer, String> entry : columnNameMap.entrySet()) {
			HSSFCell myCell = row.createCell(spalte);
			
			String title = findColumnTitleInTitleArrayIfAvailable(spalte, entry.getValue());
			
			myCell.setCellValue(new HSSFRichTextString(title));
			myCell.setCellStyle(style);
			spalte++;
		}

	
		if (this.mode==ExportUtil.MODE_DESKTOP) {
			ExportUtil.openProgressDialog(new MapToExcelProcessBarTask(Display.getDefault(), "Exportiere nach Excel...", 
					entitaetList.size(),entitaetList, dataMap, sheets, row, columnNameMap), false);
		} else {
			int rowNumber = 1;
			for (Map<String, Object> map : entitaetList) {
				int column = 0;
				for (Map.Entry<String, Object> entry : map.entrySet()) {
					try {
						if (entry.getValue() != null){
							dataMap.put(column, entry.getValue());											
						}
						column++;
					} catch (IllegalArgumentException e) {
						logger.error(e.getMessage(), e);
					} catch (NullPointerException e) {
						logger.error(e.getMessage(), e);
					} catch (Exception e) {
						logger.error(e.getMessage(), e);
					}
				}
				row = sheets.get(0).createRow((int) rowNumber);
				for (int c = 0; c < columnNameMap.size(); c++) {
					HSSFCell cell = row.createCell(c);
					try {
						if (dataMap.get(c) instanceof BigDecimal || dataMap.get(c) instanceof Integer || dataMap.get(c) instanceof Double || dataMap.get(c) instanceof Short || dataMap.get(c) instanceof Float) {
							Double cellWert = Double.parseDouble("" + dataMap.get(c));
							cell.setCellValue(cellWert);
							cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
						} else if (dataMap.get(c) instanceof java.lang.Boolean) {
							Boolean cellWert = Boolean.parseBoolean("" + dataMap.get(c));
							cell.setCellValue(cellWert);
							cell.setCellType(HSSFCell.CELL_TYPE_BOOLEAN);
						} else if (dataMap.get(c) instanceof java.sql.Date) {
							String date = "";
							Date sqlDate = (Date) dataMap.get(c);
							date = new SimpleDateFormat("dd.MM.yyyy").format(sqlDate);
							cell.setCellValue(new HSSFRichTextString(date));
							cell.setCellType(HSSFCell.CELL_TYPE_STRING);
						} else if (dataMap.get(c) instanceof java.sql.Time) {
							Time cellWert = (Time) dataMap.get(c);
							String time = "";
							time = new SimpleDateFormat("HH:mm:ss").format(cellWert);
							cell.setCellValue(new HSSFRichTextString(time));
							cell.setCellType(HSSFCell.CELL_TYPE_STRING);
						} else if (dataMap.get(c) instanceof java.sql.Timestamp) {
							Timestamp cellWert = (Timestamp) dataMap.get(c);
							String timestamp = "";
							timestamp = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss").format(cellWert);
							cell.setCellValue(new HSSFRichTextString(timestamp));
							cell.setCellType(HSSFCell.CELL_TYPE_STRING);
						}
						else /*if (dataMap.get(c) instanceof java.lang.String)*/ {
						    String cellWert = (String) dataMap.get(c);
		                    cell.setCellValue(new HSSFRichTextString(cellWert));
		                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
		                } 
					} catch (ClassCastException e) {
						logger.error(e.getMessage(), e);
					} catch (NumberFormatException e) {
						logger.error(e.getMessage(), e);
					} catch (Exception e) {
						logger.error(e.getMessage(), e);
					}
				}
				dataMap.clear();
				rowNumber++;
			}
			
		}
		
		
		for (int col = 0; col < columnNameMap.size(); col++) {
			defaultsheet.autoSizeColumn((short) col);
		}
		
		return workBook;
	}

	public File createExcelFile(HSSFWorkbook workBook) {
		return createExcelFile(workBook, null);
	}
	
	public File createExcelFile(HSSFWorkbook workBook, String fileName) {

		File tempFile;
		try {
			if (fileName != null){
				tempFile = new File(fileName);

			} else {
				File file = File.createTempFile(this.workBookTitle, "export");
				tempFile = new File(file.getAbsolutePath() + ".xls");
				file.delete();				
			}
			FileOutputStream fileOut = new FileOutputStream(tempFile);
			workBook.write(fileOut);
			fileOut.close();
			logger.info("Excel Export erfolgreich abgeschlossen." + tempFile.getAbsolutePath());
		} catch (FileNotFoundException e) {
			logger.error(e.getMessage(), e);
			return null;
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
			return null;
		}
		return tempFile;
	}
	
	
	public void setColumnTitles(String[] columnTitles) {
		this.columnTitles = columnTitles;
	}
	
	private String findColumnTitleInTitleArrayIfAvailable(int index, String alternative) {
        try {
                return this.columnTitles[index];
        } catch (Exception e) {
                return alternative;
        }              
}
}