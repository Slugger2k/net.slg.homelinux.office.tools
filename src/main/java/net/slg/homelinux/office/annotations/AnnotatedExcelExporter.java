/**
 * @author VZMUELLC
 * @since 18.12.2008
 * @version 1.1.1
 * 
 * 
 * ChangeLog:
 * ####################################################################################
 * # v 1.1.1  -  04.01.2010 - Dateiname und Pfad der Execldatei kann angegeben werden #
 * ####################################################################################
 * 
 */
package net.slg.homelinux.office.annotations;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.SortedMap;
import java.util.TreeMap;
import java.util.Vector;
import java.util.Map.Entry;

import net.slg.homelinux.office.annotations.ui.AbstractProcessBarTask;
import net.slg.homelinux.office.annotations.ui.ExcelExportProcessBarTask;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.eclipse.swt.widgets.Display;



public class AnnotatedExcelExporter {

	private static Logger logger = Logger.getLogger(AnnotatedExcelExporter.class);
	private Class<?> clazz;
	private List<?> entitaetList;

	private Map<String, String> columnNameMap;
	private SortedMap<Integer, String> orderMap;
	private Map<String, String> dateFormat;
	private Map<String, String> suppressedByMap;
	private Map<String, Object> dataMap;
	private List<String> suppressList;
	private Map<String, String> subItemMap;
	private SortedMap<Integer, String> ordercopy;
	private String dateiname = "";
	//Pfad braucht ein nachgestelltest //
	private String pfadStr = "";


	public AnnotatedExcelExporter(List<?> list) {
		startExport(list, null);
	}
	
	public AnnotatedExcelExporter(List<?> list, String[] suppress){
		if (suppress != null){
			suppressList = Arrays.asList(suppress);
		}
		startExport(list, suppressList);
	}
	
	public AnnotatedExcelExporter(List<?> list, String[] suppress, String dateiName){
		if (suppress != null){
			suppressList = Arrays.asList(suppress);
		}
		dateiname = dateiName;
		startExport(list, suppressList);
	}
	
	public AnnotatedExcelExporter(List<?> list, String[] suppress, String pfad, String dateiName){
		if (suppress != null){
			suppressList = Arrays.asList(suppress);
		}
		dateiname = dateiName;
		pfadStr = pfad;
		startExport(list, suppressList);
	}
	
	private void startExport(List<?> list, List<String> suppressList){
		logger.info("Excel Export gestartet");
		
		columnNameMap = new HashMap<String, String>(5);
		orderMap = new TreeMap<Integer, String>();
		dateFormat = new HashMap<String, String>(5);
		suppressedByMap = new HashMap<String, String>(2);
		dataMap = new HashMap<String, Object>(5);
		subItemMap = new HashMap<String, String>(2);
		ordercopy = new TreeMap<Integer, String>();
		
		if (list.size() != 0) {
			this.entitaetList = list;
			this.clazz = list.get(0).getClass();
			findAnnotations();
			createXLSMapData();
		} else {
			logger.error("Entitaetenliste leer!");
		}
	}
	
	private void findAnnotations() {
		int order = 0;
		for (Method m : this.clazz.getDeclaredMethods()) {
			if (m.isAnnotationPresent(ExcelExport.class)) {
				if (m.getName().startsWith("get") || m.getName().startsWith("is")) {
					ExcelExport excelExport = m.getAnnotation(ExcelExport.class);
					
					try {
						if (excelExport.columnName().equals("") && m.getName().startsWith("get")) {
							columnNameMap.put(m.getName(), m.getName().replaceFirst("get", ""));
						} else if (excelExport.columnName().equals("") && m.getName().startsWith("is")) {
							columnNameMap.put(m.getName(), m.getName().replaceFirst("is", ""));
						} else {
							columnNameMap.put(m.getName(), excelExport.columnName());						
						}
						
						dateFormat.put(m.getName(), excelExport.dateFormat());
						
						if (excelExport.suppressedBy() != null){
							suppressedByMap.put(m.getName(), excelExport.suppressedBy());						
						}				
						if (excelExport.order() == -1){
							orderMap.put(order, m.getName());
							order++;							
						} else {
							orderMap.put(excelExport.order(), m.getName());
						}
						if (suppressList != null){
							for (String suppress : suppressList) {
								if (excelExport.suppressedBy().equals(suppress)){
									columnNameMap.remove(m.getName());
									dateFormat.remove(m.getName());
									orderMap.remove(excelExport.order());
									suppressedByMap.remove(m.getName());
								}							
							}
						}
						if (excelExport.subItem().length() > 0){
							subItemMap.put(m.getName(), excelExport.subItem());							
						}
					} catch (Exception e){
						logger.error(e.getMessage(), e);
					}
				}
			}
		}
		
		int counter = 0;
		for (Integer zahl : orderMap.keySet()) {
			ordercopy.put(counter, orderMap.get(zahl));
			counter++;
		}
	}

	private void createXLSMapData() {
		
		HSSFWorkbook workBook = new HSSFWorkbook();
		Vector<HSSFSheet> sheets = new Vector<HSSFSheet>();
		HSSFSheet defaultsheet = workBook.createSheet(this.clazz.getSimpleName());
		String workBookTitle = "Table";
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
	    for ( Iterator<Entry<String, String>> columnName = columnNameMap.entrySet().iterator(); columnName.hasNext(); ) {
	    	if (ordercopy.get(spalte) != null){
    			HSSFCell myCell = row.createCell(spalte);
    			myCell.setCellValue(new HSSFRichTextString(columnNameMap.get(ordercopy.get(spalte))));
    			myCell.setCellStyle(style);	    		
	    		spalte++;	    			    		
	    	} 
	        columnName.next();
	    }


	    AbstractProcessBarTask processBarTask = new ExcelExportProcessBarTask(Display.getDefault(), "Exportiere nach Excel...", this.entitaetList.size(), this.entitaetList, subItemMap, dataMap, sheets, columnNameMap, ordercopy, dateFormat);
	    ExcelUtil.openProgressDialog(processBarTask, false);
	   
		for (int col = 0; col < columnNameMap.size(); col++){
			defaultsheet.autoSizeColumn((short) col);
		}

		openExcelTable(workBook, workBookTitle);
	}

	
	private void openExcelTable(HSSFWorkbook workBook, String workBookTitle) {
		if (dateiname.equals("")){
			try {
				File file = File.createTempFile(workBookTitle, "export");
				File tempFile = new File(file.getAbsolutePath() + ".xls");
				file.delete();
				tempFile.deleteOnExit();
				FileOutputStream fileOut = new FileOutputStream(tempFile);
				workBook.write(fileOut);
				fileOut.close();
				Runtime.getRuntime().exec("cmd /c start excel.exe " + tempFile.getAbsolutePath());
				logger.info("Excel Export erfolgreich abgeschlossen.");
			} catch (FileNotFoundException e) {
				logger.error(e.getMessage(), e);
			} catch (IOException e) {
				logger.error(e.getMessage(), e);
			}
		} else {
			try {
				if(!pfadStr.contains(":\\")){
					pfadStr = "c:\\" + pfadStr;
				}
				
				if (!pfadStr.endsWith("\\")){
					pfadStr += "\\";
				}
				
				File file = new File(pfadStr + dateiname + ".xls");
				FileOutputStream fileOut = new FileOutputStream(file);
				workBook.write(fileOut);
				fileOut.close();
				Runtime.getRuntime().exec("cmd /c start excel.exe " + file.getAbsolutePath());
				logger.info("Excel Export erfolgreich abgeschlossen.");
			} catch (FileNotFoundException e) {
				logger.error(e.getMessage(), e);
			} catch (IOException e) {
				logger.error(e.getMessage(), e);
			}
		}
	}
	
}
