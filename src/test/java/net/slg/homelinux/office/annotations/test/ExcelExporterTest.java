package net.slg.homelinux.office.annotations.test;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import junit.framework.Assert;
import net.slg.homelinux.office.annotations.ExcelUtil;
import net.slg.homelinux.office.annotations.ExportUtil;
import net.slg.homelinux.office.annotations.MapExcelExporter;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Before;
import org.junit.Test;


public class ExcelExporterTest {

	private List<Map<String,Object>> table;
	
	private static Logger logger = Logger.getLogger(ExcelExporterTest.class);
	
	public ExcelExporterTest() {
	}

	@Before
	public void createMockMap(){
		this.table = new ArrayList<Map<String,Object>>(10);
		
		Map<String,Object> satz = new HashMap<String, Object>();
		satz.put("Spalte1 ColumnName", new Date());
		satz.put("Spalte3 ColumnName", "Spalte 3 Value Zeile 1");
		satz.put("Spalte2 ColumnName", "Spalte 2 Value Zeile 1");
		this.table.add(satz);
		
		Map<String,Object> satz1 = new HashMap<String, Object>();
		satz1.put("Spalte2 ColumnName", Boolean.valueOf(true));
		satz1.put("Spalte1 ColumnName", "Spalte 1 Value Zeile 2");
		satz1.put("Spalte3 ColumnName", "Spalte 3 Value Zeile 2");
		this.table.add(satz1);
		
		Map<String,Object> satz2 = new HashMap<String, Object>();
		satz2.put("Spalte1 ColumnName", "Spalte 1 Value Zeile 3");
		satz2.put("Spalte3 ColumnName", new java.sql.Date(0));
		satz2.put("Spalte2 ColumnName", "Spalte 2 Value Zeile 3");
		this.table.add(satz2);
		
		Map<String,Object> satz3 = new HashMap<String, Object>();
		satz3.put("Spalte1 ColumnName", "Spalte 1 Value Zeile 4");
		satz3.put("Spalte2 ColumnName", "Spalte 2 Value Zeile 4");
		satz3.put("Spalte3 ColumnName", "Spalte 3 Value Zeile 4");
		this.table.add(satz3);
	}
	
	@Test
	public void exportExcel(){
		MapExcelExporter exporter = new MapExcelExporter();
		TreeMap<Integer, String> columnMap = exporter.findColumns(this.table);
		Assert.assertNotNull(columnMap);
		logger.debug(columnMap);
		HSSFWorkbook mapData = exporter.createXLSMapData(this.table, columnMap);
		Assert.assertNotNull(mapData);
		File excelFile = exporter.createExcelFile(mapData);
		logger.debug(excelFile.getAbsolutePath());
		Assert.assertNotNull(excelFile);
	}
	
	@Test
	public void exportExcelUtilWithMapNoSeparateColumnTitles() {
		File file = ExcelUtil.doExcelMapExportAndGetFile(ExportUtil.MODE_HEADLESS, this.table);
		Assert.assertNotNull(file);
		Assert.assertFalse(file.isDirectory());
		Assert.assertTrue(file.canRead());
	}
	
	@Test
	public void exportExcelUtilWithMapWithSeparateColumnTitles1() {
		// Test mit mehr alternativen Spaltentiteln als Spalten in der Datenquelle 
		String[] columnTitles = new String[] { "AAA", "BBB", "CCC", "DDD", "EEE" };
		File file = ExcelUtil.doExcelMapExportAndGetFile(ExportUtil.MODE_HEADLESS, this.table, columnTitles);
		Assert.assertNotNull(file);
		Assert.assertFalse(file.isDirectory());
		Assert.assertTrue(file.canRead());
	}

	@Test
	public void exportExcelUtilWithMapWithSeparateColumnTitles2() {
		// Test mit weniger alternativen Spaltentiteln als Spalten in der Datenquelle 
		String[] columnTitles = new String[] { "AAA", "BBB" };
		File file = ExcelUtil.doExcelMapExportAndGetFile(ExportUtil.MODE_HEADLESS, this.table, columnTitles);
		Assert.assertNotNull(file);
		Assert.assertFalse(file.isDirectory());
		Assert.assertTrue(file.canRead());
	}


}