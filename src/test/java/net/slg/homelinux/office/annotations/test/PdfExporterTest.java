package net.slg.homelinux.office.annotations.test;

import java.io.File;
import java.sql.Types;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import junit.framework.Assert;
import net.slg.homelinux.office.annotations.ExportUtil;
import net.slg.homelinux.office.annotations.PdfUtil;

import org.junit.Before;
import org.junit.Test;

public class PdfExporterTest {

private List<Map<String,Object>> table;

	@Before
	public void createMapsForTests(){
		this.table = new ArrayList<Map<String,Object>>(10);
		
		Map<String,Object> satz = new HashMap<String, Object>();
		satz.put("Spalte1", new Date());
		satz.put("Spalte3", "Spalte 3 Value Zeile 1");
		satz.put("Spalte2", "Spalte 2 Value Zeile 1");
		this.table.add(satz);
		
		Map<String,Object> satz1 = new HashMap<String, Object>();
		satz1.put("Spalte2", Boolean.valueOf(true));
		satz1.put("Spalte1", "Spalte 1 Value Zeile 2");
		satz1.put("Spalte3", "Spalte 3 Value Zeile 2");
		this.table.add(satz1);
		
		Map<String,Object> satz2 = new HashMap<String, Object>();
		satz2.put("Spalte1", "Spalte 1 Value Zeile 3");
		satz2.put("Spalte3", new java.sql.Date(0));
		satz2.put("Spalte2", "Spalte 2 Value Zeile 3");
		this.table.add(satz2);
		
		Map<String,Object> satz3 = new HashMap<String, Object>();
		satz3.put("Spalte1", "Spalte 1 Value Zeile 4");
		satz3.put("Spalte2", "Spalte 2 Value Zeile 4");
		satz3.put("Spalte3", "Spalte 3 Value Zeile 4");
		this.table.add(satz3);
	}

	private List<Map<String,Object>> createDummyData(int numColumns, int numRows) {
		List<Map<String,Object>> data = new ArrayList<Map<String,Object>>(numColumns);
		
		for (int row = 0; row < numRows; row++) {
			Map<String,Object> satz = new HashMap<String, Object>();
			for (int col = 0; col < numColumns+1; col++) {
				satz.put("Spalte" + col, "Value Zeile " + row + ", Spalte " + col);
			}
			data.add(satz);
		}
		
		return data;
	}
	
	
	@Test
	public void testMapPdfExporterNoSeparateColumnTitles() {
		
		File file = PdfUtil.doPdfMapExportAndGetFile(ExportUtil.MODE_HEADLESS, createDummyData(3, 250),
				new String[] {"Spalte1","Spalte2","Spalte3" }, 
				new String[] {"Title1","Title2" },
				new int[] {Types.VARCHAR , Types.VARCHAR },
				PdfUtil.PORTRAIT, 
				"Liste TEST", 
				"Dies ist ein Bottom-Text");
		
		Assert.assertNotNull(file);
		Assert.assertFalse(file.isDirectory());
		Assert.assertTrue(file.canRead());
		
	}
	
	@Test
	public void testMapPdfExporterSeparateColumnTitles() {
		
		PdfUtil.doPdfMapExportAndView(ExportUtil.MODE_HEADLESS, createDummyData(3, 250),
				new String[] {"Spalte1","Spalte2","Spalte3" }, 
				new String[] {"Title1","Title2" },
				new int[] {Types.VARCHAR , Types.VARCHAR },
				PdfUtil.LANDSCAPE, 
				"Liste TEST im Querformat", 
				"Dies ist ein Bottom-Text im Querformat");
		
		Assert.assertTrue(true);
		
	}
	
	

	
}
