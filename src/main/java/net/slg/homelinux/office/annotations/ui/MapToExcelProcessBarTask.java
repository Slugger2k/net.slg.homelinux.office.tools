package net.slg.homelinux.office.annotations.ui;

import java.math.BigDecimal;
import java.sql.Date;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.Vector;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.eclipse.swt.widgets.Display;

public class MapToExcelProcessBarTask extends AbstractProcessBarTask {

	private List<Map<String,Object>> entitaetList;
	private Map<Integer, Object> dataMap;
	private Vector<HSSFSheet> sheets;
	private HSSFRow row;
	private TreeMap<Integer, String> columnNameMap;
	
	private static Logger logger = Logger.getLogger(MapToExcelProcessBarTask.class);
	
	public MapToExcelProcessBarTask(Display display, String title, int maxValue, List<Map<String,Object>> entitaetList,
			Map<Integer, Object> dataMap, Vector<HSSFSheet> sheets, HSSFRow row, TreeMap<Integer, String> columnNameMap ) {
		super(display, title, maxValue);
		this.entitaetList = entitaetList;
		this.dataMap = dataMap;
		this.sheets = sheets;
		this.row = row;
		this.columnNameMap = columnNameMap;
	}
	

	@Override
	public void execute() throws InterruptedException {

		
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
			incementBar();
			subTask("Exportiere Eintrag: "  + rowNumber);
			rowNumber++;
			dataMap.clear();
		}
		

	}

}
