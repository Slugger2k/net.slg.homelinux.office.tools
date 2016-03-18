package net.slg.homelinux.office.annotations.ui;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.Date;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;
import java.util.SortedMap;
import java.util.Vector;

import net.slg.homelinux.office.annotations.ExcelExport;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.eclipse.swt.widgets.Display;

public class ExcelExportProcessBarTask extends AbstractProcessBarTask {

	private final List<?> entitaetList;
	private final Map<String, String> subItemMap;
	private final Map<String, Object> dataMap;
	private Logger logger = Logger.getLogger(ExcelExportProcessBarTask.class);
	private final Vector<HSSFSheet> sheets;
	private final Map<String, String> columnNameMap;
	private final SortedMap<Integer, String> ordercopy;
	private final Map<String, String> dateFormat;
	
	
	public ExcelExportProcessBarTask(Display display, String title, int maxValue, List<?> entitaetListe, Map<String, String> subItemMap, Map<String, Object> dataMap, Vector<HSSFSheet> sheets, Map<String, String> columnNameMap, SortedMap<Integer, String> ordercopy, Map<String, String> dateFormat) {
		super(display, title, maxValue);
		this.entitaetList = entitaetListe;
		this.subItemMap = subItemMap;
		this.dataMap = dataMap;
		this.sheets = sheets;
		this.columnNameMap = columnNameMap;
		this.ordercopy = ordercopy;
		this.dateFormat = dateFormat;
	}

	@Override
	public void execute() throws InterruptedException {
		int rowNumber = 1;
		
		for (Object obj : entitaetList) {
			Class<?> clasz = obj.getClass();
			for (Method m : clasz.getDeclaredMethods()) {
				if (m.isAnnotationPresent(ExcelExport.class)) {
					if (m.getName().startsWith("get") || m.getName().startsWith("is")) {
						
						String returnTypeString = m.getReturnType().getName();
						String subItemName = "";
						
						try {
							if (subItemMap.get(m.getName()) != "" && subItemMap.get(m.getName()) != null){
								if (m.getName().startsWith("get")){
									subItemName = m.getName().replaceFirst("get", "");
								} else {
									subItemName = m.getName().replaceFirst("is", "");
								}
								Class<?> clas;
								Object o = m.invoke(obj);
								clas = o.getClass();
								for (Method me : clas.getDeclaredMethods()) {
									if (me.getName().toLowerCase().endsWith(subItemMap.get(m.getName()).toLowerCase()) && ( me.getName().startsWith("get") || me.getName().startsWith("is")) ){
										
										if (returnTypeString.equals("boolean") || returnTypeString.equals("java.lang.Boolean")) {
											dataMap.put(m.getName(), new Boolean("" + me.invoke(o)));
										} else if (returnTypeString.equals("byte") || returnTypeString.equals("java.lang.Byte")) {
											dataMap.put(m.getName(), (Byte) me.invoke(o));
										} else if (returnTypeString.equals("java.sql.Date")) {
											dataMap.put(m.getName(), (Date) me.invoke(o)); 
										} else if (returnTypeString.equals("double") || returnTypeString.equals("java.lang.Double")) {
											dataMap.put(m.getName(), (Double) me.invoke(o));
										} else if (returnTypeString.equals("float") || returnTypeString.equals("java.lang.Float")) {
											dataMap.put(m.getName(), (String) me.invoke(o));
										} else if (returnTypeString.equals("int") || returnTypeString.equals("java.lang.Integer")) {
											dataMap.put(m.getName(),(Integer) me.invoke(o));
										} else if (returnTypeString.equals("short") || returnTypeString.equals("java.lang.Short")) {
											dataMap.put(m.getName(), (Double) me.invoke(o));
										} else if (returnTypeString.equals("long") || returnTypeString.equals("java.lang.Long")) {
											dataMap.put(m.getName(), (Double) me.invoke(o));
										} else if (returnTypeString.equals("java.lang.String")) {
											dataMap.put(m.getName(), (String) me.invoke(o));
										} else if (returnTypeString.equals("java.sql.Time")) {
											dataMap.put(m.getName(), (Time) me.invoke(o));
										} else if (returnTypeString.equals("java.sql.Timestamp")) {
											dataMap.put(m.getName(), (Timestamp) me.invoke(o));
										} else {								
											dataMap.put(m.getName(), "" + me.invoke(o));
										}								
									}
								}
							}
						} catch (IllegalArgumentException e) {
							logger.error(e.getMessage(), e);
						} catch (IllegalAccessException e) {
							logger.error(e.getMessage(), e);
						} catch (InvocationTargetException e) {
							logger.error(e.getMessage(), e);
						} catch (NullPointerException e){
							logger.error(e.getMessage(), e);
						} catch (Exception e){
							logger.error(e.getMessage(), e);
						}
						
						try {
							if (subItemName.equals("")){
								if (returnTypeString.equals("boolean") || returnTypeString.equals("java.lang.Boolean")) {
									dataMap.put(m.getName(), new Boolean("" + m.invoke(obj)));
								} else if (returnTypeString.equals("byte") || returnTypeString.equals("java.lang.Byte")) {
									dataMap.put(m.getName(), (Byte) m.invoke(obj));
								} else if (returnTypeString.equals("java.sql.Date")) {
									dataMap.put(m.getName(), (Date) m.invoke(obj)); 
								} else if (returnTypeString.equals("double") || returnTypeString.equals("java.lang.Double")) {
									dataMap.put(m.getName(), (Double) m.invoke(obj));
								} else if (returnTypeString.equals("float") || returnTypeString.equals("java.lang.Float")) {
									dataMap.put(m.getName(), (String) m.invoke(obj));
								} else if (returnTypeString.equals("int") || returnTypeString.equals("java.lang.Integer")) {
									dataMap.put(m.getName(),(Integer) m.invoke(obj));
								} else if (returnTypeString.equals("short") || returnTypeString.equals("java.lang.Short")) {
									dataMap.put(m.getName(), (Double) m.invoke(obj));
								} else if (returnTypeString.equals("long") || returnTypeString.equals("java.lang.Long")) {
									dataMap.put(m.getName(), (Double) m.invoke(obj));
								} else if (returnTypeString.equals("java.lang.String")) {
									dataMap.put(m.getName(), (String) m.invoke(obj));
								} else if (returnTypeString.equals("java.sql.Time")) {
									dataMap.put(m.getName(), (Time) m.invoke(obj));
								} else if (returnTypeString.equals("java.sql.Timestamp")) {
									dataMap.put(m.getName(), (Timestamp) m.invoke(obj));
								} else {								
									dataMap.put(m.getName(), "" + m.invoke(obj));
								}
							}
						} catch (ClassCastException e) {
							logger.error(e.getMessage(), e);
						} catch (IllegalArgumentException e) {
							logger.error(e.getMessage(), e);
						} catch (IllegalAccessException e) {
							logger.error(e.getMessage(), e);
						} catch (InvocationTargetException e) {
							logger.error(e.getMessage(), e);
						}
					}
				}
			}
			HSSFRow row = sheets.get(0).createRow((int) rowNumber );
			for (int c = 0; c < columnNameMap.size(); c++) {
				
				HSSFCell cell = row.createCell(c);
				
				try {
					if (dataMap.get(ordercopy.get(c)) instanceof java.lang.String){
						String cellWert = (String) dataMap.get(ordercopy.get(c));					
						cell.setCellValue(new HSSFRichTextString(cellWert));
						cell.setCellType(HSSFCell.CELL_TYPE_STRING);
					} else if (dataMap.get(ordercopy.get(c)) instanceof java.lang.Integer || dataMap.get(ordercopy.get(c)) instanceof Double || 
							   dataMap.get(ordercopy.get(c)) instanceof java.lang.Short || dataMap.get(ordercopy.get(c)) instanceof java.lang.Float){
						Double cellWert = Double.parseDouble("" + dataMap.get(ordercopy.get(c)) );					
						cell.setCellValue(cellWert);
						cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
					} else if (dataMap.get(ordercopy.get(c)) instanceof java.lang.Boolean){
						Boolean cellWert = Boolean.parseBoolean("" + dataMap.get(ordercopy.get(c)) );					
						cell.setCellValue(cellWert);
						cell.setCellType(HSSFCell.CELL_TYPE_BOOLEAN);
					} else if (dataMap.get(ordercopy.get(c)) instanceof java.sql.Date) {
						String date = "";
						Date sqlDate = (Date) dataMap.get(ordercopy.get(c));
						if (!dateFormat.get(ordercopy.get(c)).equals("dd.MM.yyyy") && dateFormat.get(ordercopy.get(c)) != null){
							date =  new SimpleDateFormat(dateFormat.get(ordercopy.get(c))).format(sqlDate);
						} else {
							date =  new SimpleDateFormat("dd.MM.yyyy").format(sqlDate);							
						}
						cell.setCellValue(new HSSFRichTextString(date));
						cell.setCellType(HSSFCell.CELL_TYPE_STRING);
					} else if (dataMap.get(ordercopy.get(c)) instanceof java.sql.Time){
						Time cellWert =  (Time) dataMap.get(ordercopy.get(c));	
						String time = "";
						if ((!dateFormat.get(ordercopy.get(c)).equals("HH:mm:ss") && dateFormat.get(ordercopy.get(c)) != null)){
							time =  new SimpleDateFormat(dateFormat.get(ordercopy.get(c))).format(cellWert);							
						} else {
							time = new SimpleDateFormat("HH:mm:ss").format(cellWert);
						}
						cell.setCellValue(new HSSFRichTextString(time));
						cell.setCellType(HSSFCell.CELL_TYPE_STRING);
					} else if (dataMap.get(ordercopy.get(c)) instanceof java.sql.Timestamp){
						Timestamp cellWert =  (Timestamp) dataMap.get(ordercopy.get(c));	
						String timestamp =""; 
						if (!dateFormat.get(ordercopy.get(c)).equals("dd.MM.yyyy HH:mm:ss") && dateFormat.get(ordercopy.get(c)) != null){
							timestamp =  new SimpleDateFormat(dateFormat.get(ordercopy.get(c))).format(cellWert);							
						} else {
							timestamp =  new SimpleDateFormat("dd.MM.yyyy HH:mm:ss").format(cellWert);
						}
						cell.setCellValue(new HSSFRichTextString(timestamp));
						cell.setCellType(HSSFCell.CELL_TYPE_STRING);
					}
				} catch (ClassCastException e) {
					logger.error(e.getMessage() ,e);
				} catch (NumberFormatException  e) {
					logger.error(e.getMessage() ,e);
				} catch (Exception e){
					logger.error(e.getMessage() ,e);
				}
			}
			incementBar();
			subTask("Exportiere Eintrag: "  + rowNumber);
			rowNumber++;
			dataMap.clear();
		}
		
		
	}

}
