/**
 * @author VZMUELLC
 */

package net.slg.homelinux.office.annotations;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.Map;

import net.slg.homelinux.office.annotations.ui.AbstractProcessBarTask;

import org.apache.log4j.Logger;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.eclipse.swt.widgets.Display;


public class ExcelUtil {

	private static Logger logger = Logger.getLogger(ExcelUtil.class);
	
	public static void doExcelExport(List<?> list){
		new AnnotatedExcelExporter(list);
	}

	public static void doExcelExport(List<?> list, String[] suppress){
		new AnnotatedExcelExporter(list, suppress);
	}
	
	/**
	 * Startet den ExcelExport
	 * 
	 * @param list
	 * @param suppress  
	 * @param dateiName Der Name der Datei ohne die Erweiterung ".xls"
	 */
	public static void doExcelExport(List<?> list, String[] suppress, String dateiName){
		new AnnotatedExcelExporter(list, suppress, dateiName);
	}
	
	/**
	 * Startet den ExcelExport
	 * 
	 * @param list
	 * @param suppress
	 * @param pfad		Der Speicherort der Datei mit oder ohne Laufwerksangabe
	 * @param dateiName Der Name der Datei ohne die Erweiterung ".xls"
	 */
	public static void doExcelExport(List<?> list, String[] suppress, String pfad, String dateiName){
		new AnnotatedExcelExporter(list, suppress, pfad, dateiName);
	}
	
	public static void doExcelMapExport(int mode, List<Map<String, Object>> list){
		new MapExcelExporter(mode, list);
	}
	
	public static File doExcelMapExportAndGetFile(int mode, List<Map<String, Object>> list, String[] columnTitles){
		MapExcelExporter exporter = new MapExcelExporter(mode);
		exporter.setColumnTitles(columnTitles);
		return exporter.createExcelFile(exporter.createXLSMapData(list, exporter.findColumns(list)));
	}
		
	public static File doExcelMapExportAndGetFile(int mode,List<Map<String, Object>> list){
		MapExcelExporter exporter = new MapExcelExporter(mode);
		return exporter.createExcelFile(exporter.createXLSMapData(list, exporter.findColumns(list)));
	}
	
	public static File doExcelMapExportAndGetFile(int mode, List<Map<String, Object>> list, String fileName){
		MapExcelExporter exporter = new MapExcelExporter(mode);
		return exporter.createExcelFile(exporter.createXLSMapData(list, exporter.findColumns(list)), fileName);
	}

	
	public static void openFileExternally(File excelFile) {
		
		String os = System.getProperty("os.name").toLowerCase();
		String command = "";
		 
		if (os.indexOf("win") >= 0) {
			// on Windows we'll burst up Microsoft Excel
			command = "cmd /c start excel.exe " + excelFile.getAbsolutePath();
		}

		if (os.indexOf("mac") >= 0) {
			// on Mac we'll try to fire up... yes... Excel
			command = "open -b com.microsoft.Excel " + excelFile.getAbsolutePath();
		}
		
		if (os.indexOf("linux") >= 0) {
			// on Linux we try to open the file with libre office / calc
			command = "libreoffice --calc " + excelFile.getAbsolutePath();
		}
		
		try {
			Runtime.getRuntime().exec(command);
		} catch (IOException e) {
			logger.error("Open in external application failed! cmd=[" + command  + "] errmsg=" + e.getMessage());
		}	
		
	}
	
	/**
	 * Oeffnet einen Fortschrittsbalkendialog und fuehrt die logik aus, 
	 * die in der Implementierung von processBarTask steht.
	 * 
	 * @param processBarTask Auszufuehrende Logik
	 * @param isCancelEnabled Kann der Prozess beendet werden
	 */
	public static void openProgressDialog(final AbstractProcessBarTask processBarTask, boolean isCancelEnabled) {

		ProgressMonitorDialog dialog = new ProgressMonitorDialog(Display.getCurrent().getActiveShell());
		
		try {
			IRunnableWithProgress op = new IRunnableWithProgress() {
				public void run(IProgressMonitor monitor) throws InvocationTargetException, InterruptedException{
					processBarTask.run(monitor);
				}	
			}; 
			dialog.run(true, isCancelEnabled, op);
		} catch (InvocationTargetException e1) {
			System.out.println("invocation");
		} catch (InterruptedException e1) {
			System.out.println("interrupted");
		} finally {
			dialog.close();
		}

	}
	
	/**
	 * Oeffnet einen Fortschrittsbalkendialog und fuehrt die logik aus, 
	 * die in der Implementierung von processBarTask steht. Gibt String mit Abbruchbedingung zurueck !
	 * 
	 * @param processBarTask Auszufuehrende Logik
	 * @param isCancelEnabled Kann der Prozess beendet werden
	 */
	public static String openProgressDialog2(final AbstractProcessBarTask processBarTask, boolean isCancelEnabled) {

		String erg = "OK";
		ProgressMonitorDialog dialog = new ProgressMonitorDialog(Display.getCurrent().getActiveShell());
		
		try {
			IRunnableWithProgress op = new IRunnableWithProgress() {
				public void run(IProgressMonitor monitor) throws InvocationTargetException, InterruptedException{
					processBarTask.run(monitor);
				}	
			}; 
			dialog.run(true, isCancelEnabled, op);
		} catch (Exception e) {
			erg = e.getMessage();
		} finally {
			dialog.close();
		}
		return erg;

	}
}