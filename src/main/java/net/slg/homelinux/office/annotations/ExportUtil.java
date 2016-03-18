package net.slg.homelinux.office.annotations;

import java.lang.reflect.InvocationTargetException;

import net.slg.homelinux.office.annotations.ui.AbstractProcessBarTask;

import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.eclipse.swt.widgets.Display;

public class ExportUtil {

	public final static int MODE_HEADLESS = 0;
	public final static int MODE_DESKTOP = 1;
	
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
	
}
