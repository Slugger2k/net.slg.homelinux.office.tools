/**
 * @author vzmuellc
 * @category ProcessBarAbstraction
 * @since 27.05.2009
 * @version 0.5
 * 
 */
package net.slg.homelinux.office.annotations.ui;


import org.apache.log4j.Logger;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.swt.widgets.Display;

public abstract class AbstractProcessBarTask {
	
	protected IProgressMonitor monitor;
	private Display display;
	private int maxValue;
	private String title;
	private Logger logger = Logger.getLogger(AbstractProcessBarTask.class);
	
	
	/**
	 * @param display DisplayInstanz
	 * @param title Text ueber dem Fortschrittsbalken
	 * @param maxValue Anzahl der Task
	 */
	public AbstractProcessBarTask(Display display, String title, int maxValue) {
		this.display = display;
		this.title = title;
		this.maxValue = maxValue;
	}
	
	/**
	 * Muss zuerst ausgefuehrt werden um den Title 
	 * und die Anzahl der Tasks zu setzen
	 * @param title
	 * @param maxValue
	 */
	private void initProcess(final int maxValue){
		if (monitor != null) {
			display.syncExec(new Runnable() {
				public void run() {
					monitor.beginTask(title, maxValue);
				}
			});
		}
	}
	
	/**
	 * zeit unter dem Fortschrittsbalken ein Label 
	 * mit den aktuellen Prozess an
	 * @param subProcessLabel
	 */
	protected void subTask(final String subProcessLabel){
		if (monitor != null) {
			display.asyncExec(new Runnable() {
				public void run() {
					monitor.subTask(subProcessLabel);
				}
			});
		}
	}
	
	/**
	 * wird von aussen aufgerufen (UIUtil)
	 * @param monitor
	 * @throws InterruptedException 
	 */
	public void run(IProgressMonitor monitor) throws InterruptedException{
		this.monitor = monitor;
		if (monitor != null){
			initProcess(maxValue);
			
			execute();

			endProccess();
		} else {
			logger.error("IProgressMonitor darf nicht null sein" );
		}
	}
	
	/**
	 * Zaehlt den Fortschrittsbalken asyncron um 1 hoch
	 */
	public void incementBar(){
		display.asyncExec(new Runnable() {
			public void run() {
				monitor.worked(1);
			}
		});	
	}
	
	/**
	 * Beendet den Dialog in einen Asyncronen Thread
	 */
	private void endProccess(){
		if (monitor != null) {
			display.asyncExec(new Runnable() {
				public void run() {
					monitor.done();
				}
			});
		}
	}

	/**
	 * In diese Methode wird die Geschaeftslogik geschreiben, 
	 * die hinter dem Fortschrittsbalken ausgefuehrt wird. 
	 * Es muss die Methode incementBar() aufgerufen werden, 
	 * damit der Fortschritt angezeigt und hochgezaehlt werden kann. 
	 * Die Methode subTask() kann Optional ausgefuehrt werden um den
	 * aktuellen Prozess anzuzeigen (Label unter dem Fortschrittsbalken)
	 * Um den Prozess zu Canceln muss folgendes in dieser Methode stehen: 
	 * 	if (monitor.isCanceled())
	 *	    throw new InterruptedException("Cancelled");
	 * 
	 * In dieser Methode ist darauf zu achten, dass keine Widget direkt angesprochen werden,
	 * da sonst ein Invald Thread Access entsteht. Am besten alle Daten aus dem Widget extrahieren
	 * und seperat in diese Methode geben. Muss ein Widget direkt angesprochen werden muss 
	 * der Display.syncExec verwendet werden. Fehlermeldungen und Dialog betrifft dies ebenfalls.
	 * 
	 * @throws InterruptedException 
	 */
	public abstract void execute() throws InterruptedException;
}
