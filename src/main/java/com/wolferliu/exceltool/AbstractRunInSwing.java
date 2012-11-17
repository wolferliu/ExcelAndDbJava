package com.wolferliu.exceltool;

import javax.swing.*;
import java.lang.reflect.InvocationTargetException;

/**
 * Created by IntelliJ IDEA.
 * User: liubing
 * Date: 2/11/11
 * Time: 11:11 AM
 * Abstract class for invoiceAndWait
 */
public abstract class AbstractRunInSwing {
	protected void runInSwing(Runnable runnable){
		try {
			SwingUtilities.invokeAndWait(runnable);
		}
		catch (InterruptedException e) {
			e.printStackTrace();
		}
		catch (InvocationTargetException e) {
			e.printStackTrace();
		}
	}
}
