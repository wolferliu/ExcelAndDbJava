/*
 * Created on 2004-3-6
 *
 * To change the template for this generated file go to
 * Window - Preferences - Java - Code Generation - Code and Comments
 */
package com.wolferliu.exceltool;

/**
 * @author SonyMusic
 *
 * To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Generation - Code and Comments
 */
public class UIChangeRunInSwing extends AbstractRunInSwing implements Runnable{
	public static final String MODE_START="START";
	public static final String MODE_END="END";
	private String mode;
	private AsyncUIChangable uiObject;
	
	UIChangeRunInSwing(AsyncUIChangable uiObject){
		this.uiObject=uiObject;
	}
	
	public void run(){
		if(MODE_START.equalsIgnoreCase(mode)){
			uiObject.startExec();
		}
		else{
			uiObject.endExec();
		}
	}

	public void startExec() {
		this.mode=MODE_START;
		runInSwing(this);
	}

	public void endExec() {
		this.mode=MODE_END;
		runInSwing(this);
	}


}
