package com.wolferliu.exceltool;

import com.wolferliu.exceltool.config.ConnectionParameter;
import com.wolferliu.exceltool.config.ParameterSet;
import com.wolferliu.exceltool.config.ParameterSetLoader;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.io.FilenameFilter;

/**
 * Export UI.
 * User: liubing
 * Date: 2/10/11
 * Time: 3:38 PM
 */
public class ExportData extends AbstractExcelTool implements AsyncUIChangable {
	JPanel parameterPanel;
	JPanel sqlPanel;
	JPanel logPanel;
	JPanel buttonPanel;

	//Sql & Log
	JTextArea textSql;

	private final static int MIN_WINDOWS_WIDTH = 800;
	private final static int MIN_WINDOWS_HEIGHT = 600;
	private JButton bExport;
	private JButton bClear;
	private JButton bCopy;
	private JButton bExit;

	public static final String APP_NAME="Excel And DB(Export)";

	public static void main(String[] args) throws ClassNotFoundException, UnsupportedLookAndFeelException, IllegalAccessException, InstantiationException {
		if (IS_MAC) {
			// take the menu bar off the jframe
			System.setProperty("apple.laf.useScreenMenuBar", "true");

			// set the name of the application menu item
			System.setProperty("com.apple.mrj.application.apple.menu.about.name", APP_NAME);
		}

		UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		new ExportData().open();
	}

	private void open() {
		loadConfig();
		mainFrame = new JFrame(APP_NAME);
		mainPanel = new JPanel(new GridBagLayout());

		prepareParameterPanel();

		prepareSqlLogPanel();

		prepareButtonPanel();

		prepareMainPanel();

		WindowListener wndCloser = new WindowAdapter() {
			public void windowClosing(WindowEvent e) {
				if (!processingFlag)
					System.exit(0);
			}


		};
		mainFrame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		mainFrame.addWindowListener(wndCloser);
		ComponentListener compListener = new ComponentAdapter() {
			@Override
			public void componentResized(ComponentEvent e) {
				Component c = (Component) e.getSource();
				if (c instanceof JFrame) {
					Dimension newSize = c.getSize();
					int newWidth = (int) newSize.getWidth();
					int newHeight = (int) newSize.getHeight();
					if (newWidth < MIN_WINDOWS_WIDTH) {
						newWidth = MIN_WINDOWS_WIDTH;
					}
					if (newHeight < MIN_WINDOWS_HEIGHT) {
						newHeight = MIN_WINDOWS_HEIGHT;
					}
					c.setSize(newWidth, newHeight);
				}
			}
		};
		mainFrame.addComponentListener(compListener);

		mainFrame.getContentPane().add(mainPanel);
		mainFrame.setSize(MIN_WINDOWS_WIDTH, MIN_WINDOWS_HEIGHT);
//		mainFrame.pack();

		fillWithParameter(currentProfileName);

		mainFrame.setVisible(true);
	}

	private void prepareParameterPanel() {
		GridBagConstraints c = new GridBagConstraints();
		c.insets = new Insets(2, 2, 2, 2);
		c.anchor = GridBagConstraints.WEST;
		parameterPanel = new JPanel(new GridBagLayout());
		parameterPanel.setBorder(BorderFactory.createTitledBorder("ConnectionParameter"));

		int x,y;


		//Add config name
		x=0;
		y=0;
		c.weightx = labelWeight;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		parameterPanel.add(new JLabel("Profile:"), c);
		c.weightx = textWeight;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		profileModel=new DefaultComboBoxModel(parameterSet.getParameterNameArray());
		comboProfile = new JComboBox(
				profileModel
		);
		comboProfile.setEditable(true);
		if(profileModel.getSize()>0){
			currentProfileName=String.valueOf(profileModel.getSelectedItem());
		}

		comboProfile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent actionEvent) {
//				System.out.println(actionEvent.getActionCommand());
				if("comboBoxChanged".equals(actionEvent.getActionCommand())){
					String pName=String.valueOf(comboProfile.getSelectedItem());
					if(parameterSet.isNameExist(pName)){
						fillWithParameter(pName);
						currentProfileName=pName;
					}
				}
				else if("comboBoxEdited".equals(actionEvent.getActionCommand())){
					String pName=String.valueOf(comboProfile.getSelectedItem());
					if(pName.length()==0){
						System.out.println("currentProfileName1 = " + currentProfileName);
						profileModel.removeElement(currentProfileName);
						parameterSet.removeParameter(currentProfileName);
						currentProfileName=String.valueOf(profileModel.getSelectedItem());
						if(currentProfileName==null){
							currentProfileName="";
						}
						System.out.println("currentProfileName2 = " + currentProfileName);
					}
				}
			}
		});
		parameterPanel.add(comboProfile, c);

		//Add Config Button
		c.weightx = labelWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		parameterPanel.add(new JLabel(""), c);
		c.weightx = textWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		c.anchor = GridBagConstraints.EAST;
		buttonSystemConfig = new JButton("System Config");
		parameterPanel.add(buttonSystemConfig, c);
		c.anchor = GridBagConstraints.WEST;
		ActionListener configListener = new ActionListener() {
			public void actionPerformed(ActionEvent ae) {
				SystemConfigDialog sd=new SystemConfigDialog(mainFrame, parameterSet);
				sd.setVisible(true);
			}
		};
		buttonSystemConfig.addActionListener(configListener);


		x=0;
		y++;

		//Add Driver
		c.weightx = labelWeight;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		parameterPanel.add(new JLabel("Driver:"), c);
		c.weightx = textWeight;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		comboDriver = new JComboBox(parameterSet.getDriverArray());
		comboDriver.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent actionEvent) {
				if(textUrl.getText().length()==0){
					String driverName=String.valueOf(comboDriver.getSelectedItem());
					textUrl.setText(parameterSet.getDefaultUrl(driverName));
				}
			}
		});
		parameterPanel.add(comboDriver, c);

		//Add File
		c.weightx = labelWeight;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		parameterPanel.add(new JLabel("Excel File:"), c);
		c.weightx = textWeight;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		textFileName = new JTextField("");
		parameterPanel.add(textFileName, c);


		x=0;
		y++;
		//Add Url
		c.weightx = labelWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		parameterPanel.add(new JLabel("URL:"), c);
		c.weightx = textWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		textUrl = new JTextField();
		parameterPanel.add(textUrl, c);

		//Add Select Button
		c.weightx = labelWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		parameterPanel.add(new JLabel(""), c);
		c.weightx = textWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		c.anchor = GridBagConstraints.EAST;
		buttonSelectFile = new JButton("Select File");
		parameterPanel.add(buttonSelectFile, c);
		c.anchor = GridBagConstraints.WEST;
		ActionListener selectListener = new ActionListener() {
			public void actionPerformed(ActionEvent ae) {
				if(IS_MAC){
					selectDistFileAWT();
				}
				else{
					selectDistFileSwing();
				}
			}
		};
		buttonSelectFile.addActionListener(selectListener);


		x=0;
		y++;
		//Add username
		c.weightx = labelWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		parameterPanel.add(new JLabel("User Name:"), c);
		c.weightx = textWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		textUsername = new JTextField();
		parameterPanel.add(textUsername, c);

		//Add title line
		c.weightx = labelWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		parameterPanel.add(new JLabel("Title Line:"), c);
		c.weightx = textWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		textTitleLine = new JTextField();
		parameterPanel.add(textTitleLine, c);

		x=0;
		y++;
		//Add password
		c.weightx = labelWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		parameterPanel.add(new JLabel("Password:"), c);
		c.weightx = textWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		textPassword = new JTextField();
		parameterPanel.add(textPassword, c);

		//Add data line
		c.weightx = labelWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		parameterPanel.add(new JLabel("Data Line:"), c);
		c.weightx = textWeight;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		textDataLine = new JTextField();
		parameterPanel.add(textDataLine, c);
	}


	private void prepareSqlLogPanel() {
		GridBagConstraints c = new GridBagConstraints();
		//sql panel
		sqlPanel = new JPanel(new GridBagLayout());
		sqlPanel.setBorder(BorderFactory.createTitledBorder("Please enter correct SQL to execute, don't include ; "));

		c.weightx = 1;
		c.weighty = 1;
		c.fill = GridBagConstraints.BOTH;
		textSql = new JTextArea();
		textSql.setLineWrap(true);
		sqlPanel.add(new JScrollPane(textSql), c);

		//Log
		logPanel = new JPanel(new GridBagLayout());
		logPanel.setBorder(BorderFactory.createTitledBorder("LOG"));

		c.weightx = 1;
		c.weighty = 1;
		c.fill = GridBagConstraints.BOTH;
		textLog = new JTextArea();
		textLog.setLineWrap(true);
		textLog.setEditable(false);
		logPanel.add(new JScrollPane(textLog), c);

	}

	private void prepareButtonPanel() {
		GridBagConstraints c = new GridBagConstraints();
		buttonPanel = new JPanel(new GridBagLayout());
		c.weightx = 0;
		c.weighty = 0;
		c.insets = new Insets(0, 5, 0, 5);
		c.fill = GridBagConstraints.NONE;
		bExport = new JButton("Export");
		buttonPanel.add(bExport, c);

		c.weightx = 1;
		c.weighty = 0;
		c.fill = GridBagConstraints.HORIZONTAL;
		buttonPanel.add(new JLabel(""), c);

		c.weightx = 0;
		c.weighty = 0;
		c.fill = GridBagConstraints.NONE;
		c.insets = new Insets(0, 5, 0, 5);
		bClear = new JButton("Clear");
		buttonPanel.add(bClear, c);

		c.weightx = 0;
		c.weighty = 0;
		c.fill = GridBagConstraints.NONE;
		c.insets = new Insets(0, 5, 0, 5);
		bCopy = new JButton("Copy");
		buttonPanel.add(bCopy, c);

		c.weightx = 0;
		c.weighty = 0;
		c.fill = GridBagConstraints.NONE;
		c.insets = new Insets(0, 5, 0, 5);
		bExit = new JButton("Exit");
		buttonPanel.add(bExit, c);

		ActionListener exportListener = new ActionListener() {
			public void actionPerformed(ActionEvent ae) {
				if(textFileName.getText().length()==0){
					JOptionPane.showMessageDialog(null, "Please select file to export!", "Error", JOptionPane.ERROR_MESSAGE);
					return;
				}
				ExportExecutor ee = new ExportExecutor();

				String pName=String.valueOf(comboProfile.getSelectedItem());

				ConnectionParameter cp=new ConnectionParameter();
				cp.setDriver(String.valueOf(comboDriver.getSelectedItem()));
				cp.setUrl(textUrl.getText());
				cp.setUsername(textUsername.getText());
				cp.setPassword(textPassword.getText());
				cp.setDataLine(Integer.parseInt(textDataLine.getText()));
				cp.setTitleLine(Integer.parseInt(textTitleLine.getText()));
				ee.setConnectionParameter(cp);
				ee.setSystemParameter(parameterSet.getSystemParameter());

				ee.setFileName(textFileName.getText());
				ee.setSQL(textSql.getText());
				try {
					if(pName.length()>0){
						parameterSet.addParameter(pName, cp);
						configLoader.saveConfig(parameterSet);
					}
					
					ee.setUiObject(ExportData.this);
					//SwingUtilities.invokeAndWait(ee);
					new Thread(ee).start();
//					startExec();
//					Thread.sleep(10000);
//					endExec();
				}
				catch (Exception e) {
					textLog.append(e.getMessage() + "\n");
				}
			}
		};
		bExport.addActionListener(exportListener);

		ActionListener clearListener = new ActionListener() {
			public void actionPerformed(ActionEvent ae) {
				textLog.setText("");

			}
		};
		bClear.addActionListener(clearListener);

		ActionListener copyListener = new ActionListener() {
			public void actionPerformed(ActionEvent ae) {
				textLog.selectAll();
				textLog.copy();
				textLog.select(0, 0);
			}
		};
		bCopy.addActionListener(copyListener);

		ActionListener exitListener = new ActionListener() {
			public void actionPerformed(ActionEvent ae) {
				System.exit(0);
			}
		};
		bExit.addActionListener(exitListener);
	}

	private void prepareMainPanel() {
		GridBagConstraints c = new GridBagConstraints();
		c.insets = new Insets(5, 5, 5, 5);
		c.weightx = 1.0;
		c.weighty = 0;
		c.gridwidth = GridBagConstraints.REMAINDER;
		c.fill = GridBagConstraints.HORIZONTAL;
		mainPanel.add(parameterPanel, c);

		c.weightx = 1.0;
		c.weighty = 1;
		c.gridwidth = GridBagConstraints.REMAINDER;
		c.fill = GridBagConstraints.BOTH;
		mainPanel.add(sqlPanel, c);

		c.weightx = 1.0;
		c.weighty = 1;
		c.gridwidth = GridBagConstraints.REMAINDER;
		c.fill = GridBagConstraints.BOTH;
		mainPanel.add(logPanel, c);

		c.weightx = 1.0;
		c.weighty = 0;
		c.gridwidth = GridBagConstraints.REMAINDER;
		c.fill = GridBagConstraints.HORIZONTAL;
		mainPanel.add(buttonPanel, c);
	}

	private boolean processingFlag;

	protected void toggleProcessing(boolean flag) {
		processingFlag = flag;
		bExport.setEnabled(!flag);
		bClear.setEnabled(!flag);
		bCopy.setEnabled(!flag);
		bExit.setEnabled(!flag);
	}

	protected void selectDistFileAWT(){
		FileDialog fd=new FileDialog(mainFrame, "Select Excel File", FileDialog.SAVE);
		fd.setFilenameFilter(new FilenameFilter(){
			public boolean accept(File dir, String name) {
				return name != null && name.endsWith("*.xls");
			}
		});
		fd.setAlwaysOnTop(true);
		fd.setVisible(true);
		String fs=fd.getFile();

		String fileName = (fs != null) ? fd.getDirectory()+fs : "";
		if (fileName.length() > 0 && !fileName.endsWith("xls")) {
			fileName += ".xls";
		}
		textFileName.setText(fileName);
	}

	private void selectDistFileSwing() {
		JFileChooser chooser = new JFileChooser();
		FileFilter filter = new FileNameExtensionFilter("Excel 2003 files (*.xls)", "xls");
		chooser.addChoosableFileFilter(filter);
		int option = chooser.showSaveDialog(mainPanel);
		if (option == JFileChooser.APPROVE_OPTION) {
			String fileName = (chooser.getSelectedFile() != null) ? chooser.getSelectedFile().getPath() : "";
			if (fileName.length() > 0 && !fileName.endsWith("xls")) {
				fileName += ".xls";
			}
			textFileName.setText(fileName);
		}
	}


}
