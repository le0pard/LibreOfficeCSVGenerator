package com.railsware;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import ooo.connector.BootstrapSocketConnector;
import jargs.gnu.CmdLineParser;

import au.com.bytecode.opencsv.CSVWriter;

import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XSheetCellCursor;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.sheet.XUsedAreaCursor;
import com.sun.star.table.XCell;
import com.sun.star.text.XText;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.CloseVetoException;

public class SheetToCSV {
	
	private static String defaultOOPath;
	private static XComponentContext xRemoteContext;
	private static XMultiComponentFactory xRemoteServiceManager;
	private static XComponentLoader xComponentLoader;
	private static XSpreadsheetDocument myDoc;
	
	private static void printUsage() {
        System.err.println(
"Usage: SheetToCSV [{-f,--file}] [{-c,--office} office exec file]\n" +
"                  [{-e, --encoding}] [{-o, --output} output folder]\n");
    }

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		//command line parser
		CmdLineParser parser = new CmdLineParser();
        CmdLineParser.Option xlsFileOpt = parser.addStringOption('f', "file");
        CmdLineParser.Option officeExec = parser.addStringOption('c', "office");
        CmdLineParser.Option fileEncOpt = parser.addStringOption('e', "encoding");
        CmdLineParser.Option outputFolderOps = parser.addStringOption('o', "output");
        
        try {
            parser.parse(args);
        }
        catch ( CmdLineParser.OptionException e ) {
            System.err.println(e.getMessage());
            printUsage();
            System.exit(2);
        }
        
        String os = System.getProperty("os.name").toLowerCase();
        if (os.indexOf("win") >= 0){
        	defaultOOPath = "C:\\\\Program Files\\LibreOffice\\soffice.exe";
        } else if (os.indexOf("mac") >= 0){
        	defaultOOPath = "/Applications/LibreOffice.app/Contents/MacOS/soffice";
        } else {
        	defaultOOPath = "/usr/bin/soffice";
        }
        String oooExeFolder = (String)parser.getOptionValue(officeExec, defaultOOPath);
        String xlsFile = (String)parser.getOptionValue(xlsFileOpt);
        String outputFolder = (String)parser.getOptionValue(outputFolderOps);
        String fileEncoding = (String)parser.getOptionValue(fileEncOpt, "Unicode (UTF-8)");
        
        if (xlsFile == null || outputFolder == null){
        	System.err.println("ERROR: Could not start working. Please provide file and output dir by '--file' and '--output' options");
            System.exit(2);
        }
        
        /* begin working with document */
		try {
			xRemoteContext = BootstrapSocketConnector.bootstrap(oooExeFolder);
		} catch (BootstrapException e) {
			System.err.println("ERROR: Could not bootstrap default Office. Please provide right way to soffice by '--office' option");
			e.printStackTrace();
			System.exit(2);
		}
        if (xRemoteContext == null) {
            System.err.println("ERROR: Could not bootstrap default Office. . Please provide right way to soffice by '--office' option");
            System.exit(2);
        }
        
        xRemoteServiceManager = xRemoteContext.getServiceManager();
        Object desktop = null;
		try {
			desktop = xRemoteServiceManager.createInstanceWithContext(
			    "com.sun.star.frame.Desktop", xRemoteContext);
		} catch (Exception e) {
			System.err.println("ERROR: Could not bootstrap default Office.");
			e.printStackTrace();
			System.exit(2);
		}
        xComponentLoader = (XComponentLoader)
            UnoRuntime.queryInterface(XComponentLoader.class, desktop);
        
        //open template
		ArrayList<PropertyValue> props = new ArrayList<PropertyValue>(); 
        PropertyValue p = null; 
        if (xlsFile != null) { 
            // Enable the use of a template document. 
            p = new PropertyValue(); 
            p.Name = "AsTemplate"; 
            p.Value = new Boolean (true); 
            props.add(p); 
        }  
        // Make the document initially invisible so the user does not 
        // have to watch it being built. 
        p = new PropertyValue(); 
        p.Name = "Hidden"; 
        p.Value = new Boolean(true); 
        props.add(p); 
        
        p = new PropertyValue(); 
        p.Name = "CharacterSet";
        p.Value = fileEncoding;
        props.add(p); 
        PropertyValue[] properties = new PropertyValue[props.size()]; 
        props.toArray(properties);
        
        System.out.println( 
                "LibreOfficeCSVGenerator: Create the document based on template " + xlsFile + ".");
        
        String templateFileURL = filePathToURL(xlsFile); 
        try {
			XComponent document = xComponentLoader.loadComponentFromURL( 
			        templateFileURL,    // URL of templateFile. 
			        "_blank",           // Target frame name (_blank creates new frame). 
			        0,                  // Search flags. 
			        properties);
			
			
			 // get window and frame
	        XModel model = (XModel) UnoRuntime.queryInterface( 
                    XModel.class, document); 
            XController c = model.getCurrentController(); 
            XFrame frame = c.getFrame(); 
            XWindow window = frame.getContainerWindow();
            
            window.setEnable(false); 
            window.setVisible(false);
            
            
            myDoc = (XSpreadsheetDocument) UnoRuntime.queryInterface(
                    XSpreadsheetDocument.class, document);
            
            if (myDoc != null) {
            	int maxCols = -1;
            	String find0 = null;
                String find1 = null;
                
            	XSpreadsheets xSheets = myDoc.getSheets();
            	String[] sheetNames = xSheets.getElementNames();
            	for (int i = 0; i < sheetNames.length; i++){
					CSVWriter writer = new CSVWriter(new FileWriter(outputFolder + File.separator + sheetNames[i] + ".csv"), ',');
	                XIndexAccess oIndexSheets = (XIndexAccess) UnoRuntime
	                                .queryInterface(XIndexAccess.class, xSheets);
	                XSpreadsheet xSheet = (XSpreadsheet) UnoRuntime
	                                .queryInterface(XSpreadsheet.class,
	                                                oIndexSheets.getByIndex(i));
	                XSheetCellCursor cursor = xSheet.createCursor();
	                XUsedAreaCursor xUsedAreaCursor = UnoRuntime.queryInterface(
	                                XUsedAreaCursor.class, cursor);
	
	                xUsedAreaCursor.gotoEndOfUsedArea(true);
	                XCellRangeAddressable xCellRangeAddressable = UnoRuntime
	                                .queryInterface(XCellRangeAddressable.class,
	                                                xUsedAreaCursor);
	                int rowCount = xCellRangeAddressable.getRangeAddress().EndRow;
	                int colCount = xCellRangeAddressable.getRangeAddress().EndColumn;
	                if (maxCols != -1 && colCount > maxCols) {
	                        colCount = maxCols;
	                }
	                for (int r = 0; r <= rowCount; r++) {
	                		List<String> entries_csv = new ArrayList<String>();
	                        if (find0 != null) {
	                                if (!xSheet.getCellByPosition(0, r).getFormula()
	                                                .contains(find0)) {
	                                        continue;
	                                }
	                        }
	                        if (find1 != null) {
	                                if (!xSheet.getCellByPosition(1, r).getFormula()
	                                                .contains(find1)) {
	                                        continue;
	                                }
	                        }
	                        for (int m = 0; m <= colCount; m++) {
	                                XCell cell = xSheet.getCellByPosition(m, r);
	                                String val = cell.getFormula();
	                                if (m > 0) {
	                                    System.out.print(',');
	                                }
	                                XText text = (XText)
	                                		UnoRuntime.queryInterface(XText.class, cell);
	                                XTextRange xTextRange = (XTextRange) UnoRuntime.queryInterface(
	                                        XTextRange.class, cell );
	                                entries_csv.add(text.getString());
	                                System.out.print(text.getString());
	                        }
	                        System.out.println();
	                        if (entries_csv.size() > 1){
	        		        	//write
	        		        	writer.writeNext(entries_csv.toArray(new String[entries_csv.size()]));
	        		        }
	                }
	                writer.close();
            	}
                // XCell cell = xSheet.getCellByPosition(0, 0);
                // System.out.println(cell.getFormula());
	        } else {
	        	System.err.println("ERROR: Could not load doc.");
				System.exit(2);
	        }
            //document.dispose();
			
			com.sun.star.util.XCloseable xCloseable = 
          	      (com.sun.star.util.XCloseable)UnoRuntime.queryInterface( 
          	        com.sun.star.util.XCloseable.class, document);
			xCloseable.close(true);

			
		} catch (com.sun.star.io.IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IndexOutOfBoundsException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (CloseVetoException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	 /** Convert a file path to URL format. */ 
    static private String filePathToURL(String file) { 
        File f = new File(file); 
        StringBuffer sb = new StringBuffer("file:///"); 
        try { 
            sb.append(f.getCanonicalPath().replace('\\', '/')); 
        } catch (IOException e) { 
        } 
        return sb.toString(); 
    } 

}
