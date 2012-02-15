package com.railsware;

import jargs.gnu.CmdLineParser;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import ooo.connector.BootstrapSocketConnector;

import au.com.bytecode.opencsv.CSVWriter;

import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.uno.Exception;
import com.sun.star.uno.XComponentContext;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextTablesSupplier;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.CloseVetoException;

public class LibreOfficeCSVGenerator {
	public enum DocFormat { 
		DOC, 
		DOCX, 
		PDF
	};
	private static String defaultOOPath;
	private static XComponentContext xRemoteContext;
	private static XMultiComponentFactory xRemoteServiceManager;
	private static XComponentLoader xComponentLoader;
	
	private static void printUsage() {
        System.err.println(
"Usage: LibreOfficeCSVGenerator [{-v,--verbose}] [{-e,--office} office exec file] [{-t,--template} document template]\n" +
"                               [{-c,--csv} csv file] [{-f,--format} output format] [{-o, --output} output folder]\n" + 
"                               [{-i,--index} field from csv file for output filenames]");
    }

	/**
	 * @param args
	 * @throws IOException 
	 */
	public static void main(String[] args) {
		//command line parser
		CmdLineParser parser = new CmdLineParser();
        CmdLineParser.Option verbose = parser.addBooleanOption('v', "verbose");
        CmdLineParser.Option officeExec = parser.addStringOption('e', "office");
        CmdLineParser.Option docTemplate = parser.addStringOption('t', "template");
        CmdLineParser.Option csvFile = parser.addStringOption('c', "csv");
        CmdLineParser.Option outputFolder = parser.addStringOption('o', "output");
        CmdLineParser.Option outputFormat = parser.addStringOption('f', "format");
        CmdLineParser.Option fileNameIndex = parser.addIntegerOption('i', "index");
        
        try {
            parser.parse(args);
        }
        catch ( CmdLineParser.OptionException e ) {
            System.err.println(e.getMessage());
            printUsage();
            System.exit(2);
        }
        
        //get options
        
        String os = System.getProperty("os.name").toLowerCase();
        if (os.indexOf("win") >= 0){
        	defaultOOPath = "C:\\\\Program Files\\LibreOffice\\soffice.exe";
        } else if (os.indexOf("mac") >= 0){
        	defaultOOPath = "/Applications/LibreOffice.app/Contents/MacOS/soffice";
        } else {
        	defaultOOPath = "/usr/bin/soffice";
        }
        Boolean verboseValue = (Boolean)parser.getOptionValue(verbose, Boolean.FALSE);
        String oooExeFolder = (String)parser.getOptionValue(officeExec, defaultOOPath);
        String templateFile = (String)parser.getOptionValue(docTemplate);
        if (null == templateFile){
        	System.err.println("ERROR: You should provide template file by --template option");
        	System.exit(2);
        }
        String csvFileValue = (String)parser.getOptionValue(csvFile);
        if (null == csvFileValue){
        	System.err.println("ERROR: You should provide css file by --csv option");
        	System.exit(2);
        }
        String outputFolderValue = (String)parser.getOptionValue(outputFolder);
        if (null == outputFolderValue){
        	System.err.println("ERROR: You should provide output folder by --folder option");
        	System.exit(2);
        }
        String outputFormatValue = (String)parser.getOptionValue(outputFormat, "doc");
        DocFormat docFormat = null;
        if (outputFormatValue.compareToIgnoreCase("doc") == 0){
        	docFormat = DocFormat.DOC;
        } else if (outputFormatValue.compareToIgnoreCase("docx") == 0) {
        	docFormat = DocFormat.DOCX;
        } else if (outputFormatValue.compareToIgnoreCase("pdf") == 0){
        	docFormat = DocFormat.PDF;
        } else {
        	System.err.println("ERROR: You provide invalid document format. Support: doc, docx, pdf, but not: " + outputFormatValue);
        	System.exit(2);
        }
        Integer fileNameIndexValue = (Integer)parser.getOptionValue(fileNameIndex, new Integer(0));
        
        
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
        
        findTextsInCSV(templateFile, csvFileValue, outputFolderValue, fileNameIndexValue, docFormat, verboseValue);
	}
	
	static private void findTextsInCSV(String templateFile, String csvFile, 
			String outputFolder, Integer fileNameIndex, 
			DocFormat docFormat, Boolean verboseMode){
		try {
			
			String files;
	        File folder = new File(outputFolder);
	        File[] listOfFiles = folder.listFiles();
	        CSVWriter writer = new CSVWriter(new FileWriter(outputFolder + File.separator + "out" + File.separator + "data.csv"), ',');
	        writer.writeNext(new String[] { "Файл:", "Сума:" , "Платник:", "Місце проживання:", "Отримувач:", "Код:", "Розрахунковий рахунок:", "МФО банку:", "Призначення платежу:", "Призначення платежу 2:" });
	        System.out.println("Files: " + listOfFiles.length);
	        for (int i = 0; i < listOfFiles.length; i++) 
	        {

		         if (listOfFiles[i].isFile() && !listOfFiles[i].getName().startsWith(".")) 
		         {
		        	 files = listOfFiles[i].getName();
		         } else {
		        	 continue;
		         }
		         templateFile = outputFolder + File.separator + files;
				
				//open template
				ArrayList<PropertyValue> props = new ArrayList<PropertyValue>(); 
		        PropertyValue p = null; 
		        if (templateFile != null) { 
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
		        PropertyValue[] properties = new PropertyValue[props.size()]; 
		        props.toArray(properties);
		        
		        if (verboseMode){
		        	System.out.println( 
                        "LibreOfficeCSVGenerator: Create the document based on template " + templateFile + "."); 
		        }
		        
		        String templateFileURL = filePathToURL(templateFile); 
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
		        
		        // Create a Search Descriptor 
		        XTextDocument xTextDocument = (XTextDocument)UnoRuntime.queryInterface(XTextDocument.class, document);
		        
		        
		        // first query the XTextTablesSupplier interface from our document
		        XTextTablesSupplier xTablesSupplier = (XTextTablesSupplier) UnoRuntime.queryInterface(
		            XTextTablesSupplier.class, xTextDocument );
		        // get the tables collection
		        XNameAccess xNamedTables = xTablesSupplier.getTextTables();
		        
		        // now query the XIndexAccess from the tables collection
		        XIndexAccess xIndexedTables = (XIndexAccess) UnoRuntime.queryInterface(
		            XIndexAccess.class, xNamedTables);
		        
		        // get the tables
//		        for (int j = 0; j < xIndexedTables.getCount(); j++) {
		            Object table = xIndexedTables.getByIndex(0);
		           
		            XCellRange xCellRange = (XCellRange) UnoRuntime.queryInterface(
		                    XCellRange.class, table );
		           XCell xCell = xCellRange.getCellByPosition( 1, 1 );
		           
		           XTextRange xTextRange = (XTextRange) UnoRuntime.queryInterface(
		        	         XTextRange.class, xCell );
		           String mainString = xTextRange.getString();
		           if (mainString.length() > 0){
		        	   String[] tmpString = mainString.split("\n");
		        	   for (int k=0;k<tmpString.length;k++){
		        		   //System.out.println("data:" + tmpString[k]);
		        	   }
		           
			           List<String> entries_csv = new ArrayList<String>();
			           entries_csv.add(files);
			           entries_csv.add(tmpString[2]);
			           entries_csv.add(tmpString[4]);
			           entries_csv.add(tmpString[6]);
			           entries_csv.add(tmpString[8].replaceAll("Назва:", "").trim());
			           
			           String tmpStr = "";
			           for (int nk=12;nk < 20;nk++){
			        	   tmpStr += tmpString[nk];
			           }
			           entries_csv.add(tmpStr);
			           tmpStr = "";
			           for (int nk=21;nk < 35;nk++){
			        	   tmpStr += tmpString[nk];
			           }
			           entries_csv.add(tmpStr);
			           tmpStr = "";
			           for (int nk=35;nk < 41;nk++){
			        	   tmpStr += tmpString[nk];
			           }
			           entries_csv.add(tmpStr);
			           String secAddr = "";
			           for (int nk=0;nk < tmpString.length;nk++){
			        	   if (tmpString[nk].compareToIgnoreCase("Призначення платежу:") == 0){
			        		   secAddr = "";
			        		   int nk2 = nk + 1;
			        		   while(tmpString[nk2].compareToIgnoreCase("Платник:") != 0 || nk2 < 4){
			        			   secAddr += tmpString[nk2];
			        			   nk2++;
			        		   }
			        	   }
			           }
			           entries_csv.add(secAddr);
			           
			           // second
			           XCell xCell2 = xCellRange.getCellByPosition( 1, 0 );
			           
			           XTextRange xTextRange2 = (XTextRange) UnoRuntime.queryInterface(
			        	         XTextRange.class, xCell2 );
			           String mainString2 = xTextRange2.getString();
			           String[] tmpString2 = mainString2.split("\n");
			           
			           secAddr = "";
			           for (int nk=0;nk < tmpString2.length;nk++){
			        	   if (tmpString2[nk].compareToIgnoreCase("Призначення платежу:") == 0){
			        		   secAddr = "";
			        		   int nk2 = nk + 1;
			        		   while(tmpString2[nk2].compareToIgnoreCase("Платник:") != 0 || nk2 < 4){
			        			   secAddr += tmpString2[nk2];
			        			   nk2++;
			        		   }
			        	   }
			           }
			           entries_csv.add(secAddr);
			           
			           writer.writeNext(entries_csv.toArray(new String[entries_csv.size()]));
		           }
//		        } 
		        
				//close doc
		        
				com.sun.star.util.XCloseable xCloseable = 
	            	      (com.sun.star.util.XCloseable)UnoRuntime.queryInterface( 
	            	        com.sun.star.util.XCloseable.class,model);
				xCloseable.close(true);
				
			}
	        
	        writer.close();
		
		} catch (IllegalArgumentException e) {
			System.err.println("ERROR: Error opening document");
			e.printStackTrace();
		} catch (IndexOutOfBoundsException e) {
			// TODO Auto-generated catch block
			System.err.println("ERROR: Exception open or save doc");
			e.printStackTrace();
		} catch (WrappedTargetException e) {
			System.err.println("ERROR: Exception open or save doc");
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (CloseVetoException e) {
			System.err.println("ERROR: Coudn't close document");
			e.printStackTrace();
		} catch (com.sun.star.io.IOException e) {
			System.err.println("ERROR: Exception open or save doc");
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.err.println("ERROR: Exception open or save doc");
			e.printStackTrace();
		} finally {
	        System.out.println("Work finished!");
			System.exit(0);
		}
	}
/*	
	static private void generateDocByCSV(String templateFile, String csvFile, 
			String outputFolder, Integer fileNameIndex, 
			DocFormat docFormat, Boolean verboseMode){
		try {
			
			CSVReader reader = new CSVReader(new FileReader(csvFile));
			String [] nextLine;
			String[] headerLine = reader.readNext();
			while ((nextLine = reader.readNext()) != null) {
				if (null == nextLine[fileNameIndex]){
	            	continue;
				}
				
				//open template
				ArrayList<PropertyValue> props = new ArrayList<PropertyValue>(); 
		        PropertyValue p = null; 
		        if (templateFile != null) { 
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
		        PropertyValue[] properties = new PropertyValue[props.size()]; 
		        props.toArray(properties);
		        
		        if (verboseMode){
		        	System.out.println( 
                        "LibreOfficeCSVGenerator: Create the document based on template " + templateFile + "."); 
		        }
		        
		        String templateFileURL = filePathToURL(templateFile); 
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
		        
		        // Create a Search Descriptor 
		        XTextDocument xTextDocument = (XTextDocument)UnoRuntime.queryInterface(XTextDocument.class, document);
		        
		        XReplaceDescriptor xReplaceDescr = null;
		        XReplaceable xReplaceable = null;
		        xReplaceable = (XReplaceable)UnoRuntime.queryInterface(XReplaceable.class, xTextDocument);
		
		        xReplaceDescr = (XReplaceDescriptor)xReplaceable.createReplaceDescriptor();
		        
		        // replace
		        for (int i=0;i<nextLine.length;i++){
	            	if (headerLine[i].length() > 0){
		            	xReplaceDescr.setSearchString("{{" + headerLine[i].trim() + "}}");
		                xReplaceDescr.setReplaceString(nextLine[i]);
		                xReplaceable.replaceAll(xReplaceDescr);
		                if (verboseMode){
		                	System.out.println( 
		                        "LibreOfficeCSVGenerator: Find '{{" + headerLine[i] + 
		                        "}}' and replace " + nextLine[i] + ".");
		                }
	            	}
	            }
		        
		        //save file
		        switch(docFormat){
		        	case DOC:
				        String outDocFile = outputFolder + File.separator + nextLine[fileNameIndex] + ".doc";
				        if (verboseMode){
				        	System.out.println( 
		                        "LibreOfficeCSVGenerator: Save the document " + outDocFile + "."); 
				        }
				        String saveDocFileURL = filePathToURL(outDocFile); 
			            XStorable storableDoc = (XStorable) UnoRuntime.queryInterface( 
			                    XStorable.class, document); 
			            PropertyValue[] docProperties = new PropertyValue[1]; 
			            PropertyValue docPValue = new PropertyValue(); 
			            docPValue.Name = "FilterName"; 
			            docPValue.Value = "MS Word 97";
			            //p2.Value = "MS Word 2007 XML";
			            docProperties[0] = docPValue; 
			            storableDoc.storeAsURL(saveDocFileURL, docProperties);
						break;
		        	case DOCX:
		        		String outDocxFile = outputFolder + File.separator + nextLine[fileNameIndex] + ".docx";
				        if (verboseMode){
				        	System.out.println( 
		                        "LibreOfficeCSVGenerator: Save the document " + outDocxFile + "."); 
				        }
				        String saveDocxFileURL = filePathToURL(outDocxFile); 
			            XStorable storableDocx = (XStorable) UnoRuntime.queryInterface( 
			                    XStorable.class, document); 
			            PropertyValue[] docxProperties = new PropertyValue[1]; 
			            PropertyValue docxPValue = new PropertyValue(); 
			            docxPValue.Name = "FilterName"; 
			            docxPValue.Value = "MS Word 2007 XML";
			            docxProperties[0] = docxPValue; 
			            storableDocx.storeAsURL(saveDocxFileURL, docxProperties);
						break;
		        	case PDF:
		        		String outPdfFile = outputFolder + File.separator + nextLine[fileNameIndex] + ".pdf";
		        		if (verboseMode){
				        	System.out.println( 
		                        "LibreOfficeCSVGenerator: Save the document " + outPdfFile + "."); 
				        }
		        		String savePdfFileURL = filePathToURL(outPdfFile);
		                XStorable storablePdf = (XStorable) UnoRuntime.queryInterface( 
		                        XStorable.class, document);
		                PropertyValue[] pdfProperties = new PropertyValue[4];
		                pdfProperties[0] = new PropertyValue();
		                pdfProperties[0].Name = "FilterName";
		                pdfProperties[0].Value = "writer_pdf_Export";

		                pdfProperties[1] = new PropertyValue();
		                pdfProperties[1].Name = "Pages";
		                pdfProperties[1].Value = "All";

		                pdfProperties[2] = new PropertyValue();
		                pdfProperties[2].Name = "Overwrite";
		                pdfProperties[2].Value = Boolean.TRUE; 
		                
		                pdfProperties[3] = new PropertyValue();
		                pdfProperties[3].Name = "CompressionMode";
		                pdfProperties[3].Value = "1"; 
		                
		                storablePdf.storeToURL(savePdfFileURL, pdfProperties);
		                break;
		        }
				//close doc
				com.sun.star.util.XCloseable xCloseable = 
	            	      (com.sun.star.util.XCloseable)UnoRuntime.queryInterface( 
	            	        com.sun.star.util.XCloseable.class,model);
				xCloseable.close(true);
				
			}
		
		} catch (IllegalArgumentException e) {
			System.err.println("ERROR: Error opening document");
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			System.err.println("ERROR: CSV file not found");
			e.printStackTrace();
		} catch (CloseVetoException e) {
			System.err.println("ERROR: Coudn't close document");
			e.printStackTrace();
		} catch (com.sun.star.io.IOException e) {
			System.err.println("ERROR: Exception open or save doc");
			e.printStackTrace();
		} catch (IOException e) {
			System.err.println("ERROR: Exception with doc");
			e.printStackTrace();
		} finally {
	        System.out.println("Work finished!");
			System.exit(0);
		}
	}
*/	
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
