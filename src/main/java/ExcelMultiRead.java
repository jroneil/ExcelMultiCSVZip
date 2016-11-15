import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
public class ExcelMultiRead {
 

	    /**
	     * @param args the command line arguments
	     */
	    public static void main(String[] args) {
	    	boolean debugOn=false;
	    	String sourceDir=args[0];
	    	if(args.length==4){
	    	 debugOn=true;
	    	}
	    	System.out.println(debugOn);
	    	 File folder = new File(sourceDir);
	    	 
	    	 File[] listOfFiles = folder.listFiles();

	   	  for (int i = 0; i < listOfFiles.length; i++) {
	   	    File file = listOfFiles[i];
	   	    if(file.exists()){
	   	    Convertfile(file,args[1],args[2],debugOn);
	   	    }
	   	  }
	       
	    }
	    
	    
	    public static void Convertfile(File inputFile,String targetDir,String zipDir, boolean debugOn){
	    	InputStream inp = null;
	    	byte[] buffer = new byte[1024];
	    	 try {
	    		 print("inputFile.getName()"+inputFile.getName(),debugOn);
		            inp = new FileInputStream(inputFile);
		            print("file read",debugOn);  
		            Workbook wb = WorkbookFactory.create(inp);
		            String filename=getFileName(inputFile);
		            print("targetDir="+targetDir,debugOn);
	    	    		    	    	print("<wb.getNumberOfSheets()="+wb.getNumberOfSheets(),debugOn);
	    	       FileOutputStream zipOut = new FileOutputStream(zipDir+"\\"+filename+".zip");
	    	       print("zipOut=="+zipOut,debugOn);
	    	       ZipOutputStream zos = new ZipOutputStream(zipOut);
		            for(int i=0;i<wb.getNumberOfSheets();i++) {
		                print(wb.getSheetAt(i).getSheetName(),debugOn);
		                File fout = new File(targetDir+"\\"+filename+"_"+wb.getSheetAt(i).getSheetName()+".csv");
		    	    	 print("fout="+fout.getName(),debugOn);
		    	    	FileOutputStream fos = new FileOutputStream(fout);
		    	    	OutputStreamWriter osw = new OutputStreamWriter(fos);

		                echoAsCSV(wb.getSheetAt(i),osw,debugOn);
		                osw.close();
		                ZipEntry ze= new ZipEntry(fout.getName());
		                zos.putNextEntry(ze);
		        		FileInputStream in = new FileInputStream(fout.getAbsolutePath());

		        		int len;
		        		while ((len = in.read(buffer)) > 0) {
		        			zos.write(buffer, 0, len);
		        		}

		        		in.close();
		        		zos.closeEntry();

		        		//remember close it
		        		
		            }
		            zos.close();
		            print("file closed",debugOn);
		        } catch (InvalidFormatException ex) {
		            Logger.getLogger(ExcelMultiRead.class.getName()).log(Level.SEVERE, null, ex);
		        } catch (FileNotFoundException ex) {
		            Logger.getLogger(ExcelMultiRead.class.getName()).log(Level.SEVERE, null, ex);
		        } catch (IOException ex) {
		            Logger.getLogger(ExcelMultiRead.class.getName()).log(Level.SEVERE, null, ex);
		        } finally {
		            try {
		                inp.close();
		            } catch (IOException ex) {
		                Logger.getLogger(ExcelMultiRead.class.getName()).log(Level.SEVERE, null, ex);
		            }
		        }
	    	
	    	
	    }
	    public static void echoAsCSV(Sheet sheet,OutputStreamWriter osw, boolean debugOn)throws IOException {
	    	 print("sheet.getSheetName()"+sheet.getSheetName(),debugOn);
	         Row row = null;
	         print("sheet.getLastRowNum()="+sheet.getLastRowNum(),debugOn);
	        for (int i = 0; i <=sheet.getLastRowNum(); i++) {
	            row = sheet.getRow(i);
	              if(row!=null){ 
	            for (int j = 0; j < row.getLastCellNum(); j++) {
	                String line="\"" + row.getCell(j) + "\";";
	                
	                osw.write(line);
	   //             print("line="+line,debugOn);
	            }
	           }
	            osw.write("");
	        }
	     
	     
	    }
	    
	 
	    
	    private static String getFileName(File file) {
	        String fileName = file.getName();
	        if(fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
	        return fileName.substring(0,fileName.lastIndexOf("."));
	        else return "";
	    }
	    
	    private static String getFileExtension(File file) {
	        String fileName = file.getName();
	        if(fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
	        return fileName.substring(fileName.lastIndexOf(".")+1);
	        else return "";
	    }
	    
	    private static void print(String str, boolean debugOn){
	     	if(debugOn){
	    		System.out.println(str);
	    		
	    	}
	    
	    }
	}