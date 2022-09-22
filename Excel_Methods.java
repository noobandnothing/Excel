import java.time.format.DateTimeFormatter;  
import java.time.LocalDateTime;    
import java.util.ArrayList;
import java.util.Date;
import java.util.Map;
import java.util.Set;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Methods {   
   
   public static void CreateBlankExcel(String File_Path) {

		try {
		    XSSFWorkbook workbook = new XSSFWorkbook();
		    FileOutputStream out;
			out = new FileOutputStream(
			     new File(File_Path));
			workbook.write(out);
		    out.close();
		    workbook.close();
		    System.out.println(File_Path+" is written successfully");
		} catch (Exception e) {
		    System.out.println("\nE : Excel_Methods(CreateBlankExcel)");
			e.printStackTrace();
		}
	      
	   }
   
	   public static void CreateBlankExcel_L(String File_Path, ArrayList<String> Sheets) {
			try {
				XSSFWorkbook workbook = new XSSFWorkbook();
			    for (int counter = 0 ; counter < Sheets.size() ; counter++)
			    {
			       workbook.createSheet(Sheets.get(counter));
			    }
			    FileOutputStream out;
				out = new FileOutputStream(
				     new File(File_Path));
				workbook.write(out);
			    out.close();
			    workbook.close();
			    System.out.println(File_Path+" is written successfully");
			} catch (Exception e) {
				System.out.println("\nE : Excel_Methods(CreateBlankExcel_L)");
				e.printStackTrace();
			}
		}
	   
	   public static int  getSheetNum(String File_Path) {
		      File file= new File(File_Path);
	    	  int x = 0;
		      if (file.isFile() && file.exists()) { 
		    	  XSSFWorkbook workbook;
					try {
						workbook = new XSSFWorkbook(file);
						x =  workbook.getNumberOfSheets() ; 
				    	workbook.close();
					} catch (Exception e) {
						System.out.println("\nE : Excel_Methods(getSheetNum)");
						e.printStackTrace();
					}
		      }else {
		         System.out.println(File_Path+ " is not Found.");
		      }
		      
		      return x;
		}
	   
	   public static ArrayList<String>  getSheetNames(String File_Path) {
		      ArrayList<String> Sheets;
		      File file= new File(File_Path);
		      if (file.isFile() && file.exists()) { 
		    	  XSSFWorkbook workbook;
		    	  Sheets = new ArrayList<String>() ;
		    	  try {
						workbook = new XSSFWorkbook(file);
						for (int counter = 0 ; counter < workbook.getNumberOfSheets() ; counter++){
							Sheets.add(workbook.getSheetName(counter));
				        }
				    	workbook.close();
				    	return Sheets;
					} catch (Exception e) {
						System.out.println("\nE : Excel_Methods(getSheetNum)");
						e.printStackTrace();
					}
			    	
			       }else {
			          System.out.println(File_Path+ " is not Found.");
			       }
		      return null;
		}
	   
	   public static void CopyExcel(String File_Path, String File_Path_Copy) {
		      File file= new File(File_Path);
		      if (file.isFile() && file.exists()) { 
		    	  XSSFWorkbook workbook;
				try {
					workbook = new XSSFWorkbook(file);
					FileOutputStream out = new FileOutputStream(
					           new File(File_Path_Copy));
					workbook.write(out);
				    out.close();
				    workbook.close();
				} catch (Exception e) {
					System.out.println("\nE : Excel_Methods(CopyExcel)");
					e.printStackTrace();
				}
		          
	              System.out.println(File_Path+" is backuped successfully");
		       }else {
		          System.out.println(File_Path+ " is not Found.");
		       }
	   }
	   
	   public static void AddExcelSheet(String File_Path, ArrayList<String> Sheets) {
		      CopyExcel(File_Path,File_Path+"_"+datetimenow());
		      File file= new File(File_Path);
	    	  XSSFWorkbook workbook;
	    	  if (file.isFile() && file.exists()) { 
	    		  try {
	    			    ArrayList<String> Sheets_names = getSheetNames(File_Path);
		  				workbook = new XSSFWorkbook(new FileInputStream(file));
		  				for (int counter = 0 ; counter < Sheets.size() ; counter++)
		  		        {
		  					if(!Sheets_names.contains(Sheets.get(counter))) {
		  						workbook.createSheet(Sheets.get(counter));
			  		        }
		  		        } 
		  				FileOutputStream out = new FileOutputStream(
			  			           new File(File_Path));
			  			workbook.write(out);
			  		    out.close();
		  				System.out.println(File_Path+" is written successfully");
		  				workbook.close();
	  			    } catch (Exception e) {
	  			    	System.out.println("\nE : Excel_Methods(AddExcelSheet)");
		  				e.printStackTrace();
	  			    }
	    	  }else {
			          System.out.println(File_Path+ " is not Found.");
			  }
		}
	   
	   
	   public static void RemoveExcelSheet(String File_Path, ArrayList<String> Sheets) {
		   CopyExcel(File_Path,File_Path+"_"+datetimenow());
		   File file= new File(File_Path);
	       XSSFWorkbook workbook;
		   try {
				workbook = new XSSFWorkbook(new FileInputStream(file));
				if (file.isFile() && file.exists()) { 
			    	  for (int counter = 0 ; counter < Sheets.size() ; counter++){
		    			  workbook.removeSheetAt(workbook.getSheetIndex(Sheets.get(counter)));
			    	  }
			          FileOutputStream out = new FileOutputStream(
			           new File(File_Path));
			          workbook.write(out);
		              out.close();
		              workbook.close();
		              System.out.println(File_Path+" is written successfully");
			       }else {
			          System.out.println(File_Path+ " is not Found.");
			       }
			} catch (Exception e) {
				System.out.println("\nE : Excel_Methods(RemoveExcelSheet)");
				e.printStackTrace();
			}
	   }
	   
	   public static void Addrecord(String File_Path,String Sheet_Name ,Map < Integer, Object[] > records) {
		      CopyExcel(File_Path,File_Path+"_"+datetimenow());
		      File file= new File(File_Path);
	    	  XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook(new FileInputStream(file));
				 if (file.isFile() && file.exists()) { 
			    	  XSSFSheet Sheet = workbook.getSheet(Sheet_Name);
		    	      XSSFRow row;
		    	  Set < Integer > keyid = records.keySet();
		    	  int rowid = Sheet.getLastRowNum()+1;
	              for (Integer key : keyid) {
		    	       row = Sheet.createRow(rowid++);
		    	       Object [] objectArr = records.get(key);
		    	       int cellid = 0;
	         
		    	       for (Object obj : objectArr) {
		    	    	   Cell cell = row.createCell(cellid++);
		    	    	   cell.setCellValue((String)obj);
		    	       }
		    	  }
		    	      FileOutputStream out = new FileOutputStream(
		    		           new File(File_Path));
		    		  workbook.write(out);
		    	      out.close();
		              System.out.println(File_Path+" is written successfully");
			       }else {
			          System.out.println(File_Path+ " is not Found.");
			       }
			} catch (Exception e) {
				System.out.println("\nE : Excel_Methods(Addrecord)");
				e.printStackTrace();
			}
		     
		}
	   
	   public static void Removerecord(String File_Path,String Sheet_Name,Integer row_num)  {
		      CopyExcel(File_Path,File_Path+"_"+datetimenow());
		      File file= new File(File_Path);
	    	  XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook(new FileInputStream(file));
				if (file.isFile() && file.exists()) { 
			    	  XSSFSheet Sheet = workbook.getSheet(Sheet_Name);
		    	      XSSFRow row;
		    	      int rowid = row_num-1;
		    	      row = Sheet.createRow(rowid);	 
		    	      Sheet.removeRow(row);
		    	      FileOutputStream out = new FileOutputStream(
		    		           new File(File_Path));
		    		  workbook.write(out);
		    	      out.close();
		              System.out.println(File_Path+" is written successfully");
			     }else {
			          System.out.println(File_Path+ " is not Found.");
			     }
			} catch (Exception e) {
				System.out.println("\nE : Excel_Methods(Removerecord)");
				e.printStackTrace();
			}
		      
		}
	   
	   public static int rows_num(String File_Path,String Sheet_Name) {
		      CopyExcel(File_Path,File_Path+"_"+datetimenow());
		      File file= new File(File_Path);
	    	  XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook(new FileInputStream(file));
				if (file.isFile() && file.exists()) { 
			    	  XSSFSheet Sheet = workbook.getSheet(Sheet_Name);
		    	     return Sheet.getLastRowNum()+1;
			    }else {
			          System.out.println(File_Path+ " is not Found.");
			    }
			} catch (Exception e) {
				System.out.println("\nE : Excel_Methods(rows_num)");
				e.printStackTrace();
			}
		    return 0;
		}
	   
	   public static int columns_num(String File_Path,String Sheet_Name) {
		      CopyExcel(File_Path,File_Path+"_"+datetimenow());
		      File file= new File(File_Path);
	    	  XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook(new FileInputStream(file));
				if (file.isFile() && file.exists()) { 
			    	  XSSFSheet Sheet = workbook.getSheet(Sheet_Name);
		    	      XSSFRow row = Sheet.getRow(0);
		    	       return row.getLastCellNum();
			     }else {
			          System.out.println(File_Path+ " is not Found.");
			     }
			} catch (Exception e) {
				System.out.println("\nE : Excel_Methods(columns_num)");
				e.printStackTrace();
			}
		    return 0;
		}
	   public static void editcell(String File_Path,String Sheet_Name ,Integer row_num,Integer column_num,Object cell_content) {
		      CopyExcel(File_Path,File_Path+"_"+datetimenow());
		      File file= new File(File_Path);
	    	  XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook(new FileInputStream(file));
				 if (file.isFile() && file.exists()) { 
			    	  XSSFSheet Sheet = workbook.getSheet(Sheet_Name);
		    	      XSSFRow row;
		    	      int rowid = row_num-1;
		    	      row = Sheet.getRow(rowid);
		    	      int cellid = column_num-1; 
		  	    	  Cell cell = row.createCell(cellid);
		  	    	  if(cell_content instanceof String || cell_content instanceof Integer)
		  	    		if (cell_content instanceof Integer)
		  	    			cell.setCellValue(Integer.valueOf(String.valueOf(cell_content)));
		  	    		else
		  	    			cell.setCellValue(String.valueOf(cell_content));
		  	    	  else if(cell_content instanceof Date)
			  	        cell.setCellValue((Date)cell_content); 
		  	    	  else if (cell_content instanceof Number)
		  	    		cell.setCellValue(Double.valueOf(String.valueOf(cell_content))); 
		    	      FileOutputStream out = new FileOutputStream(
		    		           new File(File_Path));
		    		  workbook.write(out);
		    	      out.close();
		              System.out.println(File_Path+" is written successfully");
			       }else {
			          System.out.println(File_Path+ " is not Found.");
			       }
			} catch (Exception e) {
				System.out.println("\nE : Excel_Methods(editcell)");
				e.printStackTrace();
			}
		     
		}
	   	   
	   public static String datetimenow() {
		   DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");  
		   return dtf.format(LocalDateTime.now());  
	   }
}
