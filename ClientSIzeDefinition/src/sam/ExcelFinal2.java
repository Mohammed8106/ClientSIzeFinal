package sam;


import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;




public class ExcelFinal2{

    public static void main(String[] args) 
    {
        try 
            {
        	
        	 Scanner scanner = new Scanner(System.in);
             System.out.print("Enter the path to the excel file: "+"\n");
             
             
             String filePath = scanner.nextLine();
             System.out.println("Enter the sheet name: "+"\n"); 
             
             
             String sheetName= scanner.nextLine();
             scanner.close();
             
             
            FileInputStream file = new FileInputStream(filePath);
            Workbook workbook = WorkbookFactory.create(file);

            
            Sheet sheet = workbook.getSheet(sheetName);

            int rowCount = sheet.getLastRowNum();
            for (int i = 1; i <= rowCount; i++) 
            {
                Row currentRow = sheet.getRow(i);

                if (currentRow == null) 
                {
                    continue;                                       // Skip processing empty row
                }
                

                Cell aumCell = currentRow.getCell(0);
                Cell employeeCountCell = currentRow.getCell(1);
                Cell newColumnCell = currentRow.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                double aumValue;
                double employeeCountValue;

                if (aumCell.getCellType() == CellType.NUMERIC) 
                {
                    aumValue = aumCell.getNumericCellValue();
                }
                
                
                else if (aumCell.getCellType() == CellType.STRING)
                {
                    continue;                                         // Skip processing if the cell contains a non-numeric value
                }
                else
                	
                {
                    String aumString = aumCell.getStringCellValue();
                    aumValue = aumString.isEmpty() ? 0.0 : Double.parseDouble(aumString);
                }

                
                if (employeeCountCell.getCellType() == CellType.NUMERIC) 
                {
                    employeeCountValue = employeeCountCell.getNumericCellValue();
                }
                            
                else if (employeeCountCell.getCellType() == CellType.STRING)
                {
                    continue;                                         // Skip processing if the cell contains a non-numeric value
                }
                
                else 
                
                {
                    String employeeCountString = employeeCountCell.getStringCellValue();
                    employeeCountValue = employeeCountString.isEmpty() ? 0.0 : Double.parseDouble(employeeCountString);
                }

                if ( employeeCountValue < 10 &&  employeeCountValue >0 && aumValue < 10000L &&  aumValue >0) 
                {
                    newColumnCell.setCellValue("Small");
                }
                
                else if (employeeCountValue >= 10 && employeeCountValue  <= 20 &&  aumValue >= 10000L && aumValue <= 20000L)
                {
                    newColumnCell.setCellValue("Medium");
                }
                
                else if(employeeCountValue >=10 && employeeCountValue <=20 && aumValue >=20 && aumValue<=100) {
                	newColumnCell.setCellValue("Medium");
                }
                
                
                else if  (employeeCountValue >= 20 && employeeCountValue <= 60 && aumValue >= 20000L && aumValue <= 100000L)
                {
                    newColumnCell.setCellValue("Large");
                }
                
                else if(employeeCountValue >60 && aumValue >= 20000L && aumValue <= 100000L ) 
                {
                	newColumnCell.setCellValue("Large");
                }
                
                else if (employeeCountValue > 60 && aumValue > 100000L) 
                {
                    newColumnCell.setCellValue("Mega");
                }
                             
                
            }
            

            file.close();

            FileOutputStream outFile = new FileOutputStream(filePath);
            workbook.write(outFile);
            outFile.close();

            System.out.println("Performing operations on excel file......"+"\n");
            System.out.println("Applying client size definition.........."+"\n");
            System.out.println(" Data is added in a new column successfully");
        } 
        catch (IOException e)
        {
            e.printStackTrace();
        }
    }
}