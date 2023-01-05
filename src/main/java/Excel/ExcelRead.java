package Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	

	public static void main(String[] args) throws IOException {
		
		//Assigning excel path value to excelfilepath
		String excelfilePath="C:\\Users\\Sajida\\Desktop\\javaPrgrms\\Book1.xlsx";

		
		//FileInputStream class is  to read file
		FileInputStream inputstream=new FileInputStream(excelfilePath);
		
		//class used to represent xlfile...
        XSSFWorkbook workbook=new XSSFWorkbook(inputstream);

        //  sheet variable will get the sheetname from workboook as input in getsheet method.
        XSSFSheet sheet=workbook.getSheet("Sheet1");
        
        //rows will tell us the last rowno in the sheet
        int rows=sheet.getLastRowNum();
        System.out.println(rows);
        
       //cols will get us the last cell in the row index given.
        int cols=sheet.getRow(1).getLastCellNum();
        System.out.println(cols);
        
        for(int i=0;i<=rows;i++)
        {
           XSSFRow row= sheet.getRow(i);//gets the first row in the sheet
           System.out.println();

            for(int j=0;j<cols;j++)
            {
            XSSFCell cell=  row.getCell(j);//gets the particular cell
            
            System.out.println(cell.getCellType()+"s");
            System.out.println(cell.CELL_TYPE_STRING+"j");
            
            
            switch(cell.getCellType())
            {
            case 1:
            	System.out.print(cell.getStringCellValue()+ " ");break;
            	
            case 0:
            	System.out.println(cell.getNumericCellValue());break;
            	
            default:
                System.out.println("cell");
            }
            
            
            
            
         /*   switch (cell.getCellType())
            {
           // case STRING:System.out.println(cell.getStringCellValue());break;
            
            case CELL_TYPE_STRING:
            	System.out.println(cell.getStringCellValue());break;
            	
            case CELL_TYPE_NUMERIC:
	               // case NUMERIC:
	                    System.out.println(cell.getNumericCellValue());break;
	               // case BOOLEAN:System.out.println(cell.getBooleanCellValue());break;
	                default:
	                    System.out.println("cell");

            }*/
            }

      
            }
        }
        



	

	

}

