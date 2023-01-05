package Excel;


	import java.io.FileNotFoundException;
	import java.io.FileOutputStream;
	import java.io.IOException;

	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	//Workbook--sheet--rows==cells

	public class ExcelWrite {
		


		public static void main(String[] args) throws IOException {

			XSSFWorkbook workbook=new XSSFWorkbook();
			XSSFSheet sheet=workbook.createSheet("Emp Info");
			
			Object empdata[][] = {{"Emp id", "Name","Job"},
									{101,"David","Engineer"},
									{102,"Tom","Manager"},
									{103,"Scott","Anayst"}
									
								};		
			
			int rows=empdata.length;
			int cols=empdata[0].length;
			System.out.println(rows);
			System.out.println(cols);
			
			for(int i=0;i<rows;i++)
				{
				XSSFRow row=sheet.createRow(i);
				
				for(int j=0;j<cols;j++)
				{
					XSSFCell cell=row.createCell(j);
					Object value=empdata[i][j];
					
					if(value instanceof String)
						cell.setCellValue((String)value);
					if(value instanceof Integer)
						cell.setCellValue((Integer)value);
					if(value instanceof Boolean)
						cell.setCellValue((Boolean)value);
					
				}
				
				}
			String filePath="C:\\Users\\Sajida\\Desktop\\javaPrgrms\\Book2.xlsx";
			FileOutputStream outstream=new FileOutputStream(filePath);
			workbook.write(outstream);
			outstream.close();
			

		}

	}




