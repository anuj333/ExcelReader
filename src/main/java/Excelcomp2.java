import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Excelcomp2 {

    public static void main(String[] args) {
        try {

            FileInputStream File1 = new FileInputStream(new File( "C:\\Users\\anuj.k.singh\\Documents\\Book1.xlsx"));
            FileInputStream File2 = new FileInputStream(new File("C:\\\\Users\\\\anuj.k.singh\\\\Documents\\\\Book2.xlsx"));

            XSSFWorkbook workbook1 = new XSSFWorkbook(File1);
            XSSFWorkbook workbook2 = new XSSFWorkbook(File2);

            XSSFSheet sheet1 = workbook1.getSheetAt(0);
            XSSFSheet sheet2 = workbook2.getSheetAt(0);


            if(compareTwoSheets(sheet1, sheet2)) {
                System.out.println("\n\nThe two excel sheets are Equal");
            } else {
                System.out.println("\n\nThe two excel sheets are Not Equal");
            }


            File1.close();
            File2.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    // Compare Two Sheets
    public static boolean compareTwoSheets(XSSFSheet sheet1, XSSFSheet sheet2) {
        int firstRow1 = sheet1.getFirstRowNum();
        int lastRow1 = sheet1.getLastRowNum();
        boolean equalSheets = true;
        for(int i=firstRow1; i <= lastRow1; i++) {

            System.out.println("\n\nComparing Row "+i);

            XSSFRow row1 = sheet1.getRow(i);
            XSSFRow row2 = sheet2.getRow(i);
            if(!compareTwoRows(row1, row2)) {
                equalSheets = false;
                System.out.println("Row "+i+" - Not Equal");
                break;
            } else {
                System.out.println("Row "+i+" - Equal");
            }
        }
        return equalSheets;
    }

    public static boolean compareTwoRows(XSSFRow row1, XSSFRow row2) {
        /*
         * if((row1 == null) && (row2 == null)) { return true; } else if((row1 == null)
         * || (row2 == null)) { return false; }
         */
        int firstCell1 = row1.getFirstCellNum();
        int lastCell1 = row1.getLastCellNum();
        boolean equalRows = true;

        for(int i=firstCell1; i < lastCell1; i++) {
            XSSFCell cell1 = row1.getCell(i);
            XSSFCell cell2 = row2.getCell(i);
            if(!compareTwoCells(cell1, cell2)) {
                equalRows = false;
                System.err.println("       Cell "+i+" - NOt Equal");
                break;
            } else {
                System.out.println("       Cell "+i+" - Equal");
            }
        }
        return equalRows;
    }

    public static boolean compareTwoCells(XSSFCell cell1, XSSFCell cell2) {
	        /*if((cell1 == null) && (cell2 == null)) {
	            return true;
	        } else if((cell1 == null) || (cell2 == null)) {
	            return false;
	        }*/

        boolean equalCells = false;
        //changed from int to CellType
        CellType type1 = cell1.getCellType();
        CellType type2 = cell2.getCellType();
        if (type1 == type2) {
            if (cell1.getCellStyle().equals(cell2.getCellStyle())) {

                if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                    equalCells = true;

                }
                else if(cell1.getNumericCellValue()==cell2.getNumericCellValue()) {
                    equalCells = true;
                }

            }

            else {
                return false;
            }
        } else {
            return false;
        }
        return equalCells;

        //}

    }
}
