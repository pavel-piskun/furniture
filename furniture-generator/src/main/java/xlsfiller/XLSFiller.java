/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package xlsfiller;

import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author ober
 */
public class XLSFiller {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        String fileName = "data/база шкафчиков.xlsx";
        String sheetName = "total";
        //ExcelReader excelreader = new ExcelReader();
        List<Map<String, String>> dataList = null;
        XSSFWorkbook      workBook;

        try {
                FileInputStream fis = new FileInputStream(fileName);
                workBook = new XSSFWorkbook(fis);
                int numberOfSheets = workBook.getNumberOfSheets();
                for(int i=0; i<numberOfSheets; i++){
                    System.out.println(workBook.getSheetName(i));
                }
                
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}
