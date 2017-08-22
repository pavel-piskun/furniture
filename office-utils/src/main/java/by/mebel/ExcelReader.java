package by.mebel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import by.mebel.utils.ExcelUtils;

import java.util.Arrays;

public class ExcelReader {
	
	public static void main(String[] args){            
            String fileName = "d:/CloudStorages/Dropbox/haltura/автозаполнение XLS/2016.11.1/шаблон XLS Filler.xlsm";
            List<String> resultList = Arrays.asList("фурнитура мод.");
            ExcelReader excelreader = new ExcelReader();
        try {
            //List<String> resultList = excelreader.getComponentsNameList(fileName);
           // Map<String,List<Map<String,String>>> resultData = excelreader.loadData(resultList, fileName, 1, 30,0);            
            Map<String,List<Map<String,String>>> resultData = excelreader.loadData(resultList, fileName, 3, 200, 2);
            Iterator iterator = resultData.entrySet().iterator();
            while(iterator.hasNext()){
                Entry<String, List<Map<String,String>>> entry = (Entry)iterator.next();
                System.out.println("================\n"+entry.getKey()+"\n===============");
                List<Map<String,String>> data = entry.getValue();
                Map<String,String> dataMap;
                for (int i = 0; i < data.size(); i++) {
                    System.out.println();
                    dataMap = data.get(i);
                    Iterator it = dataMap.entrySet().iterator();
                    while(it.hasNext()){                        
                        Entry<String, String> entry1 = (Entry)it.next();
                        System.out.print(entry1.toString()+" | ");
                    }
                }
             //////////////////////////////////////
               /*
                System.out.println("=========+++++++++++++++++++===========");
                List<String> resultFList = excelreader.getComponentsNameList(fileName);
                Map<String,List<Map<String,String>>> resultFData = excelreader.loadData(resultFList, fileName, 41, 38,40);
                Iterator iteratorF = resultFData.entrySet().iterator();
                while(iteratorF.hasNext()){
                    Entry<String, List<Map<String,String>>> entryF = (Entry)iteratorF.next();
                    System.out.println("================\n"+entryF.getKey()+"\n===============");
                    List<Map<String,String>> dataF = entryF.getValue();
                    Map<String,String> dataFMap;
                    for (int i = 0; i < dataF.size(); i++) {
                        System.out.println();
                        dataFMap = dataF.get(i);
                        Iterator it = dataFMap.entrySet().iterator();
                        while(it.hasNext()){                        
                            Entry<String, String> entry1 = (Entry)it.next();
                            System.out.print(entry1.toString()+" | ");
                        }
                    }
                }
            */
            
//            String fileName = "н- Масаров 2.xlsx";
//            String sheetName = "total";
//            ExcelReader excelreader = new ExcelReader();
//            List<Map<String, String>> dataList = null;
//            try {
//                    FileInputStream fis = new FileInputStream(fileName);
//                    dataList = excelreader.read(fis,sheetName);
//                    System.out.println("Count of rows: "+dataList.size());
//                    dataList = ExcelUtils.removeEmptyObjects(dataList);
//                    System.out.println("Count of rows after filtering: "+dataList.size());
//                    for(int i=0;i<dataList.size();i++){
//                            Map<String,String>  rowDataMap = dataList.get(i);
//                            Iterator iterator = rowDataMap.entrySet().iterator();
//                            while(iterator.hasNext()){
//                                    Entry<String, String> entry = (Entry)iterator.next();
//                                    System.out.println(entry.toString());
//                            }
//                            System.out.println("---------------------");
//                    }
//                    WordWriter wordWriter = new WordWriter();
//                    wordWriter.createWordDocument("template.docx","testOut.docx", dataList);
//            } catch (Exception e) {
//                    // TODO Auto-generated catch block
//                    e.printStackTrace();
//            }
            }        } catch (Exception ex) {
            ex.printStackTrace();
        }
	}
	
	public List<Map<String, String>> read(FileInputStream fis, String sheetName){
		List<Map<String, String>> result = new ArrayList<Map<String, String>>();
		boolean isHeader = true;
		List<String> headersList = null;
		XSSFWorkbook      workBook;
		try
		{	
			workBook = new XSSFWorkbook(fis);
			FormulaEvaluator  evaluator = workBook.getCreationHelper().createFormulaEvaluator();
			XSSFSheet         sheet    = workBook.getSheet(sheetName);
			Iterator<Row> rows     = sheet.rowIterator();
			while(rows.hasNext()){
				Row row = rows.next ();
				Iterator<Cell> cells = row.cellIterator();
				
				if(isHeader){
					headersList = readHeaders(cells);
					isHeader = false;
				}else{					
					result.add(readData(cells, headersList, evaluator));
				}
			}
		}catch(OutOfMemoryError ex){
			throw ex;
		}catch(Exception e){
			e.printStackTrace();
		}finally{
			try {				
				fis.close();
			} catch (IOException e) {				
				e.printStackTrace();
			}
		}
		return result;
	}
	
	private List<String> readHeaders(Iterator<Cell> cells){
		List<String> result = new ArrayList<String>();
		while(cells.hasNext()){
			Cell cell = cells.next();
			 
				switch (cell.getCellType ())
			{
				case Cell.CELL_TYPE_NUMERIC :
				{
					result.add(String.format("%d", (int) cell.getNumericCellValue()));
					break;
				}
				case Cell.CELL_TYPE_STRING:
				{
					result.add(cell.getRichStringCellValue().getString());
					break;
				}				
				default:
				{
					result.add("notRecognized");
					break;
				}
			}
		}
		return result;
	}
	private Map<String, String> readData(Iterator<Cell> cells, List<String> headers, FormulaEvaluator evaluator){
		Map<String, String> result = new HashMap<String, String>();
		int i = 0;                
		while(cells.hasNext()){
			Cell cell = cells.next();	
			CellValue cellValue = evaluator.evaluate(cell);
                        if(cellValue != null){
                            switch (cellValue.getCellType ())
                            {
                                    case Cell.CELL_TYPE_NUMERIC :
                                    {
                                            result.put(headers.get(i), String.format("%d", (int) cell.getNumericCellValue()));					
                                            break;
                                    }
                                    case Cell.CELL_TYPE_STRING:
                                    {
                                            result.put(headers.get(i), cellValue.getStringValue());					
                                            break;
                                    }
                                    case Cell.CELL_TYPE_FORMULA:
                                    {
                                            result.put(headers.get(i), "formula");
                                            break;
                                    }
                                    default:
                                    {
                                            result.put(headers.get(i), "notRecognized");
                                            break;
                                    }
                            }
                        }else{
                            result.put(headers.get(i), "");
                        }
			i++;
		}
		return result;
	}
	public List<String> getComponentsNameList(String fileName) throws Exception{
            List<String> resultList = new ArrayList<String>();            
            FileInputStream fis = null;
            XSSFWorkbook workBook = null;			
            try {
                fis = new FileInputStream(fileName);                
                workBook = new XSSFWorkbook(fis);
                int numberOfSheets = workBook.getNumberOfSheets();
                /*
                 * Не добавляем название первой страницы с исходными данными
                 */
                for(int i=1; i<numberOfSheets; i++){
                    resultList.add(workBook.getSheetName(i));
                }           
            } catch (Exception ex) {
                if(workBook != null){
                    //workBook.cloneSheet(sheetNum)
                }
                if(fis != null){
                    try {
                        fis.close();
                    } catch (IOException ex1) {
                        System.out.println("ТРЫНДЕЦ!: "+ex1.getMessage());
                    }
                }
                throw ex;
            }finally{
                try {
                    if(fis!=null){
                        fis.close();
                    }
                } catch (IOException ex) {
                     System.out.println("ТРЫНДЕЦ!: "+ex.getMessage());
                }
            }             
            return resultList;
        }
        public Map<String,List<Map<String,String>>> loadDataMultipleeTablesPerRow(List<String> componentsNameList, String fileName, int startIdx, int rowsCount, int headerIdx, int columnIdx, int columnCount) throws Exception{
            FileInputStream fis = null;
            XSSFWorkbook workBook = null;			
            List<String> componentDataHeader;
            List<Map<String,String>> componentData;
            Map<String,List<Map<String,String>>> resultMap = new HashMap<String, List<Map<String, String>>>();
            try {
                fis = new FileInputStream(fileName);                
                workBook = new XSSFWorkbook(fis);
                FormulaEvaluator  evaluator = workBook.getCreationHelper().createFormulaEvaluator();
                String componentName = null;
                for (int i = 0; i < componentsNameList.size(); i++) {
                    componentName = componentsNameList.get(i);
                    XSSFSheet sheet = workBook.getSheet(componentName);
                    componentDataHeader = getDataHeader(sheet, headerIdx, columnIdx, columnCount);
                    componentData = getData(sheet, startIdx, rowsCount, componentDataHeader, evaluator, columnIdx, columnCount);
                    resultMap.put(componentsNameList.get(i), componentData);
                }
            }catch (Exception ex) {
                if(workBook != null){
                    //workBook.cloneSheet(sheetNum)
                }
                if(fis != null){
                    try {
                        fis.close();
                    } catch (IOException ex1) {
                        System.out.println("ТРЫНДЕЦ!: "+ex1.getMessage());
                    }
                }
                throw ex;
            }finally{
                try {
                    if(fis!=null){
                        fis.close();
                    }
                } catch (IOException ex) {
                     System.out.println("ТРЫНДЕЦ!: "+ex.getMessage());
                }
            }
            return resultMap;        
        }
        
        public Map<String,List<Map<String,String>>> loadData(List<String> componentsNameList, String fileName, int startIdx, int rowsCount, int headerIdx) throws Exception{
            FileInputStream fis = null;
            XSSFWorkbook workBook = null;			
            List<String> componentDataHeader;
            List<Map<String,String>> componentData;
            Map<String,List<Map<String,String>>> resultMap = new HashMap<String, List<Map<String, String>>>();
            try {
                fis = new FileInputStream(fileName);                
                workBook = new XSSFWorkbook(fis);
                FormulaEvaluator  evaluator = workBook.getCreationHelper().createFormulaEvaluator();
                String componentName = null;
                for (int i = 0; i < componentsNameList.size(); i++) {
                    componentName = componentsNameList.get(i);
                    XSSFSheet sheet = workBook.getSheet(componentName);                    
                    componentDataHeader = getDataHeader(sheet, headerIdx);
                    componentData = getData(sheet, startIdx, rowsCount, componentDataHeader, evaluator);
                    resultMap.put(componentsNameList.get(i), componentData);
                }
            }catch (Exception ex) {
                if(workBook != null){
                    //workBook.cloneSheet(sheetNum)
                }
                if(fis != null){
                    try {
                        fis.close();
                    } catch (IOException ex1) {
                        System.out.println("ТРЫНДЕЦ!: "+ex1.getMessage());
                    }
                }
                throw ex;
            }finally{
                try {
                    if(fis!=null){
                        fis.close();
                    }
                } catch (IOException ex) {
                     System.out.println("ТРЫНДЕЦ!: "+ex.getMessage());
                }
            }
            return resultMap;
        }
        /*
        Multiply tables per row method
        */
        private List<String> getDataHeader(XSSFSheet sheet, int headerIdx, int columnIdx, int columnCount) {
            List<String> result = new ArrayList<>();
            Row row = sheet.getRow(headerIdx);
            for(int i=0; i<columnCount; i++) {
                Cell cell = row.getCell(columnIdx+i);
                switch (cell.getCellType ())
                    {
                            case Cell.CELL_TYPE_NUMERIC :
                            {
                                    result.add(String.format("%d", (int) cell.getNumericCellValue()));
                                    break;
                            }
                            case Cell.CELL_TYPE_STRING:
                            {
                                    result.add(cell.getRichStringCellValue().getString());
                                    break;
                            }				
                            default:
                            {
                                    result.add("notRecognized");
                                    break;
                            }
                    }
            }                  
            return result;
        }   
        /*
        Single table per row
        */
        private List<String> getDataHeader(XSSFSheet sheet, int headerIdx){
            Row row = sheet.getRow(headerIdx);
            Iterator<Cell> cells = row.cellIterator();
            List<String> result = new ArrayList<String>();
            while(cells.hasNext()){
                    Cell cell = cells.next();

                            switch (cell.getCellType ())
                    {
                            case Cell.CELL_TYPE_NUMERIC :
                            {
                                    result.add(String.format("%d", (int) cell.getNumericCellValue()));
                                    break;
                            }
                            case Cell.CELL_TYPE_STRING:
                            {
                                    result.add(cell.getRichStringCellValue().getString());
                                    break;
                            }				
                            default:
                            {
                                    result.add("notRecognized");
                                    break;
                            }
                    }
            }
            return result;
        }
        /*
        Multiply tables per row method
        */
        private List<Map<String, String>> getData(XSSFSheet sheet, int startIdx, int rowsCount, List<String> headers, FormulaEvaluator evaluator, int columnIdx, int columnCount) {
            List<Map<String,String>> result = new ArrayList<Map<String, String>>();
            Row row;
            Iterator<Cell> cells;
            Map<String, String> dataMap;
            for (int i = startIdx; i < startIdx + rowsCount; i++) {
                row = sheet.getRow(i);
                if(row!= null){                                      
                    dataMap = new HashMap<String, String>();
                    int j = 0;                
                    for(int k = 0; k < columnCount; k ++) {
                            Cell cell = row.getCell(columnIdx + k);
                            CellValue cellValue = evaluator.evaluate(cell);
                            if(cellValue != null){
                                switch (cellValue.getCellType ())
                                {
                                        case Cell.CELL_TYPE_NUMERIC :
                                        {
                                                
                                                dataMap.put(headers.get(j), Double.toString(cell.getNumericCellValue()));
                                                //dataMap.put(headers.get(j), String.format("%d", (int)cell.getNumericCellValue()));
                                                
                                                break;
                                        }
                                        case Cell.CELL_TYPE_STRING:
                                        {
                                                dataMap.put(headers.get(j), cellValue.getStringValue());					
                                                break;
                                        }
                                        case Cell.CELL_TYPE_FORMULA:
                                        {
                                                dataMap.put(headers.get(j), "formula");
                                                break;
                                        }
                                        default:
                                        {
                                                dataMap.put(headers.get(j), "notRecognized");
                                                break;
                                        }
                                }
                            }else{
                                
                                dataMap.put(headers.get(j), "");
                            }
                            j++;                        
                    }
                    result.add(dataMap);
                }
            }
            return ExcelUtils.removeEmptyObjects(result);
        }
        
        /*
        Single table per row
        */
        private List<Map<String,String>> getData(XSSFSheet sheet, int startIdx, int rowsCount, List<String> headers,FormulaEvaluator evaluator){
            List<Map<String,String>> result = new ArrayList<Map<String, String>>();
            Row row;
            Iterator<Cell> cells;
            Map<String, String> dataMap;
            for (int i = startIdx; i < startIdx + rowsCount; i++) {
                row = sheet.getRow(i);
                if(row!= null){
                    cells = row.cellIterator();                    
                    dataMap = new HashMap<String, String>();
                    int j = 0;                
                    while(cells.hasNext() && j < headers.size()){
                            Cell cell = cells.next();	
                            CellValue cellValue = evaluator.evaluate(cell);
                            if(cellValue != null){
                                switch (cellValue.getCellType ())
                                {
                                        case Cell.CELL_TYPE_NUMERIC :
                                        {
                                                
                                                dataMap.put(headers.get(j), Double.toString(cell.getNumericCellValue()));
                                                //dataMap.put(headers.get(j), String.format("%d", (int)cell.getNumericCellValue()));
                                                
                                                break;
                                        }
                                        case Cell.CELL_TYPE_STRING:
                                        {
                                                dataMap.put(headers.get(j), cellValue.getStringValue());					
                                                break;
                                        }
                                        case Cell.CELL_TYPE_FORMULA:
                                        {
                                                dataMap.put(headers.get(j), "formula");
                                                break;
                                        }
                                        default:
                                        {
                                                dataMap.put(headers.get(j), "notRecognized");
                                                break;
                                        }
                                }
                            }else{
                                
                                dataMap.put(headers.get(j), "");
                            }
                            j++;                        
                    }
                    result.add(dataMap);
                }
            }
            return ExcelUtils.removeEmptyObjects(result);
        }    
}
