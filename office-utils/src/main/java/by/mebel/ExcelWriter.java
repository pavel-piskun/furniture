package by.mebel;

import by.mebel.utils.ExcelUtils;
import by.mebel.utils.StringUtils;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author pavel.piskun
 */
public class ExcelWriter {
    public static final String DETAILS_TEMPLATE_SHEET_NAME = "детализация";
    public static final String FURNITURE_TEMPLATE_SHEET_NAME = "фурнитура мод.";
    public static final String MATERIALS_TEMPLATE_SHEET_NAME = "исх.данные";
    public static final String ADDITIONAL_OPERATIONS_TEMPLATE_SHEET_NAME = "операции для мод.";
    public static final Boolean DEFAULT_GENERATE_ERROR_ON_MISSING_ELEMENT_VALUE = Boolean.TRUE;
    
    public static final String FURNITURE_TITLE_COLUMN_NAME = "наименование";
    public static final String FURNITURE_ARTICLE_COLUMN_NAME = "art";
    public static final String FURNITURE_COLOR_COLUMN_NAME = "цвет";
    public static final String FURNITURE_SIZE_COLUMN_NAME = "разм.(мм)";
    public static final String FURNITURE_DIMENSION_COLUMN_NAME = "ед.изм.";
    public static final String FURNITURE_COUNT_COLUMN_NAME = "кол-во";    
    
    public static final int TEMPLATE_FURNITURE_TITLE_COLUMN_INDEX = 1;
    public static final int TEMPLATE_FURNITURE_ARTICLE_COLUMN_INDEX = 2;
    public static final int TEMPLATE_FURNITURE_COLOR_COLUMN_INDEX = 3;
    public static final int TEMPLATE_FURNITURE_SIZE_COLUMN_INDEX = 4;
    public static final int TEMPLATE_FURNITURE_DIMENSION_COLUMN_INDEX = 5;
    public static final int TEMPLATE_FURNITURE_COUNT_COLUMN_INDEX = 6; 
    
    
    public static final String ADDITIONAL_OPERATIONS_TITLE_COLUMN_NAME = "операция";
    public static final String ADDITIONAL_OPERATIONS_DIMENSION_COLUMN_NAME = "ед.изм.";
    public static final String ADDITIONAL_OPERATIONS_COUNT_COLUMN_NAME = "кол-во";
    
    public static final int TEMPLATE_ADDITIONAL_OPERATIONS_TITLE_COLUMN_INDEX = 1;
    public static final int TEMPLATE_ADDITIONAL_OPERATIONS_DIMENSION_COLUMN_INDEX = 2;
    public static final int TEMPLATE_ADDITIONAL_OPERATIONS_COUNT_COLUMN_INDEX = 3;
    
    
    
    /*
     * Мапа для хранения соответствия "название поля в шаблоне" -> "номер столбца в шаблоне"
     * для каждого вхождения в маппинг
     *  key - название поля в шаблоне
     *  value - номер столбца
    */
    public static final Map<String, Integer> TEMPLATE_FIELDS_MAPPING;
    static
    {
        TEMPLATE_FIELDS_MAPPING = new HashMap<String, Integer>();
        TEMPLATE_FIELDS_MAPPING.put("наименование дет.", new Integer(1));
        TEMPLATE_FIELDS_MAPPING.put("№ осн.", 3);
        TEMPLATE_FIELDS_MAPPING.put("текст.", 7); 
        TEMPLATE_FIELDS_MAPPING.put("1 дл.", 9); 
        TEMPLATE_FIELDS_MAPPING.put("2 дл.", 10); 
        TEMPLATE_FIELDS_MAPPING.put("длина (L)mm", new Integer(8));
        TEMPLATE_FIELDS_MAPPING.put("ширина(W)mm", new Integer(11));
        TEMPLATE_FIELDS_MAPPING.put("1 ш.", new Integer(12));
        TEMPLATE_FIELDS_MAPPING.put("2 ш.", new Integer(13));
        TEMPLATE_FIELDS_MAPPING.put("кол-во(шт)", new Integer(16));
        TEMPLATE_FIELDS_MAPPING.put("паз.", new Integer(14));
        TEMPLATE_FIELDS_MAPPING.put("четв.", new Integer(15));
        TEMPLATE_FIELDS_MAPPING.put("примечание", new Integer(20));
        TEMPLATE_FIELDS_MAPPING.put("обр. уч.№1", new Integer(17));
        TEMPLATE_FIELDS_MAPPING.put("обр. уч.№2", new Integer(18));
        TEMPLATE_FIELDS_MAPPING.put("обр. уч.№3", new Integer(19));        
    }
    
    /*
     * Мапа для хранения соответствия "название поля в шаблоне" -> "название поля в файле с БД компонентов"
     * для каждого вхождения в маппинг
     *  key - название поля в шаблоне
     *  value - название поля в БД(XLS)
    */
    public static final Map<String, String> FIELDS_MAPPING;
    static
    {
        FIELDS_MAPPING = new HashMap<String, String>();
        FIELDS_MAPPING.put("№ осн.","№ схемы");
        FIELDS_MAPPING.put("наименование дет.", "наименование дет.");
        FIELDS_MAPPING.put("текст.", "текст.");        
        FIELDS_MAPPING.put("длина (L)mm", "длина (L)mm");
        FIELDS_MAPPING.put("1 дл.", "1 дл.");
        FIELDS_MAPPING.put("2 дл.", "2 дл.");
        FIELDS_MAPPING.put("ширина(W)mm", "ширина(W)mm");
        FIELDS_MAPPING.put("1 ш.", "1 ш.");
        FIELDS_MAPPING.put("2 ш.", "2 ш.");
        FIELDS_MAPPING.put("кол-во(шт)", "кол-во(шт)");
        FIELDS_MAPPING.put("паз.", "паз.");
        FIELDS_MAPPING.put("четв.", "четв.");
        FIELDS_MAPPING.put("примечание", "примечание");
        FIELDS_MAPPING.put("обр. уч.№1", "обр. уч.№1");
        FIELDS_MAPPING.put("обр. уч.№2", "обр. уч.№2");
        FIELDS_MAPPING.put("обр. уч.№3", "обр. уч.№3");        
    }
    /*
     * Мапа для хранения соответствия "название поля в таблице "Основные материалы" в шаблоне" -> "номер столбца в шаблоне"
     * для каждого вхождения в маппинг
     *  key - название поля в шаблоне
     *  value - номер столбца
    */
    public static final Map<String, Integer> TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING;
    static
    {
        TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING = new HashMap<String, Integer>();
        TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING.put("№ основы", new Integer(0));
        TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING.put("материал", new Integer(1));
        TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING.put("цвет", new Integer(2));
        TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING.put("толщина", new Integer(3));        
    }
    /*
     * Мапа для хранения соответствия "название поля в таблице "основные материалы" в шаблоне" -> "название поля в файле с БД компонентов"
     * для каждого вхождения в маппинг
     *  key - название поля в шаблоне
     *  value - название поля в БД(XLS)
    */
    public static final Map<String, String> MAIN_MATERIALS_FIELDS_MAPPING;
    static
    {
        MAIN_MATERIALS_FIELDS_MAPPING = new HashMap<String, String>();        
        MAIN_MATERIALS_FIELDS_MAPPING.put("№ основы", "№ схемы");        
        MAIN_MATERIALS_FIELDS_MAPPING.put("материал", "материал");
        MAIN_MATERIALS_FIELDS_MAPPING.put("цвет", "цвет");
        MAIN_MATERIALS_FIELDS_MAPPING.put("толщина", "толщина");        
    }
    /*
     * Мапа для хранения соответствия "название поля в таблице "Облицовочные материалы" в шаблоне" -> "номер столбца в шаблоне"
     * для каждого вхождения в маппинг
     *  key - название поля в шаблоне
     *  value - номер столбца
    */
    public static final Map<String, Integer> TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING;
    static
    {
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING = new HashMap<String, Integer>();
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING.put("№", new Integer(0));
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING.put("материал", new Integer(1));
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING.put("цвет", new Integer(2));
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING.put("толщина/ширина", new Integer(3));        
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING.put("обозначение", new Integer(4));
    }
    /*
     * Мапа для хранения соответствия "название поля в таблице "облицовочные материалы" в шаблоне" -> "название поля в файле с БД компонентов"
     * для каждого вхождения в маппинг
     *  key - название поля в шаблоне
     *  value - название поля в БД(XLS)
    */
    public static final Map<String, String> FACING_MATERIALS_FIELDS_MAPPING;
    static
    {
        FACING_MATERIALS_FIELDS_MAPPING = new HashMap<String, String>();
        FACING_MATERIALS_FIELDS_MAPPING.put("№", "№ схемы");
        FACING_MATERIALS_FIELDS_MAPPING.put("материал", "материал");
        FACING_MATERIALS_FIELDS_MAPPING.put("цвет", "цвет");
        FACING_MATERIALS_FIELDS_MAPPING.put("толщина/ширина", "толщина/ширина");        
        FACING_MATERIALS_FIELDS_MAPPING.put("обозначение", "обозначение");
    }
    /*public static final Map<String, String> TEMPLATE_SHEETS_NAMES_MAPPING;
    static
    {
        TEMPLATE_SHEETS_NAMES_MAPPING = new HashMap<String, String>();
        TEMPLATE_SHEETS_NAMES_MAPPING.put("наименование дет.", "наимен.детали");
        TEMPLATE_SHEETS_NA6MES_MAPPING.put("длина (L)mm", "длина");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("ширина(W)mm", "ширина");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("кол-во(шт)", "кол-во");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("паз.", "паз");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("четв.", "четв.");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("примечание", "примечание");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("обр. уч.№1", "уч.1");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("обр. уч.№2", "уч.2");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("обр. уч.№3", "уч.3");        
    }*/
    public static void main(String[] args) {
        String fileName = "d:\\CloudStorges\\Dropbox\\haltura\\автозаполнение XLS\\база шкафчиков.xlsx";
        String templateFileName = "d:\\CloudStorges\\Dropbox\\haltura\\автозаполнение XLS\\шаблон beta кухня.xlsx";
        String targetFile = "d:\\test.xlsx";
        
        /*
        FileInputStream fis;
        XSSFWorkbook workBook = null;	
        try {
            fis = new FileInputStream(templateFileName);                
            workBook = new XSSFWorkbook(fis);
            System.out.println(workBook.getCTWorkbook().getSheets().getSheetArray(3).getName());
        } catch (Exception e) {
            e.printStackTrace();
        }
        
        
        
        try {
            
            ExcelReader excelReader = new ExcelReader();
            List<String> resultList = excelReader.getComponentsNameList(fileName);
            Map<String,List<Map<String,String>>> resultData = excelReader.loadData(resultList, fileName, 40, 30,39);
            //writeDataInTemplateXLS(templateFileName, targetFile, resultData, resultData, 6, 4,100);
            Map<String,Integer> calculatedFurn = calculateFurniture(resultData);
            Iterator iterator = calculatedFurn.entrySet().iterator();
            while(iterator.hasNext()){
                Map.Entry<String, Integer> entry = (Map.Entry)iterator.next();
                System.out.println(entry.getKey()+" -> "+entry.getValue());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        */
        /*
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet firstSheet = workbook.createSheet("TEST SHEET");
        HSSFRow rowA = firstSheet.createRow(0);
        HSSFCell cellA = rowA.createCell(0);
        cellA.setCellValue(new HSSFRichTextString("TEST MESSAGE"));
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(new File("export/CreateExcelDemo.xls"));
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fos != null) {
                try {
                    fos.flush();
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        */
    }

    public static void createExcelFile(String pathToExport,String fileName, List<Map<String, String>> dataList) throws RuntimeException {

        String excelFileName;
        FileOutputStream fos = null;

        Map<String, String> dataMap;
        HSSFWorkbook workbook;
        HSSFSheet firstSheet;
        HSSFRow row;
        HSSFCell cell;

        workbook = new HSSFWorkbook();
        firstSheet = workbook.createSheet("Лист1");
        for (int i = 0; i < dataList.size(); i++) {
            dataMap = (Map<String, String>) dataList.get(i);
            row = firstSheet.createRow(i);
            //добавляем первый столбец
            cell = row.createCell(0);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.ORDER_FIELD_NAME) + "," + dataMap.get(ExcelUtils.INDEX_NUMBER_FIELD_NAME)));
            //добавляем столбец
            cell = row.createCell(1);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.LENGTH_FIELD_NAME)));
            //добавляем столбец
            cell = row.createCell(2);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.WEIGTH_FIELD_NAME)));
            //добавляем столбец
            cell = row.createCell(3);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.COUNT_FIELD_NAME)));
            //добавляем столбец
            cell = row.createCell(4);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.TEXTURE_FIELD_NAME)));
            //добавляем столбец
            cell = row.createCell(5);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.PAZ_FIELD_NAME)));
            //добавляем столбец
            cell = row.createCell(6);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.COMMENT_FIELD_NAME)));
        }
        try {
            fos = new FileOutputStream(new File(pathToExport+"/" + fileName + ".xls"));
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(e.getMessage());
        } finally {
            if (fos != null) {
                try {
                    fos.flush();
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
    public static void writeDataInTemplateXLS(String templateXLSName, String targetXLSFileName,
                                                Map<String,List<Map<String,String>>> componentsDataMap, 
                                                Map<String, Integer> componentCountMap,
                                                Map<String,List<Map<String,String>>> furnitureDataMap,
                                                Map<String,List<Map<String,String>>> mainMaterial,
                                                Map<String,List<Map<String,String>>> facingMaterial,
                                                Map<String,List<Map<String,String>>> additionalOperationsDataMap,
                                                int detailsSheetStartIdx, int furnitureSheetStartIdx,
                                                int furnitureRowCount,int furnitureHeaderIdx, 
                                                int mainMaterialStartIdx, int facingMaterialStartIdx,
                                                int additionalOpertationsHeaderIdx, int additionalOpertationsStartIdx, 
                                                int additionalOpertationsRowCount) throws Exception{
        
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workBook = null;	
        //XSSFSheet sheet = workBook.getSheet(componentName);
        /*
         * Поучаем данные из страницы "фурнитура" в шаблоне
         */
        List<String> tmpSheetNameList = new ArrayList<String>();
        tmpSheetNameList.add(FURNITURE_TEMPLATE_SHEET_NAME);
        
        
        ExcelReader excelReader = new ExcelReader();                        
        Map<String,List<Map<String,String>>> templateFurnitureData = excelReader.loadData(tmpSheetNameList, templateXLSName, furnitureSheetStartIdx, furnitureRowCount, furnitureHeaderIdx);
        //System.out.println("Template Furniture Data row count: "+ getRowCount(templateFurnitureData));
        //System.out.println(StringUtils.dataToString(templateFurnitureData));
        
        tmpSheetNameList.remove(0);
        tmpSheetNameList.add(ADDITIONAL_OPERATIONS_TEMPLATE_SHEET_NAME);
        Map<String,List<Map<String,String>>> templateAdditionalOperationsDataMap = excelReader.loadData(tmpSheetNameList, templateXLSName, additionalOpertationsStartIdx, additionalOpertationsRowCount, additionalOpertationsHeaderIdx);
        System.out.println("Template Furniture Data row count: "+ getRowCount(templateAdditionalOperationsDataMap));
        System.out.println(StringUtils.dataToString(templateAdditionalOperationsDataMap));
        
        //System.out.println("Components Data row count: "+ getRowCount(componentsDataMap));
        //System.out.println("Furniture Data row count: "+ getRowCount(furnitureDataMap));
        try {
            fis = new FileInputStream(templateXLSName);                
            workBook = new XSSFWorkbook(fis);
            /**
             * Вызов метода который записывает основные материалы в шаблон
             */            
            writeDataInSheetOfTemplate(workBook, mainMaterial, null, MAIN_MATERIALS_FIELDS_MAPPING, TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING, MATERIALS_TEMPLATE_SHEET_NAME, mainMaterialStartIdx, Boolean.TRUE);
            /*
             * Вызов метода который  записывает облицовочные материалы в шаблон
             */            
            writeDataInSheetOfTemplate(workBook, facingMaterial, null, FACING_MATERIALS_FIELDS_MAPPING, TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING, MATERIALS_TEMPLATE_SHEET_NAME, facingMaterialStartIdx, Boolean.TRUE);
            /**
             * Вызов метода  который записывает данные из БД в шаблон
             */
            writeDataInSheetOfTemplate(workBook, componentsDataMap, null/*componentCountMap*/,FIELDS_MAPPING, TEMPLATE_FIELDS_MAPPING, DETAILS_TEMPLATE_SHEET_NAME, detailsSheetStartIdx, Boolean.TRUE);
            /**
             * Вызов метода который подсчитывает и записывает фурнитуру в шаблон
             */
           calculateAndWriteFurnitureDataInSheetOfTemplate(workBook,
                                                            FURNITURE_TEMPLATE_SHEET_NAME,
                                                            templateFurnitureData,
                                                            furnitureDataMap, 
                                                            componentCountMap,
                                                            furnitureSheetStartIdx,                                                             
                                                            DEFAULT_GENERATE_ERROR_ON_MISSING_ELEMENT_VALUE);            
            /**
             * Вызов метода который  подсчитывает количество и записывает в шаблон дополнительные операции
             */
            calculateAndWriteAdditionalOperationsDataInSheetOfTemplate(workBook, 
                                                                        ADDITIONAL_OPERATIONS_TEMPLATE_SHEET_NAME, 
                                                                        templateAdditionalOperationsDataMap, 
                                                                        additionalOperationsDataMap, 
                                                                        componentCountMap, 
                                                                        additionalOpertationsStartIdx, 
                                                                        DEFAULT_GENERATE_ERROR_ON_MISSING_ELEMENT_VALUE);
           /*
             * TODO: Дописать расчет общей площади объектов по метериалу, входящих в компонент
             */
            /*
             * Вычисляем все формулы в документе
             * в версии 3.10b2 не реализовано вычисление формул ссылающихся на другие XLS документы,
             * изза них будет возникать ошибка, поэтому в шаблоне их необходимо убрать.
            */
            XSSFFormulaEvaluator ev = new XSSFFormulaEvaluator(workBook);
            ev.evaluateAll();
            
            fos = new FileOutputStream(new File(targetXLSFileName));
            workBook.write(fos);
            
        }catch(Exception ex){
            if(fis != null){
                try {
                    fis.close();
                } catch (IOException ex1) {
                    System.out.println("ЭТОГО НЕ ДОЛЖНО БЫТЬ! "+ex1.getMessage());
                }
            }
            if(fos != null){
                try {
                    fos.flush();
                    fos.close();
                } catch (IOException ex1) {
                    System.out.println("ЭТОГО НЕ ДОЛЖНО БЫТЬ! "+ex1.getMessage());
                }
            }
            throw ex;
        }finally{
            if(fis != null){
                try {
                    fis.close();
                } catch (IOException ex1) {
                    System.out.println("ЭТОГО НЕ ДОЛЖНО БЫТЬ! "+ex1.getMessage());
                }
            }
            if(fos != null){
                try {
                    fos.flush();
                    fos.close();
                } catch (IOException ex1) {
                    System.out.println("ЭТОГО НЕ ДОЛЖНО БЫТЬ! "+ex1.getMessage());
                }
            }
        }
    }
    /*
    Метод записывает данные в Excel    
    */
    private static void writeDataInSheetOfTemplate(XSSFWorkbook workBook, 
                                                    Map<String,List<Map<String,String>>> dataMap, 
                                                    Map<String, Integer> componentCountMap,
                                                    Map<String, String> fieldsMappingMap, 
                                                    Map<String, Integer> templateFieldsMapping, 
                                                    String sheetName, int startIdx, Boolean enterComponentName){
        XSSFSheet sheet = workBook.getSheet(sheetName);
        Row row;
        int counter = 0;
        String componentName;
        List<Map<String,String>> dataList;
        Boolean isFirstCell = true;
        int componentsCount = 0;
        
        Iterator iterator = dataMap.entrySet().iterator();
        while(iterator.hasNext()){
            Map.Entry<String, List<Map<String,String>>> entry = (Map.Entry)iterator.next();
            componentName = entry.getKey();
            dataList = entry.getValue();
            if(componentCountMap != null){
                componentsCount = componentCountMap.get(componentName);
            }else{
                componentsCount = 1;
            }
            for(int j = 0; j < componentsCount; j++){
                isFirstCell = true;
                for (int i = 0; i < dataList.size(); i++) {                
                    row = sheet.getRow(i+startIdx+counter);
                    if(row == null){
                        row = sheet.createRow(i+startIdx);
                    }
                    Map<String, String> data = dataList.get(i);
                    Cell cell;
                    //проходимся по мапе с мапингом полей необходимых для перенесения в шаблон
                    Iterator mappingIterator = fieldsMappingMap.entrySet().iterator();

                    while(mappingIterator.hasNext()){                    
                        //для каждого вхождения в маппинг
                        // key - название поля в шаблоне
                        // value - название поля в БД(XLS)
                        Map.Entry<String, String> mappingEntry = (Map.Entry)mappingIterator.next();
                        //получаем значение поля из БД
                        String fieldValue = data.get(mappingEntry.getValue());
                        //из мапинга названия полей к их номеру в шаблоне вытягиваем номер ячейки 
                        // key - название поля в шаблоне
                        // value - порядковый номер столбца в шаблоне (0,1,2,3 и т.д.)
                        Integer fieldNum = (Integer)templateFieldsMapping.get(mappingEntry.getKey());                    
                        cell = row.getCell(fieldNum);
                        //записываем в ячейку значение из БД
                        if(enterComponentName && isFirstCell && mappingEntry.getKey().equals("наименование дет.")){
                            cell.setCellValue(componentName.substring(0, componentName.indexOf("(Сх."))+" "+fieldValue.trim());
                            isFirstCell = false;
                        }else{
                            if(fieldValue != null && !fieldValue.trim().isEmpty()){
                                if(isNumeric(fieldValue.trim())){
                                    cell.setCellValue(Double.parseDouble(fieldValue.trim()));
                                }else{
                                    cell.setCellValue(fieldValue.trim());
                                }                                
                            }                            
                        }

                    }
                }
                counter = counter + dataList.size();
            }
        }
        
    }
    private static Integer getRowCount(Map<String,List<Map<String,String>>> dataMap){
        int result = 0;
        Iterator iterator = dataMap.entrySet().iterator();
        while(iterator.hasNext()){
            Map.Entry<String, List<Map<String,String>>> entry = (Map.Entry)iterator.next();            
            List<Map<String,String>> data = entry.getValue();
            result = result + data.size();
        }
        return result;
    }
    /*
    TODO: если появится время - провести рефакторинг и сделать методы записи и подсчета универсальными
    */
    /**
     * Подсчитывает и записывает в шаблон фурнитуру
    */
    private static void calculateAndWriteFurnitureDataInSheetOfTemplate(XSSFWorkbook workBook, 
                                                                        String furnitureSheetName,
                                                                        Map<String,List<Map<String,String>>> templateFurnitureData, 
                                                                        Map<String,List<Map<String,String>>> furnitureDataMap,
                                                                        Map<String, Integer> componentCountMap,
                                                                        int furnitureStartIdx, 
                                                                        /*int furnitureHeaderIdx, 
                                                                        int furnitureRowCount, */
                                                                        Boolean generateErrorOnMissingElement) throws Exception{                
        XSSFSheet sheet = workBook.getSheet(furnitureSheetName);
        Row row;
        Cell cell;
        String calculatedFurnitureKey;
        int calculatedFurnitureValue;
        Map<String, Integer> calculatedFurniture = calculateFurniture(furnitureDataMap, componentCountMap);        
        System.out.println(calculatedFurnitureToString(calculatedFurniture));
        Iterator iterator = calculatedFurniture.entrySet().iterator();        
        int importedRowsCount = 0;
        while (iterator.hasNext()) {
            Map.Entry<String, Integer> calcFurnRow = (Map.Entry<String, Integer>) iterator.next();
            calculatedFurnitureKey = calcFurnRow.getKey();
            calculatedFurnitureValue = calcFurnRow.getValue();
            Iterator templateFurnitureDataIterator = templateFurnitureData.entrySet().iterator();
            boolean isFindKeyInTemplate = false;
            while(templateFurnitureDataIterator.hasNext()){
                Map.Entry<String, List<Map<String,String>>> templateFurnitureDataIteratorEntry = (Map.Entry)templateFurnitureDataIterator.next();
                List<Map<String,String>> templateFurnitureRowList = templateFurnitureDataIteratorEntry.getValue();                   
                for (int i = 0; i < templateFurnitureRowList.size(); i++) {//i номер строки в Template Excel
                    Map<String,String> templateFurnitureRow = templateFurnitureRowList.get(i);
                    String templateFurnitureKey = templateFurnitureRow.get(FURNITURE_TITLE_COLUMN_NAME)+templateFurnitureRow.get(FURNITURE_ARTICLE_COLUMN_NAME)+templateFurnitureRow.get(FURNITURE_COLOR_COLUMN_NAME)+templateFurnitureRow.get(FURNITURE_SIZE_COLUMN_NAME)+templateFurnitureRow.get(FURNITURE_DIMENSION_COLUMN_NAME);

                    if(templateFurnitureKey.equals(calculatedFurnitureKey)){//Если в шаблоне найдена фурнитура из БД
                        isFindKeyInTemplate = true;
                        row = sheet.getRow(i+furnitureStartIdx);
                        cell = row.getCell(TEMPLATE_FURNITURE_COUNT_COLUMN_INDEX);
                        cell.setCellValue(calculatedFurnitureValue);
                    }                    
                }
                if(isFindKeyInTemplate == false){//если в шаблоне не найдена фурнитура из БД
                    /*
                     * TODO: вставляем в конец таблицы в TemplateExcel данную фурнитуру
                     * Доработать и продумать этот момент. А пока генерим Exception.
                     * Проблема заключается в том что нет соответствия между коллекциями с рассчитанной фурнитурой и всей фурнитурой
                     * Как вариант расширить коллекцию с рассчитанной фурнитурой - хранить в ней названия и значения полей
                     * либо разделять значения в ключе каким либо разделителем ("|" например) а затем tokenizerom разбирать эту строку (минус - будет создаваться привязка к порядку следования полей) 
                    */
                    //row = sheet.getRow(furnitureStartIdx + getRowCount(templateFurnitureData));
                    //row.getCell(TEMPLATE_FURNITURE_TITLE_COLUMN_INDEX).setCellValue(furnitureSheetName);
                    //importedRowsCount++;
                    if(generateErrorOnMissingElement){
                        throw new Exception("В шаблоне отсутствует фурнитура из БД.\n Параметры фурнитуры: "+calculatedFurnitureKey);
                    }                
                }
            }
        }
        
    }
    /**
    * Подсчитывает количество фурнитуры
    */
    private static Map<String, Integer> calculateFurniture(Map<String,List<Map<String,String>>> furnitureDataMap, Map<String, Integer> componentCountMap) throws Exception{
        Map<String, Integer> resultMap = new HashMap<String, Integer>();
        Iterator iterator = furnitureDataMap.entrySet().iterator();
        String componentName;
        int componentsCount = 0;
        while(iterator.hasNext()){
            Map.Entry<String, List<Map<String,String>>> entry = (Map.Entry)iterator.next();
            componentName = entry.getKey();
            if(componentCountMap != null){
                componentsCount = componentCountMap.get(componentName);
            }else{
                componentsCount = 1;
            }
            String countValueString;
            Integer value;
            List<Map<String,String>> furnitureRowList = entry.getValue();
            for(int j = 0; j < componentsCount; j ++){
                for (int i = 0; i < furnitureRowList.size(); i++) {
                    Map<String,String> furnitureRow = furnitureRowList.get(i);
                    String key = furnitureRow.get(FURNITURE_TITLE_COLUMN_NAME)+furnitureRow.get(FURNITURE_ARTICLE_COLUMN_NAME)+furnitureRow.get(FURNITURE_COLOR_COLUMN_NAME)+furnitureRow.get(FURNITURE_SIZE_COLUMN_NAME)+furnitureRow.get(FURNITURE_DIMENSION_COLUMN_NAME);
                    countValueString = furnitureRow.get(FURNITURE_COUNT_COLUMN_NAME).trim();
                    if(countValueString!=null && !countValueString.trim().isEmpty()){
                        
                         value = (int)Double.parseDouble(countValueString);
                    }else{
                        value = 0;
                    }
                    
                    if(resultMap.get(key)==null){
                        resultMap.put(key, value);
                    }else{
                        Integer oldValue = resultMap.get(key);
                        resultMap.put(key, value+oldValue);
                    }
                }
            }
        }
        return resultMap;
    }
    /**
     * Подсчитывает и записывает в шаблон дополнительные работы
    */
    private static void calculateAndWriteAdditionalOperationsDataInSheetOfTemplate(XSSFWorkbook workBook, 
                                                                        String additionalOperationsSheetName,
                                                                        Map<String,List<Map<String,String>>> templateAdditionalOperationsData, 
                                                                        Map<String,List<Map<String,String>>> additionalOperationsDataMap,
                                                                        Map<String, Integer> componentCountMap,
                                                                        int additionalOperationsStartIdx,                                                                         
                                                                        Boolean generateErrorOnMissingElement) throws Exception{                
        XSSFSheet sheet = workBook.getSheet(additionalOperationsSheetName);
        Row row;
        Cell cell;
        String calculatedAdditionalOperationsKey;
        double calculatedAdditionalOperationsValue;
        Map<String, Double> calculatedAdditionalOperations = calculateAdditionalOperations(additionalOperationsDataMap, componentCountMap);        
        System.out.println(calculatedFurnitureToString(calculatedAdditionalOperations));
        Iterator iterator = calculatedAdditionalOperations.entrySet().iterator();        
        int importedRowsCount = 0;
        while (iterator.hasNext()) {
            Map.Entry<String, Double> calcFurnRow = (Map.Entry<String, Double>) iterator.next();
            calculatedAdditionalOperationsKey = calcFurnRow.getKey();
            calculatedAdditionalOperationsValue = calcFurnRow.getValue();
            Iterator templateAdditionalOperationsDataIterator = templateAdditionalOperationsData.entrySet().iterator();
            boolean isFindKeyInTemplate = false;
            while(templateAdditionalOperationsDataIterator.hasNext()){
                Map.Entry<String, List<Map<String,String>>> templateAdditionalOperationsDataIteratorEntry = (Map.Entry)templateAdditionalOperationsDataIterator.next();
                List<Map<String,String>> templateAdditionalOperationsRowList = templateAdditionalOperationsDataIteratorEntry.getValue();            
                for (int i = 0; i < templateAdditionalOperationsRowList.size(); i++) {//i номер строки в Template Excel
                    Map<String,String> templateAdditionalOperationsRow = templateAdditionalOperationsRowList.get(i);
                    String templateAdditionalOperationsKey = templateAdditionalOperationsRow.get(ADDITIONAL_OPERATIONS_TITLE_COLUMN_NAME)+templateAdditionalOperationsRow.get(ADDITIONAL_OPERATIONS_DIMENSION_COLUMN_NAME);

                    if(templateAdditionalOperationsKey.equals(calculatedAdditionalOperationsKey)){//Если в шаблоне найдена операция из БД
                        isFindKeyInTemplate = true;
                        row = sheet.getRow(i+additionalOperationsStartIdx);
                        cell = row.getCell(TEMPLATE_ADDITIONAL_OPERATIONS_COUNT_COLUMN_INDEX);
                        cell.setCellValue(calculatedAdditionalOperationsValue);
                    }                    
                }
                if(isFindKeyInTemplate == false){//если в шаблоне не найдена операция из БД
                    /*
                     * TODO: вставляем в конец таблицы в TemplateExcel данную операцию
                     * Доработать и продумать этот момент. А пока генерим Exception.
                     * Проблема заключается в том что нет соответствия между коллекциями с рассчитанной операцией и всей операцией
                     * Как вариант расширить коллекцию с рассчитанной операциями - хранить в ней названия и значения полей
                     * либо разделять значения в ключе каким либо разделителем ("|" например) а затем tokenizerom разбирать эту строку (минус - будет создаваться привязка к порядку следования полей) 
                    */
                    //row = sheet.getRow(furnitureStartIdx + getRowCount(templateFurnitureData));
                    //row.getCell(TEMPLATE_FURNITURE_TITLE_COLUMN_INDEX).setCellValue(furnitureSheetName);
                    //importedRowsCount++;
                    if(generateErrorOnMissingElement){
                        throw new Exception("В шаблоне отсутствует работа из БД.\n"+calculatedAdditionalOperationsKey);
                    }                
                }
            }
        }
        
    }
    /**
    * Подсчитывает количество дополнительных работ
    */
    private static Map<String, Double> calculateAdditionalOperations(Map<String,List<Map<String,String>>> additionalOperationsDataMap, Map<String, Integer> componentCountMap) throws Exception{
        Map<String, Double> resultMap = new HashMap<String, Double>();
        Iterator iterator = additionalOperationsDataMap.entrySet().iterator();
        String componentName;
        int componentsCount = 0;
        while(iterator.hasNext()){
            Map.Entry<String, List<Map<String,String>>> entry = (Map.Entry)iterator.next();
            componentName = entry.getKey();
            if(componentCountMap != null){
                componentsCount = componentCountMap.get(componentName);
            }else{
                componentsCount = 1;
            }
            String tmpCountStr;
            double value;
            List<Map<String,String>> additionalOperationsRowList = entry.getValue();
            for(int j = 0; j < componentsCount; j ++){
                for (int i = 0; i < additionalOperationsRowList.size(); i++) {
                    Map<String,String> additionalOperationsRow = additionalOperationsRowList.get(i);
                    String key = additionalOperationsRow.get(ADDITIONAL_OPERATIONS_TITLE_COLUMN_NAME)+additionalOperationsRow.get(ADDITIONAL_OPERATIONS_DIMENSION_COLUMN_NAME);
                    tmpCountStr = additionalOperationsRow.get(ADDITIONAL_OPERATIONS_COUNT_COLUMN_NAME).trim();
                    if(tmpCountStr != null && !tmpCountStr.isEmpty()){
                        value = Double.parseDouble(tmpCountStr);
                    }else{
                        value = 0;
                    }
                    if(resultMap.get(key)==null){
                        resultMap.put(key, value);
                    }else{
                        double oldValue = resultMap.get(key);
                        resultMap.put(key, value+oldValue);
                    }
                }
            }
        }
        return resultMap;
    } 
    private static String calculatedFurnitureToString (Map<String, ?> dataMap){
        StringBuffer resultStrBuff = new StringBuffer();
        Iterator iterator  = dataMap.entrySet().iterator();
        while (iterator.hasNext()) {
            Map.Entry<String, Double> entry = (Map.Entry)iterator.next();
            resultStrBuff.append(entry.getKey()+" - "+entry.getValue()+"\n");            
        }
        return resultStrBuff.toString();
    }
    public static boolean isNumeric(String str){  
      try{  
        double d = Double.parseDouble(str);  
      }catch(NumberFormatException nfe){  
        return false;  
      }  
      return true;  
    }
}
