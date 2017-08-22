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
    public static final String DETAILS_TEMPLATE_SHEET_NAME = "�����������";
    public static final String FURNITURE_TEMPLATE_SHEET_NAME = "��������� ���.";
    public static final String MATERIALS_TEMPLATE_SHEET_NAME = "���.������";
    public static final String ADDITIONAL_OPERATIONS_TEMPLATE_SHEET_NAME = "�������� ��� ���.";
    public static final Boolean DEFAULT_GENERATE_ERROR_ON_MISSING_ELEMENT_VALUE = Boolean.TRUE;
    
    public static final String FURNITURE_TITLE_COLUMN_NAME = "������������";
    public static final String FURNITURE_ARTICLE_COLUMN_NAME = "art";
    public static final String FURNITURE_COLOR_COLUMN_NAME = "����";
    public static final String FURNITURE_SIZE_COLUMN_NAME = "����.(��)";
    public static final String FURNITURE_DIMENSION_COLUMN_NAME = "��.���.";
    public static final String FURNITURE_COUNT_COLUMN_NAME = "���-��";    
    
    public static final int TEMPLATE_FURNITURE_TITLE_COLUMN_INDEX = 1;
    public static final int TEMPLATE_FURNITURE_ARTICLE_COLUMN_INDEX = 2;
    public static final int TEMPLATE_FURNITURE_COLOR_COLUMN_INDEX = 3;
    public static final int TEMPLATE_FURNITURE_SIZE_COLUMN_INDEX = 4;
    public static final int TEMPLATE_FURNITURE_DIMENSION_COLUMN_INDEX = 5;
    public static final int TEMPLATE_FURNITURE_COUNT_COLUMN_INDEX = 6; 
    
    
    public static final String ADDITIONAL_OPERATIONS_TITLE_COLUMN_NAME = "��������";
    public static final String ADDITIONAL_OPERATIONS_DIMENSION_COLUMN_NAME = "��.���.";
    public static final String ADDITIONAL_OPERATIONS_COUNT_COLUMN_NAME = "���-��";
    
    public static final int TEMPLATE_ADDITIONAL_OPERATIONS_TITLE_COLUMN_INDEX = 1;
    public static final int TEMPLATE_ADDITIONAL_OPERATIONS_DIMENSION_COLUMN_INDEX = 2;
    public static final int TEMPLATE_ADDITIONAL_OPERATIONS_COUNT_COLUMN_INDEX = 3;
    
    
    
    /*
     * ���� ��� �������� ������������ "�������� ���� � �������" -> "����� ������� � �������"
     * ��� ������� ��������� � �������
     *  key - �������� ���� � �������
     *  value - ����� �������
    */
    public static final Map<String, Integer> TEMPLATE_FIELDS_MAPPING;
    static
    {
        TEMPLATE_FIELDS_MAPPING = new HashMap<String, Integer>();
        TEMPLATE_FIELDS_MAPPING.put("������������ ���.", new Integer(1));
        TEMPLATE_FIELDS_MAPPING.put("� ���.", 3);
        TEMPLATE_FIELDS_MAPPING.put("�����.", 7); 
        TEMPLATE_FIELDS_MAPPING.put("1 ��.", 9); 
        TEMPLATE_FIELDS_MAPPING.put("2 ��.", 10); 
        TEMPLATE_FIELDS_MAPPING.put("����� (L)mm", new Integer(8));
        TEMPLATE_FIELDS_MAPPING.put("������(W)mm", new Integer(11));
        TEMPLATE_FIELDS_MAPPING.put("1 �.", new Integer(12));
        TEMPLATE_FIELDS_MAPPING.put("2 �.", new Integer(13));
        TEMPLATE_FIELDS_MAPPING.put("���-��(��)", new Integer(16));
        TEMPLATE_FIELDS_MAPPING.put("���.", new Integer(14));
        TEMPLATE_FIELDS_MAPPING.put("����.", new Integer(15));
        TEMPLATE_FIELDS_MAPPING.put("����������", new Integer(20));
        TEMPLATE_FIELDS_MAPPING.put("���. ��.�1", new Integer(17));
        TEMPLATE_FIELDS_MAPPING.put("���. ��.�2", new Integer(18));
        TEMPLATE_FIELDS_MAPPING.put("���. ��.�3", new Integer(19));        
    }
    
    /*
     * ���� ��� �������� ������������ "�������� ���� � �������" -> "�������� ���� � ����� � �� �����������"
     * ��� ������� ��������� � �������
     *  key - �������� ���� � �������
     *  value - �������� ���� � ��(XLS)
    */
    public static final Map<String, String> FIELDS_MAPPING;
    static
    {
        FIELDS_MAPPING = new HashMap<String, String>();
        FIELDS_MAPPING.put("� ���.","� �����");
        FIELDS_MAPPING.put("������������ ���.", "������������ ���.");
        FIELDS_MAPPING.put("�����.", "�����.");        
        FIELDS_MAPPING.put("����� (L)mm", "����� (L)mm");
        FIELDS_MAPPING.put("1 ��.", "1 ��.");
        FIELDS_MAPPING.put("2 ��.", "2 ��.");
        FIELDS_MAPPING.put("������(W)mm", "������(W)mm");
        FIELDS_MAPPING.put("1 �.", "1 �.");
        FIELDS_MAPPING.put("2 �.", "2 �.");
        FIELDS_MAPPING.put("���-��(��)", "���-��(��)");
        FIELDS_MAPPING.put("���.", "���.");
        FIELDS_MAPPING.put("����.", "����.");
        FIELDS_MAPPING.put("����������", "����������");
        FIELDS_MAPPING.put("���. ��.�1", "���. ��.�1");
        FIELDS_MAPPING.put("���. ��.�2", "���. ��.�2");
        FIELDS_MAPPING.put("���. ��.�3", "���. ��.�3");        
    }
    /*
     * ���� ��� �������� ������������ "�������� ���� � ������� "�������� ���������" � �������" -> "����� ������� � �������"
     * ��� ������� ��������� � �������
     *  key - �������� ���� � �������
     *  value - ����� �������
    */
    public static final Map<String, Integer> TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING;
    static
    {
        TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING = new HashMap<String, Integer>();
        TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING.put("� ������", new Integer(0));
        TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING.put("��������", new Integer(1));
        TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING.put("����", new Integer(2));
        TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING.put("�������", new Integer(3));        
    }
    /*
     * ���� ��� �������� ������������ "�������� ���� � ������� "�������� ���������" � �������" -> "�������� ���� � ����� � �� �����������"
     * ��� ������� ��������� � �������
     *  key - �������� ���� � �������
     *  value - �������� ���� � ��(XLS)
    */
    public static final Map<String, String> MAIN_MATERIALS_FIELDS_MAPPING;
    static
    {
        MAIN_MATERIALS_FIELDS_MAPPING = new HashMap<String, String>();        
        MAIN_MATERIALS_FIELDS_MAPPING.put("� ������", "� �����");        
        MAIN_MATERIALS_FIELDS_MAPPING.put("��������", "��������");
        MAIN_MATERIALS_FIELDS_MAPPING.put("����", "����");
        MAIN_MATERIALS_FIELDS_MAPPING.put("�������", "�������");        
    }
    /*
     * ���� ��� �������� ������������ "�������� ���� � ������� "������������ ���������" � �������" -> "����� ������� � �������"
     * ��� ������� ��������� � �������
     *  key - �������� ���� � �������
     *  value - ����� �������
    */
    public static final Map<String, Integer> TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING;
    static
    {
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING = new HashMap<String, Integer>();
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING.put("�", new Integer(0));
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING.put("��������", new Integer(1));
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING.put("����", new Integer(2));
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING.put("�������/������", new Integer(3));        
        TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING.put("�����������", new Integer(4));
    }
    /*
     * ���� ��� �������� ������������ "�������� ���� � ������� "������������ ���������" � �������" -> "�������� ���� � ����� � �� �����������"
     * ��� ������� ��������� � �������
     *  key - �������� ���� � �������
     *  value - �������� ���� � ��(XLS)
    */
    public static final Map<String, String> FACING_MATERIALS_FIELDS_MAPPING;
    static
    {
        FACING_MATERIALS_FIELDS_MAPPING = new HashMap<String, String>();
        FACING_MATERIALS_FIELDS_MAPPING.put("�", "� �����");
        FACING_MATERIALS_FIELDS_MAPPING.put("��������", "��������");
        FACING_MATERIALS_FIELDS_MAPPING.put("����", "����");
        FACING_MATERIALS_FIELDS_MAPPING.put("�������/������", "�������/������");        
        FACING_MATERIALS_FIELDS_MAPPING.put("�����������", "�����������");
    }
    /*public static final Map<String, String> TEMPLATE_SHEETS_NAMES_MAPPING;
    static
    {
        TEMPLATE_SHEETS_NAMES_MAPPING = new HashMap<String, String>();
        TEMPLATE_SHEETS_NAMES_MAPPING.put("������������ ���.", "������.������");
        TEMPLATE_SHEETS_NA6MES_MAPPING.put("����� (L)mm", "�����");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("������(W)mm", "������");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("���-��(��)", "���-��");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("���.", "���");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("����.", "����.");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("����������", "����������");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("���. ��.�1", "��.1");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("���. ��.�2", "��.2");
        TEMPLATE_SHEETS_NAMES_MAPPING.put("���. ��.�3", "��.3");        
    }*/
    public static void main(String[] args) {
        String fileName = "d:\\CloudStorges\\Dropbox\\haltura\\�������������� XLS\\���� ���������.xlsx";
        String templateFileName = "d:\\CloudStorges\\Dropbox\\haltura\\�������������� XLS\\������ beta �����.xlsx";
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
        firstSheet = workbook.createSheet("����1");
        for (int i = 0; i < dataList.size(); i++) {
            dataMap = (Map<String, String>) dataList.get(i);
            row = firstSheet.createRow(i);
            //��������� ������ �������
            cell = row.createCell(0);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.ORDER_FIELD_NAME) + "," + dataMap.get(ExcelUtils.INDEX_NUMBER_FIELD_NAME)));
            //��������� �������
            cell = row.createCell(1);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.LENGTH_FIELD_NAME)));
            //��������� �������
            cell = row.createCell(2);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.WEIGTH_FIELD_NAME)));
            //��������� �������
            cell = row.createCell(3);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.COUNT_FIELD_NAME)));
            //��������� �������
            cell = row.createCell(4);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.TEXTURE_FIELD_NAME)));
            //��������� �������
            cell = row.createCell(5);
            cell.setCellValue(new HSSFRichTextString(dataMap.get(ExcelUtils.PAZ_FIELD_NAME)));
            //��������� �������
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
         * ������� ������ �� �������� "���������" � �������
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
             * ����� ������ ������� ���������� �������� ��������� � ������
             */            
            writeDataInSheetOfTemplate(workBook, mainMaterial, null, MAIN_MATERIALS_FIELDS_MAPPING, TEMPLATE_MAIN_MATERIALS_FIELDS_MAPPING, MATERIALS_TEMPLATE_SHEET_NAME, mainMaterialStartIdx, Boolean.TRUE);
            /*
             * ����� ������ �������  ���������� ������������ ��������� � ������
             */            
            writeDataInSheetOfTemplate(workBook, facingMaterial, null, FACING_MATERIALS_FIELDS_MAPPING, TEMPLATE_FACING_MATERIALS_FIELDS_MAPPING, MATERIALS_TEMPLATE_SHEET_NAME, facingMaterialStartIdx, Boolean.TRUE);
            /**
             * ����� ������  ������� ���������� ������ �� �� � ������
             */
            writeDataInSheetOfTemplate(workBook, componentsDataMap, null/*componentCountMap*/,FIELDS_MAPPING, TEMPLATE_FIELDS_MAPPING, DETAILS_TEMPLATE_SHEET_NAME, detailsSheetStartIdx, Boolean.TRUE);
            /**
             * ����� ������ ������� ������������ � ���������� ��������� � ������
             */
           calculateAndWriteFurnitureDataInSheetOfTemplate(workBook,
                                                            FURNITURE_TEMPLATE_SHEET_NAME,
                                                            templateFurnitureData,
                                                            furnitureDataMap, 
                                                            componentCountMap,
                                                            furnitureSheetStartIdx,                                                             
                                                            DEFAULT_GENERATE_ERROR_ON_MISSING_ELEMENT_VALUE);            
            /**
             * ����� ������ �������  ������������ ���������� � ���������� � ������ �������������� ��������
             */
            calculateAndWriteAdditionalOperationsDataInSheetOfTemplate(workBook, 
                                                                        ADDITIONAL_OPERATIONS_TEMPLATE_SHEET_NAME, 
                                                                        templateAdditionalOperationsDataMap, 
                                                                        additionalOperationsDataMap, 
                                                                        componentCountMap, 
                                                                        additionalOpertationsStartIdx, 
                                                                        DEFAULT_GENERATE_ERROR_ON_MISSING_ELEMENT_VALUE);
           /*
             * TODO: �������� ������ ����� ������� �������� �� ���������, �������� � ���������
             */
            /*
             * ��������� ��� ������� � ���������
             * � ������ 3.10b2 �� ����������� ���������� ������ ����������� �� ������ XLS ���������,
             * ���� ��� ����� ��������� ������, ������� � ������� �� ���������� ������.
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
                    System.out.println("����� �� ������ ����! "+ex1.getMessage());
                }
            }
            if(fos != null){
                try {
                    fos.flush();
                    fos.close();
                } catch (IOException ex1) {
                    System.out.println("����� �� ������ ����! "+ex1.getMessage());
                }
            }
            throw ex;
        }finally{
            if(fis != null){
                try {
                    fis.close();
                } catch (IOException ex1) {
                    System.out.println("����� �� ������ ����! "+ex1.getMessage());
                }
            }
            if(fos != null){
                try {
                    fos.flush();
                    fos.close();
                } catch (IOException ex1) {
                    System.out.println("����� �� ������ ����! "+ex1.getMessage());
                }
            }
        }
    }
    /*
    ����� ���������� ������ � Excel    
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
                    //���������� �� ���� � �������� ����� ����������� ��� ����������� � ������
                    Iterator mappingIterator = fieldsMappingMap.entrySet().iterator();

                    while(mappingIterator.hasNext()){                    
                        //��� ������� ��������� � �������
                        // key - �������� ���� � �������
                        // value - �������� ���� � ��(XLS)
                        Map.Entry<String, String> mappingEntry = (Map.Entry)mappingIterator.next();
                        //�������� �������� ���� �� ��
                        String fieldValue = data.get(mappingEntry.getValue());
                        //�� ������� �������� ����� � �� ������ � ������� ���������� ����� ������ 
                        // key - �������� ���� � �������
                        // value - ���������� ����� ������� � ������� (0,1,2,3 � �.�.)
                        Integer fieldNum = (Integer)templateFieldsMapping.get(mappingEntry.getKey());                    
                        cell = row.getCell(fieldNum);
                        //���������� � ������ �������� �� ��
                        if(enterComponentName && isFirstCell && mappingEntry.getKey().equals("������������ ���.")){
                            cell.setCellValue(componentName.substring(0, componentName.indexOf("(��."))+" "+fieldValue.trim());
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
    TODO: ���� �������� ����� - �������� ����������� � ������� ������ ������ � �������� ��������������
    */
    /**
     * ������������ � ���������� � ������ ���������
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
                for (int i = 0; i < templateFurnitureRowList.size(); i++) {//i ����� ������ � Template Excel
                    Map<String,String> templateFurnitureRow = templateFurnitureRowList.get(i);
                    String templateFurnitureKey = templateFurnitureRow.get(FURNITURE_TITLE_COLUMN_NAME)+templateFurnitureRow.get(FURNITURE_ARTICLE_COLUMN_NAME)+templateFurnitureRow.get(FURNITURE_COLOR_COLUMN_NAME)+templateFurnitureRow.get(FURNITURE_SIZE_COLUMN_NAME)+templateFurnitureRow.get(FURNITURE_DIMENSION_COLUMN_NAME);

                    if(templateFurnitureKey.equals(calculatedFurnitureKey)){//���� � ������� ������� ��������� �� ��
                        isFindKeyInTemplate = true;
                        row = sheet.getRow(i+furnitureStartIdx);
                        cell = row.getCell(TEMPLATE_FURNITURE_COUNT_COLUMN_INDEX);
                        cell.setCellValue(calculatedFurnitureValue);
                    }                    
                }
                if(isFindKeyInTemplate == false){//���� � ������� �� ������� ��������� �� ��
                    /*
                     * TODO: ��������� � ����� ������� � TemplateExcel ������ ���������
                     * ���������� � ��������� ���� ������. � ���� ������� Exception.
                     * �������� ����������� � ��� ��� ��� ������������ ����� ����������� � ������������ ���������� � ���� ����������
                     * ��� ������� ��������� ��������� � ������������ ���������� - ������� � ��� �������� � �������� �����
                     * ���� ��������� �������� � ����� ����� ���� ������������ ("|" ��������) � ����� tokenizerom ��������� ��� ������ (����� - ����� ����������� �������� � ������� ���������� �����) 
                    */
                    //row = sheet.getRow(furnitureStartIdx + getRowCount(templateFurnitureData));
                    //row.getCell(TEMPLATE_FURNITURE_TITLE_COLUMN_INDEX).setCellValue(furnitureSheetName);
                    //importedRowsCount++;
                    if(generateErrorOnMissingElement){
                        throw new Exception("� ������� ����������� ��������� �� ��.\n ��������� ���������: "+calculatedFurnitureKey);
                    }                
                }
            }
        }
        
    }
    /**
    * ������������ ���������� ���������
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
     * ������������ � ���������� � ������ �������������� ������
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
                for (int i = 0; i < templateAdditionalOperationsRowList.size(); i++) {//i ����� ������ � Template Excel
                    Map<String,String> templateAdditionalOperationsRow = templateAdditionalOperationsRowList.get(i);
                    String templateAdditionalOperationsKey = templateAdditionalOperationsRow.get(ADDITIONAL_OPERATIONS_TITLE_COLUMN_NAME)+templateAdditionalOperationsRow.get(ADDITIONAL_OPERATIONS_DIMENSION_COLUMN_NAME);

                    if(templateAdditionalOperationsKey.equals(calculatedAdditionalOperationsKey)){//���� � ������� ������� �������� �� ��
                        isFindKeyInTemplate = true;
                        row = sheet.getRow(i+additionalOperationsStartIdx);
                        cell = row.getCell(TEMPLATE_ADDITIONAL_OPERATIONS_COUNT_COLUMN_INDEX);
                        cell.setCellValue(calculatedAdditionalOperationsValue);
                    }                    
                }
                if(isFindKeyInTemplate == false){//���� � ������� �� ������� �������� �� ��
                    /*
                     * TODO: ��������� � ����� ������� � TemplateExcel ������ ��������
                     * ���������� � ��������� ���� ������. � ���� ������� Exception.
                     * �������� ����������� � ��� ��� ��� ������������ ����� ����������� � ������������ ��������� � ���� ���������
                     * ��� ������� ��������� ��������� � ������������ ���������� - ������� � ��� �������� � �������� �����
                     * ���� ��������� �������� � ����� ����� ���� ������������ ("|" ��������) � ����� tokenizerom ��������� ��� ������ (����� - ����� ����������� �������� � ������� ���������� �����) 
                    */
                    //row = sheet.getRow(furnitureStartIdx + getRowCount(templateFurnitureData));
                    //row.getCell(TEMPLATE_FURNITURE_TITLE_COLUMN_INDEX).setCellValue(furnitureSheetName);
                    //importedRowsCount++;
                    if(generateErrorOnMissingElement){
                        throw new Exception("� ������� ����������� ������ �� ��.\n"+calculatedAdditionalOperationsKey);
                    }                
                }
            }
        }
        
    }
    /**
    * ������������ ���������� �������������� �����
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
