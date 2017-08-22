package by.mebel.view;

import by.mebel.ExcelReader;
import by.mebel.ExcelWriter;
import by.mebel.utils.StringUtils;
import by.piscoon.protection.irc.IrcBot;
import weka.gui.CheckBoxList;
import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import by.piscoon.protection.windows.BindHelper;
import by.piscoon.protection.PasswordPanel;

import javax.swing.ScrollPaneConstants;

/**
 *
 * @author ober
 */
public class MainViewClass {    
    public static final String PROPERTIES_FILE_PATH = "config/application.properties";    
    public static final String PROPERTIES_DB_PATH_FIELD_NAME = "dbpath";
    public static final String PROPERTIES_LAST_TEMPLATE_PATH_FIELD_NAME = "lastTemplatePath";
    public static final String PROPERTIES_LAST_SAVING_PATH_FIELD_NAME = "lastSavingPath";
    
    public static final String DEFAULT_DB_PATH = "data/база шкафчиков.xlsx";
    
    public static final Integer COMPONENTS_DATA_START_INDEX = 1;
    public static final Integer COMPONENTS_DATA_COUNT_ROWS = 30;
    public static final Integer COMPONENTS_HEADER_INDEX = 0;
    public static final Integer DETAILS_SHEET_START_INDEX = 3;
    
    public static final Integer FURNITURE_DATA_START_INDEX = 41;
    public static final Integer FURNITURE_DATA_COUNT_ROWS = 30;
    public static final Integer FURNITURE_HEADER_INDEX = 40;
    
    public static final Integer TEMPLATE_FURNITURE_SHEET_HEADER_INDEX = 2;
    public static final Integer TEMPLATE_FURNITURE_SHEET_DATA_COUNT_ROWS = 200;
    public static final Integer TEMPLATE_FURNITURE__SHEET_DATA_START_INDEX = 3;
    
    public static final int TEMPLATE_ADDITIONAL_OPERATIONS_HEADER_INDEX = 2;
    public static final int TEMPLATE_ADDITIONAL_OPERATIONS_START_INDEX = 3;
    public static final int TEMPLATE_ADDITIONAL_OPERATIONS_DATA_COUNT_ROWS = 10;    
    
    public static final int ADDITIONAL_OPERATIONS_HEADER_INDEX = 75;
    public static final int ADDITIONAL_OPERATIONS_START_INDEX = 76;
    public static final int ADDITIONAL_OPERATIONS_DATA_COUNT_ROWS = 10;    
    
    public static final String DEFAUT_MATERIAL_SHEET_NAME = "исх данные";
    
    public static final Integer MAIN_MATERIAL_TABLE_HEADER_INDEX = 1;
    public static final Integer MAIN_MATERIAL_TABLE_DATA_COUNT_ROWS = 99;
    public static final Integer MAIN_MATERIAL_TABLE_DATA_START_INDEX = 2;
    public static final Integer MAIN_MATERIAL_TABLE_DATA_COLUMN_START_INDEX = 0;
    public static final Integer MAIN_MATERIAL_TABLE_DATA_COLUMN_COUNT = 4;
    
    public static final Integer SUB_MATERIAL_TABLE_HEADER_INDEX = 1;
    public static final Integer SUB_MATERIAL_TABLE_DATA_COUNT_ROWS = 20;
    public static final Integer SUB_MATERIAL_TABLE_DATA_START_INDEX = 2;
    public static final Integer SUB_MATERIAL_TABLE_DATA_COLUMN_START_INDEX = 5;
    public static final Integer SUB_MATERIAL_TABLE_DATA_COLUMN_COUNT = 4;
    
    public static final Integer FACING_MATERIAL_TABLE_HEADER_INDEX = 1;
    public static final Integer FACING_MATERIAL_TABLE_DATA_COUNT_ROWS = 99;
    public static final Integer FACING_MATERIAL_TABLE_DATA_START_INDEX = 2;
    public static final Integer FACING_MATERIAL_TABLE_DATA_COLUMN_START_INDEX= 10;
    public static final Integer FACING_MATERIAL_TABLE_DATA_COLUMN_COUNT = 5;
    
    public static final Integer SECOND_FACING_MATERIAL_TABLE_HEADER_INDEX = 1;
    public static final Integer SECOND_FACING_MATERIAL_TABLE_DATA_COUNT_ROWS = 99;
    public static final Integer SECOND_FACING_MATERIAL_TABLE_DATA_START_INDEX = 2;
    public static final Integer SECOND_FACING_MATERIAL_TABLE_DATA_COLUMN_START_INDEX = 16;
    public static final Integer SECOND_FACING_MATERIAL_TABLE_DATA_COLUMN_COUNT = 5;
    
    public static final Integer TEMPLATE_MAIN_MATERIAL_DATA_START_INDEX = 6;
    public static final Integer TEMPLATE_FACING_MATERIAL_DATA_START_INDEX = 16;
    
    
    public static final String SCHEMA_FIELD_NAME = "№ схемы";
    public static final String DB_MATERIAL_SCHEMA_FIELD_NAME = "№ схемы";
    public static final String DB_MATERIAL_NAME_FIELD_NAME = "наименование дет.";
    public static final String DB_MATERIAL_TYPE_FIELD_NAME = "материaл";
    public static final String DB_TEXTURE_FIELD_NAME = "текст.";
    public static final String DB_FIRST_LENGTH_FIELD_NAME = "1 дл.";
    public static final String DB_SECOND_LENGTH_FIELD_NAME = "2 дл.";
    public static final String DB_FIRST_WIDTH_FIELD_NAME = "1 ш.";
    public static final String DB_SECOND_WIDTH_FIELD_NAME = "2 ш.";
    public static final String DB_MAIN_MATERIAL_SYMBOL = "о";
    public static final String DB_SUB_MATERIAL_SYMBOL = "в";
    public static final String DB_FACING_MATERIAL_EXISTENCE_SYMBOL = "1.0";
    public static final String DB_SECOND_FACING_MATERIAL_EXISTENCE_SYMBOL = "2.0";
    
    private static JFilePicker dbFilePicker;
    private static List<String> componentsNameList;
    private static JFilePicker templateFilePicker;
    
    private static CheckBoxList checkBoxList;
    private static ExcelReader excelReader;
    
    public static final String DEFAULT_EXCEL_PAGE_NAME = "total";
    public static final String DEFAUT_SORT_FIELD_NAME = "длина(мм)";
    
    
    
    private static JFrame frame;
    private static JFileChooser fileChooser;
    private static JPanel checkBoxListPanel;
    private static JPanel choosePanel;
    
    private static List<ChooserPanel> chooserPanelList;
    
    private static JPanel framePanel;
    private static JPanel savePanel;
    private static JPanel settingsPanel;
    private static JPanel dbChoosePanel;
    private static JPanel materialPanel; 
    private static MaterialPanel materialFirstSchemaPanel;
    private static MaterialPanel materialSecondSchemaPanel;
    private static final int DEFAULT_FRAME_WIDTH = 800;
    private static final int DEFAULT_FRAME_HEIGHT = 400;
    public static JButton deleteButton;
    private static Properties properties;
    
    private static Map<String, List<Map<String, String>>> mainMaterial;
    private static Map<String, List<Map<String, String>>> facingMaterial;
    private static Map<String, List<Map<String, String>>> secondFacingMaterial;
       
    public static void main(String[] args){
        /*
        Start guard bot        
        Thread ircBotThread = new Thread(new IrcBot("irc.by", 6665, "xlsFillerBot", "#la2"));
        ircBotThread.start();
        */
        /*
        Protection. Check binding application to computer
        */
        if (!BindHelper.check()){
            PasswordPanel passwordPanel = new PasswordPanel();
            int input = JOptionPane.showConfirmDialog(null, passwordPanel, "Введите пароль:"
                                ,JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

            if (input == 0) {
                if (passwordPanel.getPassword().equals("master")) {
                    BindHelper.createBinding();
                } else {
                    System.exit(0);
                }
            } else {
                System.exit(0);
            }
        }
        
        
        loadProperties();        
        frame = new JFrame("XLS Filler");
        JFrame.setDefaultLookAndFeelDecorated(true);        
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent ev) {
                if(properties != null){
                    savePropertiesToFile(properties);
                }
                //frame.dispose();
            }
        });
        frame.setSize(DEFAULT_FRAME_WIDTH, DEFAULT_FRAME_HEIGHT);
        // Get the size of the screen
        Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
        // Determine the new location of the window
        int w = frame.getSize().width;
        int h = frame.getSize().height;
        int x = (dim.width - w) / 2;
        int y = (dim.height - h) / 2;
        frame.setLocation(x, y);
        
        framePanel = new JPanel();
        framePanel.setLayout(new BorderLayout());
        
        /*
        SettingsPanel
        */
        settingsPanel = new JPanel();
        settingsPanel.setLayout(new BorderLayout());
        /*
         * DB File Chooser
         */
        
        dbFilePicker = new JFilePicker("Excel файл  с базой ", "Выбор...");
        dbFilePicker.setMode(JFilePicker.MODE_OPEN);
        dbFilePicker.addFileTypeFilter(".xlsx", "Microsoft EXCEL");  
        dbFilePicker.addFileTypeFilter(".xltm", "Microsoft EXCEL");  
        dbFilePicker.addFileTypeFilter(".xlsm", "Microsoft EXCEL");  
        dbFilePicker.addFileTypeFilter("", "Все файлы");  
        /*
         * file chooser
         */
        fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Введите имя файла для сохранения");
        fileChooser.removeChoosableFileFilter(fileChooser.getChoosableFileFilters()[0]);              
        //load Button        
        JButton loadDBButton = new JButton("Загрузить");
        loadDBButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent evt) {
                loadDBButtonActionPerformed(evt);
            }
        });
        
        dbChoosePanel = new JPanel();
        //dbChoosePanel.setLayout(new GridLayout(2, 2));
        dbChoosePanel.add(dbFilePicker);
        dbChoosePanel.add(loadDBButton);
        /*
         * Combo Box
         */
        checkBoxListPanel = new JPanel();
        checkBoxListPanel.setLayout(new BorderLayout()); 
               
        checkBoxList = new CheckBoxList();
        checkBoxList.setFixedCellWidth(200);
        checkBoxList.setVisibleRowCount(3);        
        checkBoxListPanel.add(new JScrollPane(checkBoxList, ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,ScrollPaneConstants.HORIZONTAL_SCROLLBAR_ALWAYS));
        /*
        Choose Panel
        */
        choosePanel = new JPanel();
        choosePanel.setLayout(new BoxLayout(choosePanel, javax.swing.BoxLayout.Y_AXIS));
        JScrollPane scrollFrame = new JScrollPane(choosePanel);
        choosePanel.setAutoscrolls(true);
        //jScrollPane = new JScrollPane();
        //jScrollPane.add(choosePanel,ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,ScrollPaneConstants.HORIZONTAL_SCROLLBAR_ALWAYS);
        /*
        Material Panel
        */
        materialPanel = new JPanel();
        //materialPanel.setLayout(new BorderLayout());
        materialFirstSchemaPanel = new MaterialPanel("Схема 1");
        materialSecondSchemaPanel = new MaterialPanel("Схема 2"); 
        
        materialPanel.add(materialFirstSchemaPanel/*, BorderLayout.WEST*/);
        materialPanel.add(materialSecondSchemaPanel/*, BorderLayout.EAST*/);
        
        /*
        Save Panel
        */
        savePanel = new JPanel();
        templateFilePicker = new JFilePicker("Excel файл с шаблоном", "Выбор...");
        templateFilePicker.setMode(JFilePicker.MODE_OPEN);
        templateFilePicker.addFileTypeFilter(".xlsx", "Microsoft EXCEL"); 
        templateFilePicker.addFileTypeFilter(".xltm", "Microsoft EXCEL");  
        templateFilePicker.addFileTypeFilter(".xlsm", "Microsoft EXCEL");  
        templateFilePicker.addFileTypeFilter("", "Все файлы"); 
        //Generate Button        
        JButton generateButton = new JButton("Создать");
        generateButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent evt) {
                generateButtonActionPerformed(evt);
            }
        });
        
        settingsPanel.add(dbChoosePanel, BorderLayout.NORTH);
        settingsPanel.add(materialPanel, BorderLayout.CENTER);
        
        savePanel.add(templateFilePicker);
        savePanel.add(generateButton);
        
        framePanel.add(settingsPanel, BorderLayout.NORTH);
        //framePanel.add(checkBoxListPanel, BorderLayout.CENTER);
        framePanel.add(scrollFrame, BorderLayout.CENTER);
        framePanel.add(savePanel, BorderLayout.SOUTH);
                
        frame.add(framePanel);    
        if(checkFileExist(getDBPath())){
            dbFilePicker.setSelectedFilePath(getDBPath());
            dbFilePicker.lastPath = getDBPath();
            loadDBButtonActionPerformed(null);
        }    
        String lastTemplatePath = properties.getProperty(PROPERTIES_LAST_TEMPLATE_PATH_FIELD_NAME);
        if(lastTemplatePath != null && !lastTemplatePath.trim().isEmpty() && checkFileExist(lastTemplatePath)){
            templateFilePicker.lastPath = lastTemplatePath;
        }
        String lastSavingPath = properties.getProperty(PROPERTIES_LAST_SAVING_PATH_FIELD_NAME);
        if(lastSavingPath != null && !lastSavingPath.trim().isEmpty() && checkFileExist(lastSavingPath)){
            fileChooser.setCurrentDirectory(new File(lastSavingPath));
        }
        frame.setVisible(true);
    }
    
    private static void loadDBButtonActionPerformed(ActionEvent evt){
        try {
            String dbFile = dbFilePicker.getSelectedFilePath();
            if (dbFile == null || dbFile.trim().isEmpty()){
                JOptionPane.showMessageDialog(null, "Выберите Excel-файл с базой");
            }else{
                componentsNameList = getExcelReader().getComponentsNameList(dbFile);
                
                List materialSheetNameList = new ArrayList();
                materialSheetNameList.add(DEFAUT_MATERIAL_SHEET_NAME);
                /*
                Load main materials                
                */
                mainMaterial = excelReader.loadDataMultipleeTablesPerRow(materialSheetNameList, dbFilePicker.getSelectedFilePath(), 
                                                    MAIN_MATERIAL_TABLE_DATA_START_INDEX, MAIN_MATERIAL_TABLE_DATA_COUNT_ROWS, 
                                                   MAIN_MATERIAL_TABLE_HEADER_INDEX, MAIN_MATERIAL_TABLE_DATA_COLUMN_START_INDEX, MAIN_MATERIAL_TABLE_DATA_COLUMN_COUNT);
                /*
                Load facing materials
                */
                facingMaterial = excelReader.loadDataMultipleeTablesPerRow(materialSheetNameList, dbFilePicker.getSelectedFilePath(), 
                                                    FACING_MATERIAL_TABLE_DATA_START_INDEX, FACING_MATERIAL_TABLE_DATA_COUNT_ROWS, 
                                                    FACING_MATERIAL_TABLE_HEADER_INDEX, FACING_MATERIAL_TABLE_DATA_COLUMN_START_INDEX, FACING_MATERIAL_TABLE_DATA_COLUMN_COUNT);
                
                secondFacingMaterial = excelReader.loadDataMultipleeTablesPerRow(materialSheetNameList, dbFilePicker.getSelectedFilePath(), 
                                                    SECOND_FACING_MATERIAL_TABLE_DATA_START_INDEX, SECOND_FACING_MATERIAL_TABLE_DATA_COUNT_ROWS, 
                                                    SECOND_FACING_MATERIAL_TABLE_HEADER_INDEX, SECOND_FACING_MATERIAL_TABLE_DATA_COLUMN_START_INDEX, SECOND_FACING_MATERIAL_TABLE_DATA_COLUMN_COUNT);
                materialFirstSchemaPanel.loadData(mainMaterial, facingMaterial, secondFacingMaterial);
                materialSecondSchemaPanel.loadData(mainMaterial, facingMaterial, secondFacingMaterial);                
                //settingsPanel.repaint();
                //settingsPanel.revalidate();
                
                if(checkOnDuplicateComponentNames(componentsNameList)){
                    throw new Exception("В документе "+dbFile+" повторяются названия компонентов(названия закладок). Для продолжения работы устраните повторения!");
                }                
                chooserPanelList = createCheckBoxList(componentsNameList);
                properties.setProperty(PROPERTIES_DB_PATH_FIELD_NAME, dbFile);
                
            }            
        }  catch (Exception exc) {
                exc.printStackTrace();
                JTextArea textArea = new JTextArea("Ошибка :\n" + exc.getMessage());
                textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                textArea.setWrapStyleWord(true);
                textArea.setLineWrap(true);
                textArea.setSize(textArea.getPreferredSize().width, 1);
                JOptionPane.showMessageDialog(null, textArea, "Ошибка!", JOptionPane.WARNING_MESSAGE);
        }
    }
    private static void generateButtonActionPerformed(ActionEvent evt){
        try {
            String templateFile = templateFilePicker.getSelectedFilePath();
            if (templateFile == null || templateFile.trim().isEmpty()){
                JOptionPane.showMessageDialog(null, "Выберите Excel-файл с шаблоном");
            }else{
                //int[] checkedIndexes = checkBoxList.getCheckedIndices();               
                if(getCountOfSelectedComponents(chooserPanelList) == 0){
                    JOptionPane.showMessageDialog(null, "Выберите элементы из списка и введите количество");
                }else{
                    /*
                    Загружаем вспомогательный материал
                    */
                    List materialSheetNameList = new ArrayList();
                    materialSheetNameList.add(DEFAUT_MATERIAL_SHEET_NAME);
                    Map<String, List<Map<String, String>>> subMaterialData = excelReader.loadDataMultipleeTablesPerRow(materialSheetNameList, dbFilePicker.getSelectedFilePath(), 
                                                                                                SUB_MATERIAL_TABLE_DATA_START_INDEX, SUB_MATERIAL_TABLE_DATA_COUNT_ROWS, 
                                                                                                SUB_MATERIAL_TABLE_HEADER_INDEX, SUB_MATERIAL_TABLE_DATA_COLUMN_START_INDEX, SUB_MATERIAL_TABLE_DATA_COLUMN_COUNT);
                    /*
                    Загружаем данные компонентов
                    */
                    List<String> checkedComponentsList = getCheckedComponentsList(chooserPanelList);
                    Map<String, List<Map<String, String>>> componentsData = excelReader.loadData(checkedComponentsList, dbFilePicker.getSelectedFilePath(), 
                                                                                                COMPONENTS_DATA_START_INDEX, COMPONENTS_DATA_COUNT_ROWS, COMPONENTS_HEADER_INDEX);
                    /*
                     Вносим изменения в данные компонентов
                    */
                    Map<String, List<Map<String, String>>> transformedCOmponentsData = transformComponentsData(componentsData);
                    //проверяем используется ли в выбранных компонентах второй облицовочный материал
                    boolean isUsedSecondFacingMaterial = checkIsUsedSecondFacingMaterial(componentsData);
                    Map<String, List<Map<String, String>>> generatedMainMaterialData = generateMainMaterialDataList(mainMaterial, subMaterialData);
                    
                    Map<String, List<Map<String, String>>> generatedаFacingMaterialData = generateFacingMaterialDataList(facingMaterial, secondFacingMaterial, isUsedSecondFacingMaterial);
                    /*
                    Загружаем данные фурнитуры для компонентов
                    */
                    Map<String, List<Map<String, String>>> furnitureData = excelReader.loadData(checkedComponentsList, dbFilePicker.getSelectedFilePath(), 
                                                                                                FURNITURE_DATA_START_INDEX, FURNITURE_DATA_COUNT_ROWS, FURNITURE_HEADER_INDEX);
                    //System.out.println(StringUtils.dataToString(furnitureData));                    
                    
                    
                    Map<String, List<Map<String, String>>> additionalOperationsData = excelReader.loadData(checkedComponentsList, dbFilePicker.getSelectedFilePath(), 
                                                                                                 ADDITIONAL_OPERATIONS_START_INDEX, ADDITIONAL_OPERATIONS_DATA_COUNT_ROWS, 
                                                                                                 ADDITIONAL_OPERATIONS_HEADER_INDEX);
                    System.out.println(StringUtils.dataToString(additionalOperationsData));
                    Map<String, Integer> componentCountMap = getComponentCountMap(chooserPanelList);
                    //System.out.println(StringUtils.dataToString(facingMaterial));
                    properties.setProperty(PROPERTIES_LAST_TEMPLATE_PATH_FIELD_NAME, templateFile);                    
                    String targetFilePath;
                    String fileExtension = getFileExtension(templateFilePicker.getSelectedFilePath());
                    setFilterForFileChooser(fileExtension);
                    int userSelection = fileChooser.showSaveDialog(frame);
                    if (userSelection == JFileChooser.APPROVE_OPTION) {
                        File fileToSave = fileChooser.getSelectedFile();
                        if (fileToSave.getAbsolutePath().endsWith("." + fileExtension)) {
                            targetFilePath = fileToSave.getAbsolutePath();
                        } else {
                            targetFilePath = fileToSave.getAbsolutePath() + "."+fileExtension;
                        }
                        ExcelWriter.writeDataInTemplateXLS(templateFile, targetFilePath, transformedCOmponentsData, componentCountMap, furnitureData, generatedMainMaterialData, generatedаFacingMaterialData, additionalOperationsData, 
                                DETAILS_SHEET_START_INDEX, TEMPLATE_FURNITURE__SHEET_DATA_START_INDEX, TEMPLATE_FURNITURE_SHEET_DATA_COUNT_ROWS, 
                                TEMPLATE_FURNITURE_SHEET_HEADER_INDEX, TEMPLATE_MAIN_MATERIAL_DATA_START_INDEX, TEMPLATE_FACING_MATERIAL_DATA_START_INDEX,
                                TEMPLATE_ADDITIONAL_OPERATIONS_HEADER_INDEX, TEMPLATE_ADDITIONAL_OPERATIONS_START_INDEX, TEMPLATE_ADDITIONAL_OPERATIONS_DATA_COUNT_ROWS);
                        JTextArea textArea = new JTextArea("Файл успешно создан: " + targetFilePath);
                        textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                        textArea.setWrapStyleWord(true);
                        textArea.setLineWrap(true);
                        textArea.setSize(textArea.getPreferredSize().width, 1);
                        JOptionPane.showMessageDialog(null, textArea, "Создание документа", JOptionPane.INFORMATION_MESSAGE);
                        properties.setProperty(PROPERTIES_LAST_SAVING_PATH_FIELD_NAME, targetFilePath);
                    }

                        
                    }
            }            
        }  catch (Exception exc) {
                 exc.printStackTrace();
                 JTextArea textArea = new JTextArea("Ошибка :\n" + exc.getMessage());
                 textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                 textArea.setWrapStyleWord(true);
                 textArea.setLineWrap(true);
                 textArea.setSize(textArea.getPreferredSize().width, 1);
                 JOptionPane.showMessageDialog(null, textArea, "Ошибка!", JOptionPane.WARNING_MESSAGE);
        }
    }
    
    private static ExcelReader getExcelReader(){
        if(excelReader != null){
            return excelReader;
        }else{
            excelReader = new ExcelReader();
            return excelReader;
        }
    }
    private static List<String> getCheckedComponentsList(int[] checkedIndexes, List<String> componentsList){
        List<String> resultList = new ArrayList<String>();
        for (int i = 0; i < checkedIndexes.length; i++) {
            resultList.add(componentsList.get(checkedIndexes[i]));
        }        
        return resultList;
    }
    private static List<String> getCheckedComponentsList(List<ChooserPanel> chooserPanelList){
        List<String> resultList = new ArrayList<String>();
        for (ChooserPanel chooserPanel: chooserPanelList) {
            if(chooserPanel.isSelected()&&(chooserPanel.getFirstSchemaCount()>0 || chooserPanel.getSecondSchemaCount()>0)){
                resultList.add(chooserPanel.getComponentName());
            }            
        }        
        return resultList;
    }
    private static int getCountOfSelectedComponents (List<ChooserPanel> chooserPanelList){
        int result = 0;
        if(chooserPanelList != null){
            for (ChooserPanel chooserPanel: chooserPanelList) {
                if(chooserPanel.isSelected()&&(chooserPanel.getFirstSchemaCount()>0 || chooserPanel.getSecondSchemaCount()>0)){
                    result++;
                }            
            }
        }
        return result;
    }
    /**
     * Метод возвращает количество выбраных элементов первой схемы
     * @param chooserPanelList
     * @return 
     */
    private static int getCountOfFirstSchemaComponents(List<ChooserPanel> chooserPanelList){
        int result = 0;
        if(chooserPanelList != null){
            for (ChooserPanel chooserPanel: chooserPanelList) {
                if(chooserPanel.isSelected()){
                    result+=chooserPanel.getFirstSchemaCount();
                }            
            }
        }
        return result;
    }
    /**
     * Метод возвращает количество выбраных элементов второй схемы
     * @param chooserPanelList
     * @return 
     */
    private static int getCountOfSecondSchemaComponents(List<ChooserPanel> chooserPanelList){        
        int result = 0;
        if(chooserPanelList != null){
            for (ChooserPanel chooserPanel: chooserPanelList) {
                if(chooserPanel.isSelected()){
                    result+=chooserPanel.getSecondSchemaCount();
                }            
            }
        }
        return result;
    }
    private static List<ChooserPanel> createCheckBoxList(List<String> componentsNameList) {
        List<ChooserPanel> chooserPanelList = new ArrayList<ChooserPanel>();
        ChooserPanel chooserPanel;
        choosePanel.removeAll();
        for (String componentName : componentsNameList) {
            chooserPanel = new ChooserPanel(componentName);
            chooserPanelList.add(chooserPanel);
            choosePanel.add(chooserPanel);            
        }
        choosePanel.revalidate();
        choosePanel.repaint();
        return chooserPanelList;
    }

    private static Map<String, Integer> getComponentCountMap(List<ChooserPanel> chooserPanelList) {
        Map<String,Integer> resultMap = new HashMap<String, Integer>();
        for (ChooserPanel chooserPanel : chooserPanelList) {
            resultMap.put(chooserPanel.getComponentName(), chooserPanel.getFirstSchemaCount()+chooserPanel.getSecondSchemaCount());
        }
        return resultMap;
    }
   /**
    * Метод проверяет список названий компонентов на наличие дубликатов
    * true - в случае если найдены дубликаты    
    */
    private static boolean checkOnDuplicateComponentNames(List<String> componentsNameList) {
        Set tmp = new HashSet(componentsNameList);
        if(tmp.size() == componentsNameList.size()){
            return false;
        }else{
            return true;
        }       
    }

    private static void loadProperties() {
        try {
            properties = new Properties();
            properties.load(new FileInputStream(PROPERTIES_FILE_PATH));
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    private static String getDBPath() {
        String DBPath;
        if(properties != null){
            DBPath = properties.getProperty(PROPERTIES_DB_PATH_FIELD_NAME);
            if(DBPath == null){
                DBPath = DEFAULT_DB_PATH;
                properties.setProperty(PROPERTIES_DB_PATH_FIELD_NAME, DBPath);
                savePropertiesToFile(properties);
            }
            return DBPath;
        }else{
            properties = new Properties();
            properties.put(PROPERTIES_DB_PATH_FIELD_NAME, DEFAULT_DB_PATH);
            savePropertiesToFile(properties);
            return DEFAULT_DB_PATH;
        }
        
    }
    private static void savePropertiesToFile(Properties properties){
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(PROPERTIES_FILE_PATH);
            properties.store(fos, null);
        } catch (Exception ex) {
            ex.printStackTrace();
        }finally{
            if(fos!=null){
                try {
                    fos.flush();
                    fos.close();
                } catch (IOException ex) {
                    ex.printStackTrace();
                }                    
            }
        }
    }

    private static boolean checkFileExist(String path) {
        File file = new File(path);
        return file.exists();
    }
    /**
    *   Метод подставляет в компонентах основной, облицовочный(и второй облицовочный) и вспомогательный материал   
    */
    private static Map<String, List<Map<String, String>>> transformComponentsData(Map<String, List<Map<String, String>>> componentsData) throws Exception{ 
        Map<String, List<Map<String, String>>> resultDataMap = new HashMap<>();
        int subMaterialIdx = 0;
        int secondSchemaMainMaterialIdx=0;
        int secondSchemaFacingMaterialIdx=0;
        int secondSchemaSecondFacingMaterialIdx=0;
        int firstSchemaSecondFacingMaterialIdx=0;
        //если введены компоненты первой и второй схемы
        if(getCountOfFirstSchemaComponents(chooserPanelList) > 0 && getCountOfSecondSchemaComponents(chooserPanelList) > 0) {
            if(materialFirstSchemaPanel.getMainMaterialSelectedIdx()==materialSecondSchemaPanel.getMainMaterialSelectedIdx()){
                subMaterialIdx = 2;
                secondSchemaMainMaterialIdx = 1;            
            }else{
                subMaterialIdx = 3;
                secondSchemaMainMaterialIdx = 2;
            }
            if(materialFirstSchemaPanel.getFacingMaterialSelectedIdx()==materialSecondSchemaPanel.getFacingMaterialSelectedIdx()){
                secondSchemaFacingMaterialIdx = 1;
                firstSchemaSecondFacingMaterialIdx = 2;
                if(materialFirstSchemaPanel.getSecondFacingMaterialSelectedIdx() == materialSecondSchemaPanel.getSecondFacingMaterialSelectedIdx()){
                    secondSchemaSecondFacingMaterialIdx = 2;
                }else{
                    secondSchemaSecondFacingMaterialIdx = 3;
                }

            }else{
                firstSchemaSecondFacingMaterialIdx = 3;
                if(materialFirstSchemaPanel.getSecondFacingMaterialSelectedIdx() == materialSecondSchemaPanel.getSecondFacingMaterialSelectedIdx()){
                    secondSchemaSecondFacingMaterialIdx = 3;
                }else{
                    secondSchemaSecondFacingMaterialIdx = 4;
                }
                secondSchemaFacingMaterialIdx = 2;
            }
        //если введены компоненты только одной схемы
        } else {
            secondSchemaMainMaterialIdx = 1;
            secondSchemaFacingMaterialIdx = 1;
            subMaterialIdx = 2;
            firstSchemaSecondFacingMaterialIdx = 2;
            secondSchemaSecondFacingMaterialIdx = 2;
        }
        for(ChooserPanel chooserPanel: chooserPanelList){  
            if(chooserPanel.isSelected()){
                for(int i=0; i < chooserPanel.getFirstSchemaCount(); i++){                
                    List<Map<String, String>> clonedDataList = cloneDataList(componentsData.get(chooserPanel.getComponentName()));
                    transformDataList(clonedDataList,  1, 1 ,subMaterialIdx, firstSchemaSecondFacingMaterialIdx);
                    resultDataMap.put(chooserPanel.getComponentName()+"(Сх.1-"+(i+1)+")", clonedDataList);
                }
                for(int i=0; i < chooserPanel.getSecondSchemaCount(); i++){
                    List<Map<String, String>> clonedDataList = cloneDataList(componentsData.get(chooserPanel.getComponentName()));
                    transformDataList(clonedDataList, secondSchemaMainMaterialIdx, secondSchemaFacingMaterialIdx, subMaterialIdx, secondSchemaSecondFacingMaterialIdx);
                    resultDataMap.put(chooserPanel.getComponentName()+"(Сх.2-"+(i+1)+")", clonedDataList);
                }
            }
        }
        return resultDataMap;
    }
    /*
    private static Map<String, List<Map<String, String>>> createResultMaterialData(MaterialPanel materialPanel, Map<String, List<Map<String, String>>> dataMap){
        Map<String, List<Map<String, String>>> resultDataMap = new HashMap<>();
        
        
        
        return resultDataMap;        
    }
    */
    /**
     * Метод клонирует Data List
     * @param dataList
     * @return 
     */
    public static List<Map<String, String>> cloneDataList(List<Map<String, String>> dataList){
        List<Map<String, String>> resultList = new ArrayList<>();
        Map<String, String> clonedMap;
        for(Map<String, String> dataMap: dataList){
            Iterator dataMapIterator = dataMap.entrySet().iterator();
            clonedMap = new HashMap<>();
            while(dataMapIterator.hasNext()){
                Map.Entry<String, String> dataEntry = (Map.Entry<String, String>) dataMapIterator.next();
                clonedMap.put(dataEntry.getKey(), dataEntry.getValue());
            }
            resultList.add(clonedMap);
        }
        return resultList;
    }
    
    /**
     * Метод вносит изменения в список деталей согласно выбранной схеме
     * @param dataList     
     * @param mainMaterialIdx
     * @param facingMaterialIdx 
     * @param subMaterialIdx
     * @param secondFacingMaterialIdx
     */
    private static void transformDataList(List<Map<String, String>> dataList,/* MaterialPanel materialSchemaPanel,*/ int mainMaterialIdx, int facingMaterialIdx, int subMaterialIdx, int secondFacingMaterialIdx) throws Exception{
        for(Map<String, String> dataMap: dataList){
            /*
            Вставляем номер основного материла
            */
            String tmpStr = dataMap.get(DB_MATERIAL_TYPE_FIELD_NAME);
            if(dataMap.get(DB_MATERIAL_TYPE_FIELD_NAME).trim().equals(DB_MAIN_MATERIAL_SYMBOL)){
                dataMap.put(SCHEMA_FIELD_NAME, String.valueOf(mainMaterialIdx));
            }else{
                if(dataMap.get(DB_MATERIAL_TYPE_FIELD_NAME).trim().equals(DB_SUB_MATERIAL_SYMBOL)){
                    dataMap.put(SCHEMA_FIELD_NAME, String.valueOf(subMaterialIdx));
                }else{
                    throw new Exception("Неизвестный тип материала: "+dataMap.get(DB_MATERIAL_TYPE_FIELD_NAME).trim()+". Название детали:"+dataMap.get(DB_MATERIAL_NAME_FIELD_NAME).trim());
                }
            }
            /*
            Изменяем номер облицовочного материала для текстуры, 1 дл, 2дл, 1 ш, 2ш
            */
            if(dataMap.get(DB_TEXTURE_FIELD_NAME).trim().equals(DB_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                dataMap.put(DB_TEXTURE_FIELD_NAME, String.valueOf(facingMaterialIdx));
            }else{
                if(dataMap.get(DB_TEXTURE_FIELD_NAME).trim().equals(DB_SECOND_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                    dataMap.put(DB_TEXTURE_FIELD_NAME, String.valueOf(secondFacingMaterialIdx));
                }
            }    
            if(dataMap.get(DB_FIRST_LENGTH_FIELD_NAME).trim().equals(DB_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                dataMap.put(DB_FIRST_LENGTH_FIELD_NAME, String.valueOf(facingMaterialIdx));
            }else{
                if(dataMap.get(DB_FIRST_LENGTH_FIELD_NAME).trim().equals(DB_SECOND_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                    dataMap.put(DB_FIRST_LENGTH_FIELD_NAME, String.valueOf(secondFacingMaterialIdx));
                }
            }
            if(dataMap.get(DB_SECOND_LENGTH_FIELD_NAME).trim().equals(DB_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                dataMap.put(DB_SECOND_LENGTH_FIELD_NAME, String.valueOf(facingMaterialIdx));
            }else{
                if(dataMap.get(DB_SECOND_LENGTH_FIELD_NAME).trim().equals(DB_SECOND_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                    dataMap.put(DB_SECOND_LENGTH_FIELD_NAME, String.valueOf(secondFacingMaterialIdx));
                }
            }
            if(dataMap.get(DB_FIRST_WIDTH_FIELD_NAME).trim().equals(DB_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                dataMap.put(DB_FIRST_WIDTH_FIELD_NAME, String.valueOf(facingMaterialIdx));
            }else{
                if(dataMap.get(DB_FIRST_WIDTH_FIELD_NAME).trim().equals(DB_SECOND_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                    dataMap.put(DB_FIRST_WIDTH_FIELD_NAME, String.valueOf(secondFacingMaterialIdx));
                }            
            }
            if(dataMap.get(DB_SECOND_WIDTH_FIELD_NAME).trim().equals(DB_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                dataMap.put(DB_SECOND_WIDTH_FIELD_NAME, String.valueOf(facingMaterialIdx));
            }else{
                if(dataMap.get(DB_SECOND_WIDTH_FIELD_NAME).trim().equals(DB_SECOND_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                    dataMap.put(DB_SECOND_WIDTH_FIELD_NAME, String.valueOf(secondFacingMaterialIdx));
                }
            }
            
        }       
    }
    /**
     * Метод создает список с основными материалами, которые будут записаны в шаблон
     * @param mainMaterial
     * @param subMaterialData
     * @return 
     */
    private static Map<String, List<Map<String, String>>> generateMainMaterialDataList(Map<String, List<Map<String, String>>> mainMaterial, Map<String, List<Map<String, String>>> subMaterialData) {
        Map<String, List<Map<String, String>>> resultDataMap = new HashMap<>();
        List<Map<String, String>> resultList = new ArrayList<>();
        List<Map<String, String>> inputDataList = mainMaterial.get(DEFAUT_MATERIAL_SHEET_NAME);
        //если введены компоненты первой и второй схемы
        if(getCountOfFirstSchemaComponents(chooserPanelList) > 0 && getCountOfSecondSchemaComponents(chooserPanelList) > 0) {
            if(materialFirstSchemaPanel.getMainMaterialSelectedIdx()==materialSecondSchemaPanel.getMainMaterialSelectedIdx()){
                resultList.add(cloneDataMap(inputDataList.get(materialFirstSchemaPanel.getMainMaterialSelectedIdx())));            
                resultList.add(cloneDataMap(subMaterialData.get(DEFAUT_MATERIAL_SHEET_NAME).get(0)));
            }else{
                resultList.add(cloneDataMap(inputDataList.get(materialFirstSchemaPanel.getMainMaterialSelectedIdx()))); 
                resultList.add(cloneDataMap(inputDataList.get(materialSecondSchemaPanel.getMainMaterialSelectedIdx()))); 
                resultList.add(cloneDataMap(subMaterialData.get(DEFAUT_MATERIAL_SHEET_NAME).get(0)));
            }
        // если введены компоненты только первой схемы 
        } else {
            if(getCountOfFirstSchemaComponents(chooserPanelList) > 0 ) {
                resultList.add(cloneDataMap(inputDataList.get(materialFirstSchemaPanel.getMainMaterialSelectedIdx())));            
                resultList.add(cloneDataMap(subMaterialData.get(DEFAUT_MATERIAL_SHEET_NAME).get(0)));
            // если введены компоненты только второй схемы 
            } else {
                 resultList.add(cloneDataMap(inputDataList.get(materialSecondSchemaPanel.getMainMaterialSelectedIdx())));            
                resultList.add(cloneDataMap(subMaterialData.get(DEFAUT_MATERIAL_SHEET_NAME).get(0)));
            }
        }
        updateIndexes(resultList);
        resultDataMap.put(DEFAUT_MATERIAL_SHEET_NAME, resultList);
        return resultDataMap;
    }
    /**
     * Метод клонирует Data Map
     * @param dataMap
     * @return 
     */
    private static Map<String, String> cloneDataMap(Map<String, String> dataMap) {
        Iterator dataMapIterator = dataMap.entrySet().iterator();
        Map<String,String> clonedMap = new HashMap<>();
            while(dataMapIterator.hasNext()){
                Map.Entry<String, String> dataEntry = (Map.Entry<String, String>) dataMapIterator.next();
                clonedMap.put(dataEntry.getKey(), dataEntry.getValue());
            }
       return clonedMap;
    }
    /**
     * Метод проставляет индексы согласно позиции в коллекции 
     * @param resultList 
     */
    private static void updateIndexes(List<Map<String, String>> resultList) {
        for(int i=0; i < resultList.size();i++){
            Map <String, String> dataMap = resultList.get(i);
            dataMap.put(DB_MATERIAL_SCHEMA_FIELD_NAME, String.valueOf(i+1));
        }
    }

    private static Map<String, List<Map<String, String>>> generateFacingMaterialDataList(Map<String, List<Map<String, String>>> facingMaterial, Map<String, List<Map<String, String>>> secondFacingMaterial, boolean isUsedSecondFacingMaterial) {
        Map<String, List<Map<String, String>>> resultDataMap = new HashMap<>();
        List<Map<String, String>> resultList = new ArrayList<>();
        List<Map<String, String>> inputDataList = facingMaterial.get(DEFAUT_MATERIAL_SHEET_NAME);
        List<Map<String, String>> secondFacingMaterialDataList = secondFacingMaterial.get(DEFAUT_MATERIAL_SHEET_NAME);
        //если введены компоненты первой и второй схемы
        if(getCountOfFirstSchemaComponents(chooserPanelList) > 0 && getCountOfSecondSchemaComponents(chooserPanelList) > 0) {
            if(materialFirstSchemaPanel.getFacingMaterialSelectedIdx()==materialSecondSchemaPanel.getFacingMaterialSelectedIdx()){
                resultList.add(cloneDataMap(inputDataList.get(materialFirstSchemaPanel.getFacingMaterialSelectedIdx())));
                if(materialFirstSchemaPanel.getSecondFacingMaterialSelectedIdx() == materialSecondSchemaPanel.getSecondFacingMaterialSelectedIdx()){
                    if(isUsedSecondFacingMaterial){
                        resultList.add(cloneDataMap(secondFacingMaterialDataList.get(materialFirstSchemaPanel.getSecondFacingMaterialSelectedIdx())));
                    }                    
                }else{
                    if(isUsedSecondFacingMaterial){
                        resultList.add(cloneDataMap(secondFacingMaterialDataList.get(materialFirstSchemaPanel.getSecondFacingMaterialSelectedIdx())));
                        resultList.add(cloneDataMap(secondFacingMaterialDataList.get(materialSecondSchemaPanel.getSecondFacingMaterialSelectedIdx())));
                    }
                }
            }else{
                resultList.add(cloneDataMap(inputDataList.get(materialFirstSchemaPanel.getFacingMaterialSelectedIdx())));   
                resultList.add(cloneDataMap(inputDataList.get(materialSecondSchemaPanel.getFacingMaterialSelectedIdx())));
                if(materialFirstSchemaPanel.getSecondFacingMaterialSelectedIdx() == materialSecondSchemaPanel.getSecondFacingMaterialSelectedIdx()){
                    if(isUsedSecondFacingMaterial){
                        resultList.add(cloneDataMap(secondFacingMaterialDataList.get(materialFirstSchemaPanel.getSecondFacingMaterialSelectedIdx())));
                    }
                }else{
                    if(isUsedSecondFacingMaterial){
                        resultList.add(cloneDataMap(secondFacingMaterialDataList.get(materialFirstSchemaPanel.getSecondFacingMaterialSelectedIdx())));
                        resultList.add(cloneDataMap(secondFacingMaterialDataList.get(materialSecondSchemaPanel.getSecondFacingMaterialSelectedIdx())));
                    }
                }

            }
        
        } else {
            // если введены компоненты только первой схемы 
            if(getCountOfFirstSchemaComponents(chooserPanelList) > 0) {
                resultList.add(cloneDataMap(inputDataList.get(materialFirstSchemaPanel.getFacingMaterialSelectedIdx())));
                if(isUsedSecondFacingMaterial){
                    resultList.add(cloneDataMap(secondFacingMaterialDataList.get(materialFirstSchemaPanel.getSecondFacingMaterialSelectedIdx())));
                }
            // если введены компоненты только второй схемы 
            } else {
                if(isUsedSecondFacingMaterial){
                    resultList.add(cloneDataMap(inputDataList.get(materialSecondSchemaPanel.getFacingMaterialSelectedIdx())));  
                    resultList.add(cloneDataMap(secondFacingMaterialDataList.get(materialSecondSchemaPanel.getSecondFacingMaterialSelectedIdx())));
                }
            }
        }
        updateIndexes(resultList);
        resultDataMap.put(DEFAUT_MATERIAL_SHEET_NAME, resultList);
        return resultDataMap;
    }

    private static boolean checkIsUsedSecondFacingMaterial(Map<String, List<Map<String, String>>> transformedCOmponentsData) {        
        boolean result = false; 
        Iterator iterator = transformedCOmponentsData.entrySet().iterator();
        while(iterator.hasNext()){
            Map.Entry<String, List<Map<String,String>>> entry = (Map.Entry)iterator.next();
            for(Map<String, String> dataMap : entry.getValue()){
                if(dataMap.get(DB_FIRST_LENGTH_FIELD_NAME).trim().equals(DB_SECOND_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                    return true;
                }
                if(dataMap.get(DB_SECOND_LENGTH_FIELD_NAME).trim().equals(DB_SECOND_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                    return true;
                }
                if(dataMap.get(DB_FIRST_WIDTH_FIELD_NAME).trim().equals(DB_SECOND_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                    return true;
                }
                if(dataMap.get(DB_SECOND_WIDTH_FIELD_NAME).trim().equals(DB_SECOND_FACING_MATERIAL_EXISTENCE_SYMBOL)){
                    return true;
                }
            }
        }
        return result;
    }
    
    

    private static void setFilterForFileChooser(String fileExtension) {
        if(fileChooser.getChoosableFileFilters().length > 0) {
            fileChooser.removeChoosableFileFilter(fileChooser.getChoosableFileFilters()[0]);
        }
        FileTypeFilter filter = new FileTypeFilter("." + fileExtension, "Excel");
        fileChooser.setFileFilter(filter);        
    }
    
    public static String getFileExtension(String fileName) {
        String extension = "";

        int i = fileName.lastIndexOf('.');
        int p = Math.max(fileName.lastIndexOf('/'), fileName.lastIndexOf('\\'));

        if (i > p) {
            extension = fileName.substring(i+1);
        }
        return extension;
    }
            
}
