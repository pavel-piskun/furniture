package by.mebel.view;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import by.mebel.ExcelReader;
import by.mebel.ExcelWriter;
import by.mebel.WordWriter;
import by.mebel.exceptions.FieldFormatException;
import by.mebel.utils.ExcelUtils;
import by.piscoon.protection.PasswordPanel;
import by.piscoon.protection.windows.BindHelper;
import by.piscoon.protection.windows.WinRegistry;
import java.text.SimpleDateFormat;
import java.util.Date;


/**
 *
 * @author pavel.piskun
 */
public class EntryClass {

    private static JFilePicker filePicker;
    private static List<JFilePicker> filePickerList;
    private static ExcelReader excelReader;
    private static WordWriter wordWriter;
    public static final String DEFAULT_EXCEL_PAGE_NAME = "total";
    public static final String DEFAUT_SORT_FIELD_NAME = "длина(мм)";
    private static JTextField pageNameTextField;
    private static JTextField sortNameField;
    private static JFrame frame;
    private static JFileChooser fileChooser;
    private static JPanel pickerPanel;
    private static int filePickerCounter = 0;
    private static JPanel framePanel;
    private static int frameWidgth = 520;
    private static int frameHeigth = 200;
    public static JButton deleteButton;

    public static void main(String[] args) {
        /*
         * Protection
         
        try {
            String remainLounchCount = WinRegistry.readString(WinRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\\stickers\\protection", "remainLounchCount");
            if(remainLounchCount == null){
                remainLounchCount = "20";
                WinRegistry.createKey(WinRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\\stickers\\protection");
                WinRegistry.writeStringValue(WinRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\\stickers\\protection", "remainLounchCount", remainLounchCount);
                
                JTextArea textArea = new JTextArea("Осталось " + remainLounchCount+ " запусков.");
                textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                textArea.setWrapStyleWord(true);
                textArea.setLineWrap(true);
                textArea.setSize(textArea.getPreferredSize().width, 1);
                JOptionPane.showMessageDialog(null, textArea, "Внимание!", JOptionPane.WARNING_MESSAGE);
            }else{
                if((Integer.parseInt(remainLounchCount) -1) < 0 ){
                    JTextArea textArea = new JTextArea("Количество запусков для ознакомления закончилось.");
                    textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                    textArea.setWrapStyleWord(true);
                    textArea.setLineWrap(true);
                    textArea.setSize(textArea.getPreferredSize().width, 1);
                    JOptionPane.showMessageDialog(null, textArea, "Внимание!", JOptionPane.WARNING_MESSAGE);
                    System.exit(1);
                }else{
                    WinRegistry.writeStringValue(WinRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\\stickers\\protection", "remainLounchCount", new String(new Integer((Integer.parseInt(remainLounchCount)-1)).toString()));
                    JTextArea textArea = new JTextArea("Осталось " + remainLounchCount+ " запусков.");
                    textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                    textArea.setWrapStyleWord(true);
                    textArea.setLineWrap(true);
                    textArea.setSize(textArea.getPreferredSize().width, 1);
                    JOptionPane.showMessageDialog(null, textArea, "Внимание!", JOptionPane.WARNING_MESSAGE);
                }                
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
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
        /*
         * Init objects
         */
        excelReader = new ExcelReader();
        wordWriter = new WordWriter();
        filePickerList = new ArrayList<JFilePicker>();
        /*
         * 
         */
        frame = new JFrame("Stickers maker");
        JFrame.setDefaultLookAndFeelDecorated(true);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(frameWidgth, frameHeigth);
        // Get the size of the screen
        Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
        // Determine the new location of the window
        int w = frame.getSize().width;
        int h = frame.getSize().height;
        int x = (dim.width - w) / 2;
        int y = (dim.height - h) / 2;
        frame.setLocation(x, y);
        /*
         * file chooser
         */
        fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Введите имя файла для сохранения");
        fileChooser.removeChoosableFileFilter(fileChooser.getChoosableFileFilters()[0]);
        FileTypeFilter filter = new FileTypeFilter(".docx", "Microsoft Word");
        fileChooser.addChoosableFileFilter(filter);
        /*
         *Frame panel 
         */
        framePanel = new JPanel();
        framePanel.setLayout(new BorderLayout());
        /*
         * File picker list
         */
        filePickerCounter++;
        filePicker = new JFilePicker("Excel файл " + filePickerCounter, "Выбор...");
        filePicker.setMode(JFilePicker.MODE_OPEN);
        filePicker.addFileTypeFilter(".xlsx", "Microsoft EXCEL");        
        filePicker.addFileTypeFilter(".xltm", "Microsoft EXCEL");
        filePicker.addFileTypeFilter(".xlsm", "Microsoft EXCEL");
        filePicker.addFileTypeFilter("", "Все файлы");
        filePickerList.add(filePicker);
        /*
         * picker Panel
         */
        pickerPanel = new JPanel();
        pickerPanel.setLayout(new BoxLayout(pickerPanel, BoxLayout.Y_AXIS));
        pickerPanel.add(filePicker);
        /*
         * Generate Button
         */
        JButton generateButton = new JButton("Создать");
        generateButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent evt) {
                buttonActionPerformed(evt);
            }
        });
        /*
         * Generate Button
         */
        JButton convertButton = new JButton("Преобразовать");
        convertButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent evt) {
                convertButtonActionPerformed(evt);
            }
        });
        /*
         * add panel Button
         */
        JButton addButton = new JButton("Добавить файл");
        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent evt) {
                addButtonAction(evt);
            }
        });

        deleteButton = new JButton("Удалить файл");
        deleteButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent evt) {
                deleteButtonAction(evt);
            }
        });
        deleteButton.setVisible(false);
        /*
         * Button panel
         */
        JPanel buttonPanel = new JPanel();
        buttonPanel.setLayout(new FlowLayout(FlowLayout.CENTER, 5, 5));
        buttonPanel.add(generateButton);
        buttonPanel.add(convertButton);
        buttonPanel.add(addButton);
        buttonPanel.add(deleteButton);

        /*
         * text fields
         */
        pageNameTextField = new JTextField(DEFAULT_EXCEL_PAGE_NAME, 10);
        pageNameTextField.setToolTipText("Введите название страницы");
        JLabel pageNameLabel = new JLabel("Название страницы");
        sortNameField = new JTextField(DEFAUT_SORT_FIELD_NAME, 10);
        sortNameField.setToolTipText("Введите название поля по которому будет происходить сортировка");
        JLabel sortFieldLabel = new JLabel("Поле для сортировки");
        /*
         * text field panel
         */
        JPanel textFieldPanel = new JPanel();
        textFieldPanel.setLayout(new FlowLayout(FlowLayout.CENTER, 5, 5));
        textFieldPanel.add(pageNameLabel);
        textFieldPanel.add(pageNameTextField);
        textFieldPanel.add(sortFieldLabel);
        textFieldPanel.add(sortNameField);

        framePanel.add(pickerPanel, BorderLayout.NORTH);
        framePanel.add(textFieldPanel, BorderLayout.CENTER);
        framePanel.add(buttonPanel, BorderLayout.SOUTH);
        frame.add(framePanel);
        frame.setVisible(true);
    }

    private static void buttonActionPerformed(ActionEvent evt) {
        //String filePath = filePicker.getSelectedFilePath();
        if (!checkOnEmptyFields(filePickerList)) {
            JOptionPane.showMessageDialog(null, "Выберите Excel-файл");
        } else {
            String excelFileName = null;
            JFilePicker filePicker = null;
            try {
                //List<Map<String,String>> dataList = excelReader.read(new FileInputStream(filePath), pageNameTextField.getText().trim());
                List<Map<String, String>> totalDataList = new ArrayList<Map<String, String>>();
                for (Iterator<JFilePicker> iterator = filePickerList.iterator(); iterator.hasNext();) {
                    List<Map<String, String>> tmpDataList;
                    filePicker = iterator.next();
                    excelFileName = filePicker.getSelectedFilePath();
                    tmpDataList = excelReader.read(new FileInputStream(excelFileName), pageNameTextField.getText().trim());
                    totalDataList.addAll(ExcelUtils.prepareDataList(tmpDataList));
                    totalDataList = ExcelUtils.groupByColorAndSortGroups(totalDataList, sortNameField.getText().trim());
                    if (!filePicker.lastPath.trim().equals("")) {
                        fileChooser.setCurrentDirectory(new File(filePicker.lastPath));
                    }
                }
                String wordFilePath = null;
                int userSelection = fileChooser.showSaveDialog(frame);
                if (userSelection == JFileChooser.APPROVE_OPTION) {
                    File fileToSave = fileChooser.getSelectedFile();
                   
                    if (fileToSave.getAbsolutePath().endsWith(".docx")) {
                        wordFilePath = fileToSave.getAbsolutePath();
                    } else {
                        wordFilePath = fileToSave.getAbsolutePath() + ".docx";
                    }
                    wordWriter.createWordDocument("template/template.docx", wordFilePath, totalDataList);
                    JTextArea textArea = new JTextArea("Файл с наклейками успешно создан: " + wordFilePath);
                    textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                    textArea.setWrapStyleWord(true);
                    textArea.setLineWrap(true);
                    textArea.setSize(textArea.getPreferredSize().width, 1);
                    JOptionPane.showMessageDialog(null, textArea, "Создание документа", JOptionPane.INFORMATION_MESSAGE);
                }
                
            } catch (OutOfMemoryError ex) {
                JTextArea textArea = new JTextArea("Приложению не хватает памяти для обработки файла:\n" + excelFileName);
                textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                textArea.setWrapStyleWord(true);
                textArea.setLineWrap(true);
                textArea.setSize(textArea.getPreferredSize().width, 1);
                JOptionPane.showMessageDialog(null, textArea, "Ошибка!", JOptionPane.WARNING_MESSAGE);
            } catch (FieldFormatException fex) {
                JTextArea textArea = new JTextArea("Ошибка при обработке файла:\n" + excelFileName + "\n" + fex.getMessage());
                fex.printStackTrace();
                textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                textArea.setWrapStyleWord(true);
                textArea.setLineWrap(true);
                textArea.setSize(textArea.getPreferredSize().width, 1);
                JOptionPane.showMessageDialog(null, textArea, "Ошибка!", JOptionPane.WARNING_MESSAGE);            
            } catch (Exception e) {
                JTextArea textArea = new JTextArea("Ошибка при обработке файла:\n" + excelFileName + "\n" + e.getMessage());
                e.printStackTrace();
                textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                textArea.setWrapStyleWord(true);
                textArea.setLineWrap(true);
                textArea.setSize(textArea.getPreferredSize().width, 1);
                JOptionPane.showMessageDialog(null, textArea, "Ошибка!", JOptionPane.WARNING_MESSAGE);
            }
        }
    }

    private static void convertButtonActionPerformed(ActionEvent evt) {
        if (!checkOnEmptyFields(filePickerList)) {
            JOptionPane.showMessageDialog(null, "Выберите Excel-файл");
        } else {
            String sourceExcelFileName = null;
            String targetExcelFileName = null;
            Map<String, List<Map<String, String>>> groupedData;
            try {
                //List<Map<String,String>> dataList = excelReader.read(new FileInputStream(filePath), pageNameTextField.getText().trim());
                List<Map<String, String>> totalDataList = new ArrayList<Map<String, String>>();
                for (Iterator<JFilePicker> iterator = filePickerList.iterator(); iterator.hasNext();) {
                    List<Map<String, String>> tmpDataList;
                    sourceExcelFileName = iterator.next().getSelectedFilePath();
                    tmpDataList = excelReader.read(new FileInputStream(sourceExcelFileName), pageNameTextField.getText().trim());

                    totalDataList.addAll(ExcelUtils.removeEmptyObjects(tmpDataList));
                }
                groupedData = ExcelUtils.groupByMaterialAndColor(totalDataList);
                Iterator<Map.Entry<String, List<Map<String, String>>>> dataIterator = groupedData.entrySet().iterator();
                String pathToExport = createCurrentDateDir("export");
                while (dataIterator.hasNext()) {
                    Map.Entry<String, List<Map<String, String>>> entry = dataIterator.next();
                    targetExcelFileName = entry.getKey();
                    ExcelWriter.createExcelFile(pathToExport,targetExcelFileName, entry.getValue());
                }
                JTextArea textArea = new JTextArea("Excel-файлы для программы распила успешно созданы в каталоге " + pathToExport);
                textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                textArea.setWrapStyleWord(true);
                textArea.setLineWrap(true);
                textArea.setSize(textArea.getPreferredSize().width, 1);
                JOptionPane.showMessageDialog(null, textArea, "Преобразование файлов", JOptionPane.INFORMATION_MESSAGE);
                
            } catch (OutOfMemoryError ex) {
                JTextArea textArea = new JTextArea("Приложению не хватает памяти для обработки файла:\n" + sourceExcelFileName);
                textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                textArea.setWrapStyleWord(true);
                textArea.setLineWrap(true);
                textArea.setSize(textArea.getPreferredSize().width, 1);
                JOptionPane.showMessageDialog(null, textArea, "Ошибка!", JOptionPane.WARNING_MESSAGE);
            } catch (RuntimeException e) {
                JTextArea textArea = new JTextArea("Ошибка при создании Excel файла:\n" + targetExcelFileName + ".xls\n" + e.getMessage());
                textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                textArea.setWrapStyleWord(true);
                textArea.setLineWrap(true);
                textArea.setSize(textArea.getPreferredSize().width, 1);
                JOptionPane.showMessageDialog(null, textArea, "Ошибка!", JOptionPane.WARNING_MESSAGE);
            } catch(UnknownError ue){
                JTextArea textArea = new JTextArea("Ошибка при создании каталога.\n" + ue.getMessage());
                textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                textArea.setWrapStyleWord(true);
                textArea.setLineWrap(true);
                textArea.setSize(textArea.getPreferredSize().width, 1);
                JOptionPane.showMessageDialog(null, textArea, "Ошибка!", JOptionPane.WARNING_MESSAGE);
            } catch (Exception exc) {
                JTextArea textArea = new JTextArea("Ошибка при создании файла:\n" + sourceExcelFileName + "\n" + exc.getMessage());
                textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
                textArea.setWrapStyleWord(true);
                textArea.setLineWrap(true);
                textArea.setSize(textArea.getPreferredSize().width, 1);
                JOptionPane.showMessageDialog(null, textArea, "Ошибка!", JOptionPane.WARNING_MESSAGE);
            }
        }
    }

    private static String createCurrentDateDir(String parentDir) {         
        SimpleDateFormat dateFormater = new SimpleDateFormat("dd-MM-yyyy HH-mm-ss");        
        Boolean success = new File(parentDir+"/"+dateFormater.format(new Date())).mkdirs();        
        if(success){
            return parentDir+"/"+dateFormater.format(new Date());
        }else{
            throw new UnknownError("Не удалось создать каталог: "+ parentDir+"/"+dateFormater.format(new Date()));
        }
    }

    private static void addButtonAction(ActionEvent evt) {
        filePickerCounter++;
        filePicker = new JFilePicker("Excel файл " + filePickerCounter, "Выбор...");
        filePicker.setMode(JFilePicker.MODE_OPEN);
        filePicker.addFileTypeFilter(".xlsx", "Microsoft EXCEL");
        filePicker.addFileTypeFilter(".xltm", "Microsoft EXCEL");
        filePicker.addFileTypeFilter(".xlsm", "Microsoft EXCEL");
        filePicker.addFileTypeFilter("", "Все файлы");
        pickerPanel.add(filePicker);
        filePickerList.add(filePicker);
        frameHeigth = frameHeigth + 35;
        frame.setSize(frameWidgth, frameHeigth);
        if (filePickerCounter > 1) {
            deleteButton.setVisible(true);
        }
        pickerPanel.revalidate();
        pickerPanel.repaint();
    }

    public static void deleteButtonAction(ActionEvent evt) {

        filePickerCounter--;
        pickerPanel.remove(pickerPanel.getComponentCount() - 1);
        filePickerList.remove(filePickerList.size() - 1);
        frameHeigth = frameHeigth - 35;
        frame.setSize(frameWidgth, frameHeigth);
        pickerPanel.revalidate();
        pickerPanel.repaint();
        if (filePickerCounter <= 1) {
            deleteButton.setVisible(false);
        }
    }

    public static boolean checkOnEmptyFields(List<JFilePicker> filePickerList) {
        boolean result = true;
        for (Iterator<JFilePicker> iterator = filePickerList.iterator(); iterator.hasNext();) {
            JFilePicker filePicker = iterator.next();
            String filePath = filePicker.getSelectedFilePath();
            if (filePath == null || filePath.trim().isEmpty()) {
                result = false;
            }
        }
        return result;
    }
}
