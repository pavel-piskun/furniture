package by.mebel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTTblImpl;

import by.mebel.utils.ExcelUtils;
import by.mebel.utils.StringUtils;



public class WordWriter {
	public static final String ALLIGMENT_PROPERTY_NAME = "Alligment";
	public static final String FONT_FAMILY_PROPERTY_NAME = "FontFamily";
	public static final String FONT_SIZE_PROPERTY_NAME = "FontSize";
	
	public static final int DEFAULT_COLUMN_COUNT = 3;
	public static final String TEST_KEY = "№ поз.";
	
	private int templateRowHeight;
	private int templateColumnCount;
	private List<String> templateLines;
	private List<List<XWPFRun>> templateRunsList;
	private List<XWPFParagraph> templateParagraphs;
	
	private List<Map<String,Object>> textPropertyMapList;
	
	public static void main(String[] args){
		WordWriter wordWriter = new WordWriter();
		try {
			wordWriter.readTemplate("template.docx");
		} catch (Exception e) {
			
			e.printStackTrace();
		}
	}
	
	
	public void createWordDocument(String templateFileName ,String targetFileName, List<Map<String, String>> dataList) throws Exception{
		
		
		//загружаем шаблонный документ
		FileInputStream fis = new FileInputStream(templateFileName);
		XWPFDocument document = new XWPFDocument(fis);	
		
		XWPFTable table = document.getTables().get(0);		
		
		//получаем количество столбцов
		templateColumnCount = table.getRow(0).getTableCells().size();
		
		//получаем высоту строки
		templateRowHeight = table.getRow(0).getHeight();
		
		//System.out.println("-----"+columnCount+"-------");
		
		//получаем строки из шаблона
		XWPFTableCell templateCell = table.getRow(0).getCell(0);
		templateParagraphs = templateCell.getParagraphs();
		templateLines = new ArrayList<String>();
		templateRunsList = new ArrayList<List<XWPFRun>>();	
		textPropertyMapList = new ArrayList<Map<String,Object>>();
		for(XWPFParagraph paragraph:templateParagraphs){
			templateLines.add(paragraph.getText());
			templateRunsList.add(paragraph.getRuns());
			if(!paragraph.getText().trim().isEmpty()){
				Map<String, Object> propertyMap = new HashMap<String, Object>();
				propertyMap.put(ALLIGMENT_PROPERTY_NAME, paragraph.getAlignment());			
				propertyMap.put(FONT_FAMILY_PROPERTY_NAME, paragraph.getRuns().get(0).getFontFamily());
				propertyMap.put(FONT_SIZE_PROPERTY_NAME, paragraph.getRuns().get(0).getFontSize());
				textPropertyMapList.add(propertyMap);
			}else{
				textPropertyMapList.add(null);
			}
				
			
		}
		//вычисляем количество строк
		int countOfStickerTableRows = 0;
		if((dataList.size() % templateColumnCount) != 0){
			countOfStickerTableRows = (dataList.size() / templateColumnCount) +1;  
		}else{
			countOfStickerTableRows = (dataList.size() / templateColumnCount);
		}
		
		//удаляем первую строку
		table.removeRow(0);
		
		//Создаем новые строки и заполняем их данными
		for(int i =0;i<countOfStickerTableRows;i++){
			XWPFTableRow row = table.createRow();
			row.setHeight(templateRowHeight);
			for(int j=0; j < templateColumnCount; j++){
				if((i*templateColumnCount+j)<dataList.size()){				
					List<String> dataStringList = StringUtils.createDataStringList(templateLines, dataList.get(i*templateColumnCount+j));
					XWPFTableCell cell;
					if(i==0){					
						cell =  row.createCell();
					}else{					
						cell = row.getCell(j);
					}
					for (int k = 0; k < dataStringList.size(); k++) {
						XWPFParagraph newPara = new XWPFParagraph(cell.getCTTc().addNewP(), cell);	
		                XWPFRun run=newPara.createRun();
		                if(textPropertyMapList.get(k) != null){
		                	newPara.setAlignment((ParagraphAlignment) textPropertyMapList.get(k).get(ALLIGMENT_PROPERTY_NAME));
			                run.setFontFamily((String) textPropertyMapList.get(k).get(FONT_FAMILY_PROPERTY_NAME));
			                run.setFontSize((Integer) textPropertyMapList.get(k).get(FONT_SIZE_PROPERTY_NAME));
		                }		                
		                run.setText(dataStringList.get(k));
					}
				}
			}
		}
		
        FileOutputStream outStream = null;
        try {
            outStream = new FileOutputStream(targetFileName);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
 
        try {
            document.write(outStream);
            outStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
	
	public void readTemplate(String templateFileName)throws Exception{
		FileInputStream fis = new FileInputStream(templateFileName);
		XWPFDocument document = new XWPFDocument(fis);	
		
		/*XWPFDocument tmpDocument = new XWPFDocument();
		XWPFTable tmpTable = tmpDocument.createTable(1,4);
		*/
		final XWPFTable table = document.getTables().get(0);
		
		final XWPFTableCell cell = table.getRow(0).getCell(0);
		
		//XWPFTable innerTable = cell.getTables().get(0);
		
		int columnCount = table.getRow(0).getTableCells().size();
		
		int rowHeight = table.getRow(0).getHeight();
		System.out.println(rowHeight);
		List<XWPFParagraph> paragraphs = cell.getParagraphs();
		
		/*
		 for(XWPFParagraph paragraph: paragraphs){			
			System.out.println(paragraph.getText());
			System.out.println(paragraph.getRuns().size());
			for(XWPFRun run: paragraph.getRuns()){
				System.out.println(run.getTextPosition()+" "+run.getColor()+" "+run.getFontFamily()+" "+run.getFontSize());
			}
			
		}
		*/
		table.removeRow(0);
		for(int i=0;i<10;i++){
			XWPFTableRow row = table.createRow();
			row.setHeight(rowHeight);
			for(int j=0; j<columnCount;j++){
				if(i==0){
					XWPFTableCell newCell =  row.createCell();
					CTTbl ctTable = CTTbl.Factory.newInstance();
					ctTable.addNewTr().addNewTc();
					XWPFTable newXWPFTable = new XWPFTable(ctTable, newCell);
					newXWPFTable.getRow(0).getCell(0).setText("hello");
					newXWPFTable.getRow(0).getCell(0).setColor("FF0000");
					
					
					//newCell.insertTable(0, newXWPFTable);
					/*CTTbl newTbl = newCell.getCTTc().addNewTbl();
					CTRow ctRow = newTbl.addNewTr();
					CTTc newCt = ctRow.addNewTc();*/
					//newCt.
									
					newCell.setText("test");
					//newCell.getTables().add(tmpTable);
				}else{
					XWPFTableCell newCell =row.getCell(j); 
					newCell.setText("test");
					//newCell.getTables().add(tmpTable);
				}
				
				
			}
		}
		document.write(new FileOutputStream("test1.docx"));
	}
	
}
