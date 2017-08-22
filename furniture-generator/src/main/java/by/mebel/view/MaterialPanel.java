package by.mebel.view;

import java.awt.Color;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import javax.swing.BorderFactory;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;

/**
 *
 * @author pavel.piskun
 */
public class MaterialPanel extends JPanel{
    public static final String MAIN_MATERIAL_NAME_COLUMN = "материал";
    public static final String MAIN_MATERIAL_COLOR_COLUMN = "цвет";
    public static final String MAIN_MATERIAL_THICKNESS_COLUMN = "толщина";
    public static final String MAIN_MATERIAL_NUMBER_COLUMN = "№ схемы";    
    
    public static final String FACING_MATERIAL_NAME_COLUMN = "материал";
    public static final String FACING_MATERIAL_COLOR_COLUMN = "цвет";
    public static final String FACING_MATERIAL_SIZE_COLUMN = "толщина/ширина";
    public static final String FACING_MATERIAL_NUMBER_COLUMN = "№ схемы";
    
    private JComboBox mainMaterialComboBox;
    private JComboBox facingMaterialComboBox;
    private JComboBox secondFacingMaterialComboBox;
    private JLabel componentNameLabel;
    //private JPanel mainPanel;
    //private JPanel facingPanel;
    public MaterialPanel(String componentName){
        this.setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
        this.setLayout(new GridBagLayout());
        componentNameLabel = new JLabel(componentName);
        mainMaterialComboBox = new JComboBox();
        facingMaterialComboBox = new JComboBox(); 
        secondFacingMaterialComboBox = new JComboBox();
        
        //mainPanel = new JPanel();
        //mainPanel.setLayout(new GridBagLayout());
        //facingPanel = new JPanel();
        GridBagConstraints c = new GridBagConstraints();
        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridx = 0;
        c.gridy = 0;
        c.gridwidth =2;
        c.weightx = 0.5;
        this.add(componentNameLabel, c);
        //this.add(new JLabel(""));//add empty label
        
        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridx = 0;
        c.gridy = 1;
        c.gridwidth =1;
        c.weightx = 0.0;
        this.add(new JLabel("Осн. м-л: "), c);
        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridx = 1;
        c.gridy = 1;
        this.add(mainMaterialComboBox, c);
        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridx = 0;
        c.gridy = 2;
        this.add(new JLabel("Обл. м-л: "),c);
        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridx = 1;
        c.gridy = 2;
        this.add(facingMaterialComboBox,c);
        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridx = 0;
        c.gridy = 3;
        this.add(new JLabel("2 Обл. м-л: "),c);
        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridx = 1;
        c.gridy = 3;
        this.add(secondFacingMaterialComboBox,c);
        
        //this.add(mainPanel);
        //this.add(facingPanel);
    }
    
    public void loadData(Map<String, List<Map<String, String>>> mainMaterialData,  Map<String, List<Map<String, String>>> facingMaterialData,  Map<String, List<Map<String, String>>> secondFacingMaterialData){
        /*
        Clean boxes
        */
        this.mainMaterialComboBox.removeAllItems();
        this.facingMaterialComboBox.removeAllItems();        
        /*
        Load main materials
        */
        Iterator mainMaterialDataIterator = mainMaterialData.entrySet().iterator();        
        String [] mainMaterialViewArray = null;
        while(mainMaterialDataIterator.hasNext()){
            Map.Entry<String, List<Map<String,String>>> mainMaterialDataIteratorEntry = (Map.Entry)mainMaterialDataIterator.next();
            List<Map<String,String>> mainMaterialRowList = mainMaterialDataIteratorEntry.getValue();                   
            mainMaterialViewArray = new String[mainMaterialRowList.size()];
            StringBuilder tempStr = new StringBuilder();
            for (int i = 0; i < mainMaterialRowList.size(); i++) {
                tempStr.delete(0, tempStr.length());
                Map<String,String> mainMaterialRow = mainMaterialRowList.get(i);
                tempStr.append(mainMaterialRow.get(MAIN_MATERIAL_NUMBER_COLUMN)).append(" ");
                tempStr.append(mainMaterialRow.get(MAIN_MATERIAL_NAME_COLUMN)).append(" ");
                tempStr.append(mainMaterialRow.get(MAIN_MATERIAL_COLOR_COLUMN)).append(" ");
                tempStr.append(mainMaterialRow.get(MAIN_MATERIAL_THICKNESS_COLUMN));
                this.mainMaterialComboBox.addItem(tempStr.toString());                
            }
        }        
        /*
        Load fasing materials
        */
        Iterator facingMaterialDataIterator = facingMaterialData.entrySet().iterator();        
        String [] facingMaterialViewArray = null;
        while(facingMaterialDataIterator.hasNext()){
            Map.Entry<String, List<Map<String,String>>> facingMaterialDataIteratorEntry = (Map.Entry)facingMaterialDataIterator.next();
            List<Map<String,String>> facingMaterialRowList = facingMaterialDataIteratorEntry.getValue();                   
            facingMaterialViewArray = new String[facingMaterialRowList.size()];
            StringBuilder tempStr = new StringBuilder();
            for (int i = 0; i < facingMaterialRowList.size(); i++) {
                tempStr.delete(0, tempStr.length());
                Map<String,String> facingMaterialRow = facingMaterialRowList.get(i);
                tempStr.append(facingMaterialRow.get(FACING_MATERIAL_NUMBER_COLUMN)).append(" ");
                tempStr.append(facingMaterialRow.get(FACING_MATERIAL_NAME_COLUMN)).append(" ");
                tempStr.append(facingMaterialRow.get(FACING_MATERIAL_COLOR_COLUMN)).append(" ");
                tempStr.append(facingMaterialRow.get(FACING_MATERIAL_SIZE_COLUMN));
                this.facingMaterialComboBox.addItem(tempStr.toString());                            
            }
        }
       /*
        Load second facing materials
        */ 
        Iterator secondFacingMaterialDataIterator = secondFacingMaterialData.entrySet().iterator();        
        String [] secondFacingMaterialViewArray = null;
        while(secondFacingMaterialDataIterator.hasNext()){
            Map.Entry<String, List<Map<String,String>>> secondFacingMaterialDataIteratorEntry = (Map.Entry)secondFacingMaterialDataIterator.next();
            List<Map<String,String>> secondFacingMaterialRowList = secondFacingMaterialDataIteratorEntry.getValue();                   
            secondFacingMaterialViewArray = new String[secondFacingMaterialRowList.size()];
            StringBuilder tempStr = new StringBuilder();
            for (int i = 0; i < secondFacingMaterialRowList.size(); i++) {
                tempStr.delete(0, tempStr.length());
                Map<String,String> secondFacingMaterialRow = secondFacingMaterialRowList.get(i);
                tempStr.append(secondFacingMaterialRow.get(FACING_MATERIAL_NUMBER_COLUMN)).append(" ");
                tempStr.append(secondFacingMaterialRow.get(FACING_MATERIAL_NAME_COLUMN)).append(" ");
                tempStr.append(secondFacingMaterialRow.get(FACING_MATERIAL_COLOR_COLUMN)).append(" ");
                tempStr.append(secondFacingMaterialRow.get(FACING_MATERIAL_SIZE_COLUMN));
                this.secondFacingMaterialComboBox.addItem(tempStr.toString());                            
            }
        }
    } 
    public int getMainMaterialSelectedIdx(){        
        return  mainMaterialComboBox.getSelectedIndex();
    }
    public int getFacingMaterialSelectedIdx(){
        return facingMaterialComboBox.getSelectedIndex();
    }
     public int getSecondFacingMaterialSelectedIdx(){
        return secondFacingMaterialComboBox.getSelectedIndex();
    }
    /*
    public MaterialPanel(String componentName, Map<String, List<Map<String, String>>> mainMaterialData,  Map<String, List<Map<String, String>>> facingMaterialData){
        
    }
    */
}
