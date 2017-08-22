/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package by.mebel.view;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.FlowLayout;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import javax.swing.BorderFactory;
import javax.swing.JCheckBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

/**
 *
 * @author Pavel.Piskun
 */
public class ChooserPanel extends JPanel{
    private JCheckBox checkBox;
    private JTextField firstSchemaCountTextField;
    private JTextField secondSchemaCountTextField;
    private JLabel componentNameLabel;
    private String componentName;
    private JPanel leftPanel;
    private JPanel rightPanel;
    
    
    public ChooserPanel(String componentName){
        //super();
        this.componentName = componentName;
        setLayout(new BorderLayout());
        setBorder(BorderFactory.createLineBorder(Color.lightGray));
        leftPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        rightPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        firstSchemaCountTextField = new JTextField(2);
        secondSchemaCountTextField = new JTextField(2);
        checkBox = new JCheckBox();
        checkBox.addItemListener(new ItemListener() {

            @Override
            public void itemStateChanged(ItemEvent e) {
                int state = e.getStateChange();
                if (state == ItemEvent.SELECTED) {
                    firstSchemaCountTextField.setEnabled(true);
                    secondSchemaCountTextField.setEnabled(true);                    
                }
                if(state == ItemEvent.DESELECTED){
                    firstSchemaCountTextField.setEnabled(false);
                    secondSchemaCountTextField.setEnabled(false);
                }
            }
        });
        
        firstSchemaCountTextField.setText("0");
        firstSchemaCountTextField.setEnabled(checkBox.isSelected());
        
        
        secondSchemaCountTextField.setText("0");
        secondSchemaCountTextField.setEnabled(checkBox.isSelected());
        
        componentNameLabel = new JLabel(componentName);
        leftPanel.add(checkBox);
        leftPanel.add(componentNameLabel);
        rightPanel.add(new JLabel("—хема 1, количество: "));
        rightPanel.add(firstSchemaCountTextField);
        
        rightPanel.add(new JLabel("—хема 2, количество: "));
        rightPanel.add(secondSchemaCountTextField);
        
        this.add(leftPanel, BorderLayout.WEST);
        this.add(rightPanel, BorderLayout.EAST);
        
    }
    public boolean isSelected(){
        return checkBox.isSelected();
    }
    public String getComponentName(){
        return this.componentName;
    }
    public Integer getFirstSchemaCount(){
        return Integer.parseInt(firstSchemaCountTextField.getText().trim());
    }
    
    public Integer getSecondSchemaCount(){
        return Integer.parseInt(secondSchemaCountTextField.getText().trim());
    }
}
