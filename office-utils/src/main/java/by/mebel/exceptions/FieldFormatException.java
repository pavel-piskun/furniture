/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package by.mebel.exceptions;

/**
 *
 * @author ober
 */
public class FieldFormatException extends Exception{
    private String fieldName;
    private String fieldValue;
    private String lineNumber;
    
    public FieldFormatException(String fieldName, String fieldValue, String lineNumber, Exception ex){
        super(ex);
        this.fieldName = fieldName;
        this.fieldValue = fieldValue;        
        this.lineNumber = lineNumber;
    }
    @Override
    public String getMessage(){
        return "Ошибка при обработке значения в столбце '"+this.fieldName+"', № позиции: "+lineNumber+". Значение: "+fieldValue;
    }
}
