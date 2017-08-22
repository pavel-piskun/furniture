/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package by.piscoon.protection.windows;

/**
 *
 * @author pavel.piskun
 */
public class BindHelper {
    public static void main (String[] args) {
        BindHelper.deleteBinding();
    }
    
    public static boolean check(){
        boolean result = false;
        try{
            String bindingString = WinRegistry.readString(WinRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\\stickers\\protection", "binding");
            if (bindingString == null) {
                return false;
            } else {
                return true;
            }
        } catch (Exception ex) {
            ex.printStackTrace();
            return false;
        }        
    }
    
    public static void createBinding(){
        try {
            WinRegistry.createKey(WinRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\\stickers\\protection");
            WinRegistry.writeStringValue(WinRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\\stickers\\protection", "binding", "1");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
    
    public static void deleteBinding(){
        try {            
            WinRegistry.deleteValue(WinRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\\stickers\\protection", "binding");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
}
