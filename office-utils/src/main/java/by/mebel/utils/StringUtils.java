package by.mebel.utils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StringUtils {
	public static void main(String[] args){
		String template = "{заказ} * {№ поз.} ({кол-во(шт)}) {код маршрута}";
		Map<String, String> dataMap = new HashMap<String, String>();		
		dataMap.put("1длина"," ");
		dataMap.put("налич.дет.","1.0");
		dataMap.put("толщ.(мм)","18.0");
		dataMap.put("материал","ДСП");
		dataMap.put("кол-во(шт)","1.0");
		dataMap.put("код маршрута","12 ");
		dataMap.put("ширина(мм)","655.0");
		dataMap.put("текстура","Не задана");
		dataMap.put("2ширина"," "); 
		dataMap.put("цвет","Серый");
		dataMap.put("1ширина"," "); 
		dataMap.put("паз",""); 
		dataMap.put("№ поз.","1.0");
		dataMap.put("2длина"," "); 
		dataMap.put("примечание", "рис 1");  
		dataMap.put("наименование дет.","столешница");
		dataMap.put("заказ","Н Масаров");
		dataMap.put("длина(мм)","1850.0");
		
		System.out.println(replaceTokens(template, dataMap));
	}
	
	public static String replaceTokens(String text,
		Map<String, String> replacements) {
		Pattern pattern = Pattern.compile("\\{(.+?)\\}");
		Matcher matcher = pattern.matcher(text);
		StringBuffer buffer = new StringBuffer();
		while (matcher.find()) {
			String replacement = replacements.get(matcher.group(1));
			if (replacement != null) {
				// matcher.appendReplacement(buffer, replacement);
				// see comment
				matcher.appendReplacement(buffer, "");
				buffer.append(replacement);
			}
		}
		matcher.appendTail(buffer);
		return buffer.toString();
	}
	
	public static List<String> createDataStringList(List<String> templateStringList, Map<String, String> dataMap){
		List<String> result = new ArrayList<String>();
		for (int i = 0; i < templateStringList.size(); i++) {
			result.add(replaceTokens(templateStringList.get(i), dataMap));
		}		
		return result;
	}
        public static String dataToString(Map<String, List<Map<String, String>>> inDataMap){
            Iterator iterator = inDataMap.entrySet().iterator();
            StringBuffer strBuff = new StringBuffer();
            while(iterator.hasNext()){
                Map.Entry<String, List<Map<String,String>>> entry = (Map.Entry)iterator.next();
                strBuff.append("================\n"+entry.getKey()+"\n===============");
                List<Map<String,String>> data = entry.getValue();
                Map<String,String> dataMap;
                for (int i = 0; i < data.size(); i++) {
                    strBuff.append("\n");
                    dataMap = data.get(i);
                    Iterator it = dataMap.entrySet().iterator();
                    while(it.hasNext()){                        
                        Map.Entry<String, String> entry1 = (Map.Entry)it.next();
                        strBuff.append(entry1.toString()+" | ");
                    }
                }
            }
            return strBuff.toString();
        }       
}
