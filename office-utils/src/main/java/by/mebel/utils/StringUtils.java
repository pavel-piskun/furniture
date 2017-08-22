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
		String template = "{�����} * {� ���.} ({���-��(��)}) {��� ��������}";
		Map<String, String> dataMap = new HashMap<String, String>();		
		dataMap.put("1�����"," ");
		dataMap.put("�����.���.","1.0");
		dataMap.put("����.(��)","18.0");
		dataMap.put("��������","���");
		dataMap.put("���-��(��)","1.0");
		dataMap.put("��� ��������","12 ");
		dataMap.put("������(��)","655.0");
		dataMap.put("��������","�� ������");
		dataMap.put("2������"," "); 
		dataMap.put("����","�����");
		dataMap.put("1������"," "); 
		dataMap.put("���",""); 
		dataMap.put("� ���.","1.0");
		dataMap.put("2�����"," "); 
		dataMap.put("����������", "��� 1");  
		dataMap.put("������������ ���.","����������");
		dataMap.put("�����","� �������");
		dataMap.put("�����(��)","1850.0");
		
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
