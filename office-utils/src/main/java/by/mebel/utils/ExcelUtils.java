package by.mebel.utils;

import by.mebel.exceptions.FieldFormatException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
/**
 *
 * @author pavel.piskun
 */
public class ExcelUtils {

    public static final String COUNT_FIELD_NAME = "кол-во(шт)";
    public static final String FIRST_LENGTH_FIELD_NAME = "1длина";
    public static final String SECOND_LENGTH_FIELD_NAME = "2длина";
    public static final String FIRST_WEIGTH_FIELD_NAME = "1ширина";
    public static final String SECOND_WEIGTH_FIELD_NAME = "2ширина";
    public static final String COLOR_FIELD_NAME = "цвет";
    public static final String MATERIAL_FIELD_NAME = "материал";
    public static final String NUMBER_FIELD_NAME = "number";
    public static final String ORDER_FIELD_NAME = "заказ";
    public static final String INDEX_NUMBER_FIELD_NAME = "№ поз.";
    public static final String LENGTH_FIELD_NAME = "длина(мм)";
    public static final String WEIGTH_FIELD_NAME = "ширина(мм)";
    public static final String TEXTURE_FIELD_NAME = "текстура";
    public static final String PAZ_FIELD_NAME = "паз";
    public static final String COMMENT_FIELD_NAME = "примечание";
    public static final String THICKNESS_FIELD_NAME = "толщ.(мм)";
    
    
    /*
     * Remove empty rows.
     * Row is empty if count of empty fields greater than half of total count fields.
     */
    public static List<Map<String, String>> removeEmptyObjects(List<Map<String, String>> dataList) {
        List<Map<String, String>> result = new ArrayList<Map<String, String>>();
        //boolean isEmplty = false;
        int emptyFieldsCounter = 0;

        for (int i = 0; i < dataList.size(); i++) {
            Map<String, String> rowDataMap = dataList.get(i);
            Iterator iterator = rowDataMap.entrySet().iterator();
            emptyFieldsCounter = 0;
            while (iterator.hasNext()) {
                Entry<String, String> entry = (Entry) iterator.next();
                if (entry.getValue() == null || entry.getValue().trim().isEmpty()) {
                    emptyFieldsCounter++;
                }
            }
            //if count of empty fields greater than third of total count fields then 'remove' this object            
            if (emptyFieldsCounter > (rowDataMap.size() / 1.5)) {
                
            }else{
                result.add(rowDataMap);
            }
        }
        return result;
    }

    public static List<Map<String, String>> clone(List<Map<String, String>> dataList) throws FieldFormatException {
        List<Map<String, String>> result = new ArrayList<Map<String, String>>();
        //dataList = sortBy(dataList, sortFieldName);
        for (int i = 0; i < dataList.size(); i++) {
            Map<String, String> rowDataMap = dataList.get(i);
            Integer countOfclons;
            try {
                countOfclons = Integer.parseInt(rowDataMap.get(COUNT_FIELD_NAME));
                for (int j = 1; j <= countOfclons; j++) {
                    Map<String, String> clonedMap = cloneMap(rowDataMap);
                    clonedMap.put(NUMBER_FIELD_NAME, String.format("%d", j));
                    result.add(clonedMap);
                }
            } catch (Exception e) {
                throw new FieldFormatException(COUNT_FIELD_NAME, rowDataMap.get(COUNT_FIELD_NAME), rowDataMap.get(INDEX_NUMBER_FIELD_NAME), e);
            }
            

            /*
             Iterator iterator = rowDataMap.entrySet().iterator();			
             while(iterator.hasNext()){
             Entry<String, String> entry = (Entry)iterator.next();				
             if(entry.getValue()==null || entry.getValue().trim().isEmpty()){
             emptyFieldsCounter++;
             }
             }
             //if count of empty fields greater than half of total count fields then 'remove' this object
             if (emptyFieldsCounter < (rowDataMap.size() / 2)){
             result.add(rowDataMap);
             }*/
        }

        return result;
    }

    public static List<Map<String, String>> sortBy(List<Map<String, String>> dataList, final String sortFieldName) {
        Comparator<Map<String, String>> mapComparator = new Comparator<Map<String, String>>() {
            public int compare(Map<String, String> m1, Map<String, String> m2) {
                return Integer.parseInt(m2.get(sortFieldName)) - Integer.parseInt(m1.get(sortFieldName));
            }
        };

        Collections.sort(dataList, mapComparator);
        return dataList;
    }

    public static Map<String, String> cloneMap(Map<String, String> dataMap) {
        Map<String, String> resultMap = new HashMap<String, String>();

        Iterator iterator = dataMap.entrySet().iterator();

        while (iterator.hasNext()) {
            Entry<String, String> entry = (Entry) iterator.next();
            resultMap.put(entry.getKey(), entry.getValue());
        }

        return resultMap;
    }

    public static void updateEmptyFields(List<Map<String, String>> dataList) {

        for (int i = 0; i < dataList.size(); i++) {
            if (dataList.get(i).get(FIRST_LENGTH_FIELD_NAME).trim().isEmpty()) {
                updateMap(dataList.get(i), FIRST_LENGTH_FIELD_NAME, "           ");
            }
            if (dataList.get(i).get(SECOND_LENGTH_FIELD_NAME).trim().isEmpty()) {
                updateMap(dataList.get(i), SECOND_LENGTH_FIELD_NAME, "           ");
            }
            if (dataList.get(i).get(FIRST_WEIGTH_FIELD_NAME).trim().isEmpty()) {
                updateMap(dataList.get(i), FIRST_WEIGTH_FIELD_NAME, "           ");
            }
            if (dataList.get(i).get(SECOND_WEIGTH_FIELD_NAME).trim().isEmpty()) {
                updateMap(dataList.get(i), SECOND_WEIGTH_FIELD_NAME, "           ");
            }
        }
    }

    public static void updateMap(Map<String, String> dataMap, String key, String newValue) {

        for (Entry<String, String> entry : dataMap.entrySet()) {
            if (entry.getKey().equals(key)) {
                entry.setValue(newValue);
                break;
            }
        }
    }

    public static List<Map<String, String>> prepareDataList(List<Map<String, String>> dataList) throws Exception{
        // delete empty rows
        dataList = ExcelUtils.removeEmptyObjects(dataList);

        //clone and create additional field
        dataList = ExcelUtils.clone(dataList);

        //update some empty fields
        ExcelUtils.updateEmptyFields(dataList);

        return dataList;
    }

    public static List<Map<String, String>> groupByColorAndSortGroups(List<Map<String, String>> dataList, String sortFieldName) {
        List<Map<String, String>> resultDataList = new ArrayList<Map<String, String>>();
        Map<String, List<Map<String, String>>> groupedData = new HashMap<String, List<Map<String, String>>>();
        for (Iterator<Map<String, String>> iterator = dataList.iterator(); iterator.hasNext();) {
            Map<String, String> dataMap = (Map<String, String>) iterator.next();
            if (groupedData.get(dataMap.get(COLOR_FIELD_NAME)) != null) {
                groupedData.get(dataMap.get(COLOR_FIELD_NAME)).add(dataMap);
            } else {
                groupedData.put(dataMap.get(COLOR_FIELD_NAME), new ArrayList<Map<String, String>>());
                groupedData.get(dataMap.get(COLOR_FIELD_NAME)).add(dataMap);
            }
        }
        Iterator<Entry<String, List<Map<String, String>>>> iterator = groupedData.entrySet().iterator();
        while (iterator.hasNext()) {
            Entry<String, List<Map<String, String>>> entry = iterator.next();
            resultDataList.addAll(sortBy(entry.getValue(), sortFieldName));
        }
        return resultDataList;
    }

    public static Map<String, List<Map<String, String>>> groupByMaterialAndColor(List<Map<String, String>> dataList) {
        Map<String, List<Map<String, String>>> resultGroupedData = new HashMap<String, List<Map<String, String>>>();
        for (Iterator<Map<String, String>> iterator = dataList.iterator(); iterator.hasNext();) {
            Map<String, String> dataMap = (Map<String, String>) iterator.next();
            if (resultGroupedData.get(dataMap.get(MATERIAL_FIELD_NAME) + " " + dataMap.get(COLOR_FIELD_NAME) + "("+dataMap.get(THICKNESS_FIELD_NAME)+")") != null) {
                resultGroupedData.get(dataMap.get(MATERIAL_FIELD_NAME) + " " + dataMap.get(COLOR_FIELD_NAME)+ "("+dataMap.get(THICKNESS_FIELD_NAME)+")").add(dataMap);
            } else {
                resultGroupedData.put(dataMap.get(MATERIAL_FIELD_NAME) + " " + dataMap.get(COLOR_FIELD_NAME)+ "("+dataMap.get(THICKNESS_FIELD_NAME)+")", new ArrayList<Map<String, String>>());
                resultGroupedData.get(dataMap.get(MATERIAL_FIELD_NAME) + " " + dataMap.get(COLOR_FIELD_NAME)+ "("+dataMap.get(THICKNESS_FIELD_NAME)+")").add(dataMap);
            }
        }
        return resultGroupedData;
    }
}
