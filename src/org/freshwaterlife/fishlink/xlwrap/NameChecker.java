/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package org.freshwaterlife.fishlink.xlwrap;

import at.jku.xlwrap.spreadsheet.Sheet;
import java.util.ArrayList;
import java.util.HashMap;
import org.freshwaterlife.fishlink.MasterFactory;

/**
 *
 * @author Christian
 */
public class NameChecker {

     private HashMap<String,ArrayList<String>> categories;
     private HashMap<String,ArrayList<String>> subcategories;
     private HashMap<String,ArrayList<String>> constants;

     NameChecker() throws XLWrapMapException{
        Sheet masterSheet =  MasterFactory.getMasterListSheet();
        categories = new HashMap<String,ArrayList<String>>();
        int zeroColumn = -1; //-1 as getnames starts by increasing it.
        zeroColumn = getNames(masterSheet, categories, zeroColumn);
        subcategories = new HashMap<String,ArrayList<String>>();
        zeroColumn = getNames(masterSheet, subcategories, zeroColumn);
        constants = new HashMap<String,ArrayList<String>>();
        zeroColumn = getNames(masterSheet, constants, zeroColumn);
     }

     private int getNames(Sheet masterSheet, HashMap<String,ArrayList<String>> hashMap, int zeroColumn) throws XLWrapMapException{
         String rangeName;
         do {
            zeroColumn++;
            rangeName =  MasterFactory.getTextZeroBased(masterSheet, zeroColumn, 0);
            int zeroRow = 1;
            String fieldName = MasterFactory.getTextZeroBased(masterSheet, zeroColumn, zeroRow);
            ArrayList<String> feilds = new ArrayList<String>();
            while (!fieldName.isEmpty()) {
                zeroRow++;
                feilds.add(fieldName);
                fieldName = MasterFactory.getTextZeroBased(masterSheet, zeroColumn, zeroRow);
            } while (!fieldName.isEmpty());
            hashMap.put(rangeName, feilds);
         } while (!rangeName.isEmpty());
         return zeroColumn;
    }

    boolean isCategory (String field) {
        return categories.containsKey(field);
    }

    void checkName (String sheetInfo, String category, String field) throws XLWrapMapException{
        ArrayList<String> fields = categories.get(category);
        if (fields == null){
            throw new XLWrapMapException("Map used catagory "+ category + " which is not in the Master");
        }
        if (!fields.contains(field)){
            throw new XLWrapMapException(sheetInfo + " used field " + field + " in catagory "+ category +
                    " which is not in the Master");
        }
    }

    void checkSubType (String sheetInfo, String category, String subType) throws XLWrapMapException{
        ArrayList<String> subTypes = subcategories.get(category + "SubType");
        if (subTypes == null){
            for (String key: subcategories.keySet()) {
                System.out.println (key);
            }
            throw new XLWrapMapException("Map used catagory "+ category + 
                    " which does not have a subtype in the Master");
        }
        if (!subTypes.contains(subType)){
            throw new XLWrapMapException(sheetInfo + " used subtype " + subType + " for catagory "+ category +
                    " which is not in the Master");
        }
    }
}
