/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package org.freshwaterlife.fishlink.xlwrap;

import org.freshwaterlife.fishlink.FishLinkException;
import at.jku.xlwrap.spreadsheet.Sheet;
import java.util.ArrayList;
import java.util.HashMap;
import org.freshwaterlife.fishlink.FishLinkUtils;

/**
 *
 * @author Christian
 */
public class NameChecker {

     private HashMap<String,ArrayList<String>> categories;
     private HashMap<String,ArrayList<String>> subcategories;
     private HashMap<String,ArrayList<String>> constants;

     NameChecker(Sheet masterListSheet) throws FishLinkException{
        categories = new HashMap<String,ArrayList<String>>();
        int zeroColumn = -1; //-1 as getnames starts by increasing it.
        zeroColumn = getNames(masterListSheet, categories, zeroColumn);
        subcategories = new HashMap<String,ArrayList<String>>();
        zeroColumn = getNames(masterListSheet, subcategories, zeroColumn);
        constants = new HashMap<String,ArrayList<String>>();
        zeroColumn = getNames(masterListSheet, constants, zeroColumn);
     }

     private int getNames(Sheet masterSheet, HashMap<String,ArrayList<String>> hashMap, int zeroColumn) throws FishLinkException{
         String rangeName;
         do {
            zeroColumn++;
            rangeName =  FishLinkUtils.getTextZeroBased(masterSheet, zeroColumn, 0);
            int zeroRow = 1;
            String fieldName = FishLinkUtils.getTextZeroBased(masterSheet, zeroColumn, zeroRow);
            ArrayList<String> feilds = new ArrayList<String>();
            while (!fieldName.isEmpty()) {
                zeroRow++;
                feilds.add(fieldName);
                fieldName = FishLinkUtils.getTextZeroBased(masterSheet, zeroColumn, zeroRow);
            } while (!fieldName.isEmpty());
            hashMap.put(rangeName, feilds);
         } while (!rangeName.isEmpty());
         return zeroColumn;
    }

    boolean isCategory (String field) {
        return categories.containsKey(field);
    }

    void checkName (String sheetInfo, String category, String field) throws FishLinkException{
        ArrayList<String> fields = categories.get(category);
        if (fields == null){
            throw new FishLinkException("Map used catagory "+ category + " which is not in the Master");
        }
        if (!fields.contains(field)){
            throw new FishLinkException(sheetInfo + " used field " + field + " in catagory "+ category +
                    " which is not in the Master");
        }
    }

    void checkSubType (String sheetInfo, String category, String subType) throws FishLinkException{
        ArrayList<String> subTypes = subcategories.get(category + "SubType");
        if (subTypes == null){
            throw new FishLinkException("Map used catagory "+ category + 
                    " which does not have a subtype in the Master");
        }
        if (!subTypes.contains(subType)){
            throw new FishLinkException(sheetInfo + " used subtype " + subType + " for catagory "+ category +
                    " which is not in the Master");
        }
    }

    void checkConstant (String sheetInfo, String constant, String value) throws FishLinkException{
        ArrayList<String> values = constants.get(constant);
        if (values == null){
            for (String valuex : constants.keySet()){
                System.out.println(valuex);
            }
            throw new FishLinkException("Map used constant field "+ constant + 
                    " which is not in the Master");
        }
        if (!values.contains(value)){
            throw new FishLinkException(sheetInfo + " used value " + value + " for constant "+ constant +
                    " which is not in the Master");
        }
    }
}
