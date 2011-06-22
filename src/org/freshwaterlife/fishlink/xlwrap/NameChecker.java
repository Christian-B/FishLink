/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package org.freshwaterlife.fishlink.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.util.ArrayList;
import java.util.HashMap;
import org.freshwaterlife.fishlink.metadatacreator.MetaDataCreator;

/**
 *
 * @author Christian
 */
public class NameChecker {

     private HashMap<String,ArrayList<String>> categories;

     NameChecker() throws XLWrapException, XLWrapEOFException{
        Sheet masterSheet =  MetaDataCreator.getMasterListSheet();
        categories = new HashMap<String,ArrayList<String>>();
        int zeroColumn = 0;
        String rangeName =  MetaDataCreator.getTextZeroBased(masterSheet, zeroColumn, 0);
        while (!rangeName.isEmpty()) {
            int zeroRow = 1;
            String fieldName = MetaDataCreator.getTextZeroBased(masterSheet, zeroColumn, zeroRow);
            ArrayList<String> feilds = new ArrayList<String>();
            while (!fieldName.isEmpty()) {
                zeroRow++;
                feilds.add(fieldName);
                fieldName = MetaDataCreator.getTextZeroBased(masterSheet, zeroColumn, zeroRow);
            } while (!fieldName.isEmpty());
            categories.put(rangeName, feilds);
            zeroColumn++;
            rangeName = MetaDataCreator.getTextZeroBased(masterSheet, zeroColumn, 0);
        }
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
}
