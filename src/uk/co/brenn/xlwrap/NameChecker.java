/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.util.ArrayList;
import java.util.HashMap;
import uk.co.brenn.metadata.MetaDataCreator;

/**
 *
 * @author Christian
 */
public class NameChecker {

     HashMap<String,ArrayList<String>> categories;

     NameChecker() throws XLWrapException, XLWrapEOFException{
        //stem.out.println("MasterNameChecker");
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

    public void checkName (String sheetInfo, String category, String field) throws XLWrapMapException{
        //ystem.out.println("checking name");
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
