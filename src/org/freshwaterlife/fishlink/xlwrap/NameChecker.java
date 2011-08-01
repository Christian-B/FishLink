package org.freshwaterlife.fishlink.xlwrap;

import org.freshwaterlife.fishlink.FishLinkException;
import at.jku.xlwrap.spreadsheet.Sheet;
import java.util.ArrayList;
import java.util.HashMap;
import org.freshwaterlife.fishlink.FishLinkUtils;

/**
 * Wraps the MetaMaster Names List to allow vocabulary checking.
 *
 * @author Christian
 */
public class NameChecker {

    /**
     * Lists the Categories and the Fields.
     */
    private HashMap<String,ArrayList<String>> categories;
    /**
     * Lists the SubCatgerories by Category.
     */
    private HashMap<String,ArrayList<String>> subcategories;
    /**
     * List the legal values for each Constant.
     */
    private HashMap<String,ArrayList<String>> constants;

    /**
     * Constructs a nameChecker based on a List Sheet from the MetaMaster.
     *
     * Builds maps holding all the vocabulary
     *
     * @param masterListSheet  Lists Sheet from the MetaMaster
     * @throws FishLinkException Thrown if the MetaMaster is corrupt
     */
    NameChecker(Sheet masterListSheet) throws FishLinkException{
        categories = new HashMap<String,ArrayList<String>>();
        int zeroColumn = -1; //-1 as getnames starts by increasing it.
        zeroColumn = getNames(masterListSheet, categories, zeroColumn);
        subcategories = new HashMap<String,ArrayList<String>>();
        zeroColumn = getNames(masterListSheet, subcategories, zeroColumn);
        constants = new HashMap<String,ArrayList<String>>();
        zeroColumn = getNames(masterListSheet, constants, zeroColumn);
     }

    /**
     * Support function which loads a Hashmap with all the values found in the columns up to the empty 
     * @param masterListSheet  Lists Sheet from the MetaMaster
     * @param hashMap Map to be filled
     * @param zeroColumn column before first to check (ZeroBased index)
     * @return Empty column at the end of this series of columns
     * @throws FishLinkException If the MasterSheet is corrupt.
     */
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

    /**
     * Tests to see if the field name matches a category.
     * @param field Fields as in annotation sheet (Without the vocab:has)
     * @return True if and only if the field matches a Category name, otherwise False.
     */
    boolean isCategory (String field) {
        return categories.containsKey(field);
    }

    /**
     * Checks that the given field is in the given Category and corrects any case differences.
     * 
     * @param sheetInfo Info about the sheet used for error reporting only
     * @param category name of the Category
     * @param field Fields as in annotation sheet (Without the vocab:has)
     * @return Field in the case format in the vocabulary
     * @throws FishLinkException If either the Category is unknown or the Fields is not applicable to this category.
     */
    String checkField (String sheetInfo, String category, String field) throws FishLinkException{
        ArrayList<String> fields = categories.get(category);
        if (fields == null){
            throw new FishLinkException("Map used catagory "+ category + " which is not in the Master");
        }
        if (!fields.contains(field)){
            for (String theField: fields){
                if (theField.equalsIgnoreCase(field)){
                    //Close enough
                    return theField;
                }
            }
            throw new FishLinkException(sheetInfo + " used field " + field + " in catagory "+ category +
                    " which is not in the Master");
        }
        return field;
    }

    /**
     * Checks that the given subType is in the given Category
     * @param sheetInfo Info about the sheet used for error reporting only
     * @param category name of the Category
     * @param subType name of the subType
     * @throws FishLinkException If either the Category has no subTypes or the subType is not applicable to this category.
     */
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

    /**
     * Checks that the given value is part of the Constant vocabulary
     * @param sheetInfo Info about the sheet used for error reporting only
     * @param constant Name of the Constant Type
     * @param value Specific value to check
     * @throws FishLinkException IF either the Constant Type has no vocabulary or the Value is not applicable to this type.
     */
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
