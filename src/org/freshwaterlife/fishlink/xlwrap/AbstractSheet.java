package org.freshwaterlife.fishlink.xlwrap;

import org.freshwaterlife.fishlink.FishLinkException;
import at.jku.xlwrap.common.Utils;
import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.map.expr.val.XLExprValue;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import org.freshwaterlife.fishlink.FishLinkConstants;
import org.freshwaterlife.fishlink.FishLinkUtils;

/**
 * Abstract superclass for both Data Sheet and Master Sheets.
 * 
 * Wraps an xlwrap Sheet.
 * <p>Finds the rows on which various items are found. See field summary for more details.
 * 
 * @author Christian
 */
public class AbstractSheet {

    /**
     * XLWrap sheet being wrapped.
     */
    Sheet sheet;

    /**
     * Flag to say this value was not yet found.
     */
    static int NOT_FOUND = -1;

    /**
     * First row of the AnnotationSheet that holds data to be copied into the RDF.
     * Uses Excel indexing.
     */
    int firstData;
    
    //Excell columns and Rows used here are Excell based and not 0 based as XLWrap uses internally
 
    /**
     * Row row of the AnnotationSheet that holds the Category
     * Uses Excel indexing.
     */
    int categoryRow = NOT_FOUND;
    /**
     * Row row of the AnnotationSheet that holds the Field
     * Uses Excel indexing.
     */
    int fieldRow = NOT_FOUND;
    /**
     * Row row of the AnnotationSheet that holds the id/Value Link
     * Uses Excel indexing.
     */
    int idValueLinkRow = NOT_FOUND;
    /**
     * Row row of the AnnotationSheet that holds the id/Value Link
     * Uses Excel indexing.
     */    
    int externalSheetRow = NOT_FOUND;
    /**
     * Row row of the AnnotationSheet that holds the Zero/Null indication
     * Uses Excel indexing.
     */        
    int ZeroNullRow = NOT_FOUND;
    
    /**
     * First row of the AnnotationSheet that holds the other Annotations/dropdowns.
     * Currently these include Factor, Unit, DerivationFactor and SubType
     * Uses Excel indexing.
     */
    int firstConstant = NOT_FOUND;
    /**
     * Last row of the AnnotationSheet that holds the other Annotations/dropdowns.
     * Currently these include Factor, Unit, DerivationFactor and SubType
     * Uses Excel indexing.
     */
    int lastConstant = NOT_FOUND;
    /**
     * Last Columnof the Data Sheet that holds the other Annotations/dropdowns.
     * Currently these include Factor, Unit, DerivationFactor and SubType
     * Uses Excel indexing.
     */
    String lastDataColumn;

    /**
     * Constructs an AbstractSheet, finding and checking all the fields.
     * 
     * @param theSheet xlwrap sheet
     * @throws FishLinkException An errors finding and checking teh fields.
     */
    AbstractSheet (Sheet theSheet) throws FishLinkException {
        sheet = theSheet;
        findAndCheckMetaSplits();
    }

    //TODO remove this
    private enum SplitType{
        NONE, CONSTANT
    }

    //TODO remove this
    private void endMataSplit(int row, SplitType splitType){
         switch (splitType){
            case NONE:
                return;
            case CONSTANT:
                lastConstant = row - 1;
        }
    }

    /**
     * Finds all the values to the fields.
     * This approach allows for some changes to the MetaMaster without having to change the code.
     * @throws FishLinkException Unexpected MetaNaster format.
     */
    private void findMetaSplits() throws FishLinkException{
        int row = 1;
        SplitType splitType = SplitType.NONE;
        do {
            String columnA = getCellValue("A",row);
           if (columnA != null){
                if (columnA.equalsIgnoreCase(FishLinkConstants.CATEGORY_LABEL)){
                    categoryRow = row;
                } else if (columnA.equalsIgnoreCase(FishLinkConstants.FIELD_LABEL)){
                    fieldRow = row;
                } else if (columnA.startsWith(FishLinkConstants.ID_VALUE_LABEL)){
                    idValueLinkRow = row;
                } else if (columnA.equalsIgnoreCase(FishLinkConstants.EXTERNAL_LABEL)){
                    externalSheetRow = row;
                } else if (columnA.equalsIgnoreCase(FishLinkConstants.ZEROS_VS_NULLS_LABEL)){
                    ZeroNullRow = row;
                } else if (columnA.contains("links")){
                    throw new FishLinkException ("Links not currently supported");
                } else if (columnA.contains(FishLinkConstants.CONSTANTS_DIVIDER)){
                    firstConstant = row + 1;
                    endMataSplit(row, splitType);
                    splitType = SplitType.CONSTANT;
                } else if (columnA.equalsIgnoreCase(FishLinkConstants.HEADER_LABEL)) {
                    endMataSplit(row, splitType);
                    firstData = row + 1;
                    return;
                } else if (columnA.isEmpty()){
                    endMataSplit(row, splitType);
                    splitType = SplitType.NONE;
                } else if (splitType == SplitType.NONE){
                    throw new FishLinkException ("Found unexpected \"" + columnA + "\" before " + FishLinkConstants.CONSTANTS_DIVIDER);                    
                }
              } else {
                   endMataSplit(row, splitType);
                   splitType = SplitType.NONE;
            }
            row++;
        } while (true); //will return out when finished
    }

    /**
     * Finds all the values to the fields, checking the critical ones.
     * 
     * This approach allows for some changes to the MetaMaster without having to change the code.
     * 
     * @throws FishLinkException Unexpected MetaNaster format.
     */
    private void findAndCheckMetaSplits() throws FishLinkException{
        lastDataColumn = FishLinkUtils.indexToAlpha(sheet.getColumns());
        findMetaSplits();
        if (categoryRow == NOT_FOUND) {
            throw new FishLinkException("Unable to find \"" + FishLinkConstants.CATEGORY_LABEL + "\" in column A.");
        }
        if (fieldRow == NOT_FOUND) {
            throw new FishLinkException("Unable to find \"" + FishLinkConstants.FIELD_LABEL + "\" in column A.");
        }
        if (idValueLinkRow == NOT_FOUND) {
            throw new FishLinkException("Unable to find \"" + FishLinkConstants.ID_VALUE_LABEL + "\" in column A.");
        }
    }

    /**
     * Wraps xlwraps getCell (based on zerobased indexes) and catches any Exceptions
     * @param col column zero based
     * @param row Zero based
     * @return xlwrap Cell
     * @throws FishLinkException An thrown Exception Wrapped. 
     */
    private Cell getZeroBasedCell(int col, int row) throws FishLinkException{
        try{
            return sheet.getCell(col, row);
        } catch (NullPointerException ex) {
            throw new FishLinkException ("Unable to find cell " + col + "  " + row + " in sheet " + 
                    sheet.getSheetInfo(), ex);
        } catch (XLWrapException ex) {
            throw new FishLinkException ("Unable to find cell " + col + "  " + row + " in sheet " + 
                    sheet.getSheetInfo(), ex);
        } catch (XLWrapEOFException ex) {
            throw new FishLinkException ("Unable to find cell " + col + "  " + row + " in sheet " + 
                    sheet.getSheetInfo(), ex);
        }        
    }
   
    /**
     * Gets a cell's value (based on zerobased indexes) and catches any Exceptions
     * @param col column zero based
     * @param row Zero based
     * @return String value of the Cells contents or null if cell was empty
     * @throws FishLinkException An thrown Exception Wrapped. 
     */
   private String getZeroBasedCellValue (int col, int actualRow) throws FishLinkException{
        Cell cell = getZeroBasedCell(col, actualRow);
        XLExprValue<?> value;
        try {
            value = Utils.getXLExprValue(cell);
        } catch (XLWrapException ex) {
            throw new FishLinkException("Unable to get value from cell. ", ex);
        }
        if (value == null){
            return null;
        }
        //remove the quotes that get added and we don't want here.
        return value.toString().replace("\"","");
    }

   /**
     * Gets a cell's value (based on Excel indexes) and catches any Exceptions
     * @param column Using Excel names
     * @param row Using Excel counting
     * @return String value of the Cells contents or null if cell was empty
     * @throws FishLinkException An thrown Exception Wrapped. 
     */
    String getCellValue (String column, int row) throws FishLinkException{
        int col = Utils.alphaToIndex(column);
        int actualRow = row - 1;
        return getZeroBasedCellValue (col, actualRow);
    }

}


