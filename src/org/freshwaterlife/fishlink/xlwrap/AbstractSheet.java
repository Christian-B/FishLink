package org.freshwaterlife.fishlink.xlwrap;

import org.freshwaterlife.fishlink.FishLinkException;
import at.jku.xlwrap.common.Utils;
import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.map.expr.val.XLExprValue;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import org.freshwaterlife.fishlink.FishLinkUtils;

/**
 *
 * @author Christian
 */
public class AbstractSheet {

    Sheet sheet;

    int firstData;

    //Excell columns and Rows used here are Excell based and not 0 based as XLWrap uses internally
    int categoryRow = -1;
    int fieldRow = -1;
    int idTypeRow = -1;
    int externalSheetRow = -1;
    int ZeroNullRow = -1;
    int firstConstant = -1;
    int lastConstant = -1;
    String lastDataColumn;

    AbstractSheet (Sheet theSheet) throws FishLinkException {
        sheet = theSheet;
        findAndCheckMetaSplits();
    }

    private enum SplitType{
        NONE, CONSTANT
    }

    private void endMataSplit(int row, SplitType splitType){
         switch (splitType){
            case NONE:
                return;
            case CONSTANT:
                lastConstant = row - 1;
        }
    }

    private void findMetaSplits() throws FishLinkException{
        int row = 1;
        SplitType splitType = SplitType.NONE;
        do {
            String columnA = getCellValue("A",row);
           if (columnA != null){
                if (columnA.equalsIgnoreCase(Constants.CATEGORY_LABEL)){
                    categoryRow = row;
                } else if (columnA.equalsIgnoreCase(Constants.FIELD_LABEL)){
                    fieldRow = row;
                } else if (columnA.startsWith(Constants.ID_VALUE_LABEL)){
                    idTypeRow = row;
                } else if (columnA.equalsIgnoreCase(Constants.EXTERNAL_LABEL)){
                    externalSheetRow = row;
                } else if (columnA.equalsIgnoreCase(Constants.ZEROS_VS_NULLS_LABEL)){
                    ZeroNullRow = row;
                } else if (columnA.contains("links")){
                    throw new FishLinkException ("Links not currently supported");
                } else if (columnA.contains(Constants.CONSTANTS_DIVIDER)){
                    firstConstant = row + 1;
                    endMataSplit(row, splitType);
                    splitType = SplitType.CONSTANT;
                } else if (columnA.equalsIgnoreCase(Constants.HEADER_LABEL)) {
                    endMataSplit(row, splitType);
                    firstData = row + 1;
                    return;
                } else if (columnA.isEmpty()){
                    endMataSplit(row, splitType);
                    splitType = SplitType.NONE;
                } else if (splitType == SplitType.NONE){
                    throw new FishLinkException ("Found unexpected \"" + columnA + "\" before " + Constants.CONSTANTS_DIVIDER);                    
                }
              } else {
                   endMataSplit(row, splitType);
                   splitType = SplitType.NONE;
            }
            row++;
        } while (true); //will return out when finished
    }

    private void findAndCheckMetaSplits() throws FishLinkException{
        lastDataColumn = FishLinkUtils.indexToAlpha(sheet.getColumns());
        findMetaSplits();
        if (categoryRow == -1) {
            throw new FishLinkException("Unable to find \"" + Constants.CATEGORY_LABEL + "\" in column A.");
        }
        if (fieldRow == -1) {
            throw new FishLinkException("Unable to find \"" + Constants.FIELD_LABEL + "\" in column A.");
        }
        if (idTypeRow == -1) {
            throw new FishLinkException("Unable to find \"" + Constants.ID_VALUE_LABEL + "\" in column A.");
        }
    }

    private static Cell getCell(Sheet metaData, int col, int row) throws FishLinkException{
        try{
            return metaData.getCell(col, row);
        } catch (NullPointerException ex) {
            throw new FishLinkException ("Unable to find cell " + col + "  " + row + " in sheet " + 
                    metaData.getSheetInfo(), ex);
        } catch (XLWrapException ex) {
            throw new FishLinkException ("Unable to find cell " + col + "  " + row + " in sheet " + 
                    metaData.getSheetInfo(), ex);
        } catch (XLWrapEOFException ex) {
            throw new FishLinkException ("Unable to find cell " + col + "  " + row + " in sheet " + 
                    metaData.getSheetInfo(), ex);
        }        
    }
    

   private String getZeroBasedCellValue (int col, int actualRow) throws FishLinkException{
        Cell cell = getCell(sheet, col, actualRow);
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

    String getCellValue (String column, int row) throws FishLinkException{
        int col = Utils.alphaToIndex(column);
        int actualRow = row - 1;
        return getZeroBasedCellValue (col, actualRow);
    }

}


