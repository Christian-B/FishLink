package org.freshwaterlife.fishlink.xlwrap;

import at.jku.xlwrap.common.Utils;
import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.map.expr.val.XLExprValue;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.freshwaterlife.fishlink.POI_Utils;

/**
 *
 * @author Christian
 */
public class AbstractSheet {

    protected Sheet metaSheet;

    protected int firstData;

    //Excell columns and Rows used here are Excell based and not 0 based as XLWrap uses internally
    protected int categoryRow = -1;
    protected int fieldRow = -1;
    protected int idTypeRow = -1;
    protected int externalSheetRow = -1;
    protected int ignoreZerosRow = -1;
    protected int firstConstant = -1;
    protected int lastConstant = -1;
    protected String lastDataColumn;

    public AbstractSheet (Workbook metaWorkbook, String sheetName) throws XLWrapMapException {
        try {
            metaSheet = metaWorkbook.getSheet(sheetName);
        } catch (XLWrapException ex) {
            throw new XLWrapMapException("Unable to find sheet " + sheetName, ex);
        }
        findAndCheckMetaSplits();    
    }

    public AbstractSheet (Sheet theSheet) throws XLWrapMapException {
        metaSheet = theSheet;
        findAndCheckMetaSplits();
    }

    private enum SplitType{
        NONE, CONSTANT, HEADER
    }

    private void endMataSplit(int row, SplitType splitType){
         switch (splitType){
            case NONE:
                return;
            case CONSTANT:
                lastConstant = row - 1;
        }
    }

    private void findMetaSplits() throws XLWrapMapException{
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
                } else if (columnA.equalsIgnoreCase(Constants.EXTERNAL_LABEL)){
                    ignoreZerosRow = row;
                } else if (columnA.contains("links")){
                    throw new XLWrapMapException ("Links not currently supproted");
                } else if (columnA.contains(Constants.CONSTANTS_DIVIDER)){
                    firstConstant = row + 1;
                    endMataSplit(row, splitType);
                    splitType = SplitType.CONSTANT;
                } else if (columnA.equalsIgnoreCase(Constants.HEADER_LABEL)) {
                    endMataSplit(row, splitType);
                    splitType = SplitType.HEADER;
                } else if (columnA.isEmpty()){
                    endMataSplit(row, splitType);
                    splitType = SplitType.NONE;
                } else if (splitType == SplitType.HEADER){
                    try{
                        firstData = Integer.parseInt(columnA);
                        return;
                    } catch (Exception e){
                        System.err.println("Expected a number after \"header\" but found " + columnA);
                    }
                } 
            } else {
                   endMataSplit(row, splitType);
                   splitType = SplitType.NONE;
            }
            row++;
        } while (true); //will return out when finished
    }

    private void findAndCheckMetaSplits() throws XLWrapMapException{
        lastDataColumn = POI_Utils.indexToAlpha(metaSheet.getColumns() -1);
        findMetaSplits();
        if (categoryRow == -1) {
            throw new XLWrapMapException("Unable to find \"" + Constants.CATEGORY_LABEL + "\" in column A.");
        }
        if (fieldRow == -1) {
            throw new XLWrapMapException("Unable to find \"" + Constants.FIELD_LABEL + "\" in column A.");
        }
        if (idTypeRow == -1) {
            throw new XLWrapMapException("Unable to find \"" + Constants.ID_VALUE_LABEL + "\" in column A.");
        }
    }

   private String getZeroBasedCellValue (int col, int actualRow) throws XLWrapMapException{
        Cell cell = POI_Utils.getCell(metaSheet, col, actualRow);
        XLExprValue<?> value;
        try {
            value = Utils.getXLExprValue(cell);
        } catch (XLWrapException ex) {
            throw new XLWrapMapException("Unable to get value from cell. ", ex);
        }
        if (value == null){
            return null;
        }
        //remove the quotes that get added and we don't want here.
        return value.toString().replace("\"","");
    }

    protected String getCellValue (String column, int row) throws XLWrapMapException{
        int col = Utils.alphaToIndex(column);
        int actualRow = row - 1;
        return getZeroBasedCellValue (col, actualRow);
    }

    protected String getMetaCellValueOnDataColumn (String dataColumn, int row) throws XLWrapMapException{
        int col = Utils.alphaToIndex(dataColumn) + 1;
        int actualRow = row - 1;
        return getZeroBasedCellValue (col, actualRow);
    }
}


