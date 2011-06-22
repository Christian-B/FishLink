package org.freshwaterlife.fishlink.xlwrap;

import at.jku.xlwrap.common.Utils;
import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.map.expr.val.XLExprValue;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import org.freshwaterlife.fishlink.metadatacreator.POI_Utils;

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

    public AbstractSheet (Sheet theSheet)
            throws XLWrapException, XLWrapEOFException, XLWrapMapException {
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

    private void findMetaSplits() throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        int row = 1;
        SplitType splitType = SplitType.NONE;
        do {
            String columnA = getCellValue("A",row);
            if (columnA != null){
                String columnALower = columnA.toLowerCase();
                if (columnALower.equals("category")){
                    categoryRow = row;
                } else if (columnALower.equals("field")){
                    fieldRow = row;
                } else if (columnALower.startsWith("id/value column")){
                    idTypeRow = row;
                } else if (columnALower.equals("external sheet")){
                    externalSheetRow = row;
                } else if (columnALower.equals("ignore zeros")){
                    ignoreZerosRow = row;
                } else if (columnALower.contains("links")){
                    throw new XLWrapMapException ("Links not currently supproted");
                } else if (columnALower.contains("constant")){
                    firstConstant = row + 1;
                    endMataSplit(row, splitType);
                    splitType = SplitType.CONSTANT;
                } else if (columnALower.equals("header")) {
                    endMataSplit(row, splitType);
                    splitType = SplitType.HEADER;
                } else if (columnALower.isEmpty()){
                    endMataSplit(row, splitType);
                    splitType = SplitType.NONE;
                } else if (splitType == SplitType.HEADER){
                    try{
                        firstData = Integer.parseInt(columnALower);
                        return;
                    } catch (Exception e){
                        System.err.println("Expected a number after \"header\" but found " + columnALower);
                    }
                } 
            } else {
                   endMataSplit(row, splitType);
                   splitType = SplitType.NONE;
            }
            row++;
        } while (true); //will return out when finished
    }

    private void findAndCheckMetaSplits() throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        lastDataColumn = POI_Utils.indexToAlpha(metaSheet.getColumns() -1);
        findMetaSplits();
        if (categoryRow == -1) {
            throw new XLWrapException("Unable to find \"category\" in column A.");
        }
        if (fieldRow == -1) {
            throw new XLWrapException("Unable to find \"field\" in column A.");
        }
        if (idTypeRow == -1) {
            throw new XLWrapException("Unable to find \"Id/Value column\" in column A.");
        }
    }

   private String getZeroBasedCellValue (int col, int actualRow) throws XLWrapException, XLWrapEOFException{
        Cell cell = metaSheet.getCell(col, actualRow);
        XLExprValue<?> value = Utils.getXLExprValue(cell);
        if (value == null){
            return null;
        }
        //remove the quotes that get added and we don't want here.
        return value.toString().replace("\"","");
    }

    protected String getCellValue (String column, int row) throws XLWrapException, XLWrapEOFException{
        int col = Utils.alphaToIndex(column);
        int actualRow = row - 1;
        return getZeroBasedCellValue (col, actualRow);
    }

    protected String getMetaCellValueOnDataColumn (String dataColumn, int row) throws XLWrapException, XLWrapEOFException{
        int col = Utils.alphaToIndex(dataColumn) + 1;
        int actualRow = row - 1;
        return getZeroBasedCellValue (col, actualRow);
    }
}


