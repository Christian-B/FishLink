/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap;

import at.jku.xlwrap.common.Utils;
import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.map.expr.val.XLExprValue;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.BufferedWriter;
import java.io.IOException;
import uk.co.brenn.metadata.POI_Utils;

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
    //private int externalColumnRow = -1;
    protected int ignoreZerosRow = -1;
    protected int firstLink = -1;
    protected int lastLink = -1;
    protected int firstConstant = -1;
    protected int lastConstant = -1;
    protected String lastDataColumn;

    //, String mapFileName, String rdfFileName
    public AbstractSheet (Sheet theSheet)
            throws XLWrapException, XLWrapEOFException, XLWrapMapException {
        metaSheet = theSheet;
        findAndCheckMetaSplits();
    }

    private enum SplitType{
        NONE, LINKS, CONSTANT, HEADER
    }

    private void endMataSplit(int row, SplitType splitType){
        //ystem.out.println (splitType);
        switch (splitType){
            case NONE:
                return;
            case LINKS:
                lastLink = row -1;
            case CONSTANT:
                lastConstant = row - 1;
        }
    }

    private void findMetaSplits() throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        int row = 1;
        SplitType splitType = SplitType.NONE;
        do {
            String columnA = getCellValue("A",row);
            //ystem.out.println(row + columnA);
            //ystem.out.println(row + ": " + columnA);
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
                //} else if (columnALower.equals("external column")){
                //    externalColumnRow = row;
                } else if (columnALower.contains("links")){
                    throw new XLWrapMapException ("Links not currently supproted");
                    //firstLink = row + 1;
                    //endMataSplit(row, splitType);
                    //splitType = SplitType.LINKS;
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
        //ystem.out.println(lastDataColumn + " " + metaSheet.getColumns() + "  " + metaSheet.getSheetInfo());
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

    private String template(){
        return metaSheet.getName() + "template";
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
        //String sheetName = null; //null is first (0) sheet.
        int col = Utils.alphaToIndex(column);
        int actualRow = row - 1;
        return getZeroBasedCellValue (col, actualRow);
    }

    protected String getMetaCellValueOnDataColumn (String dataColumn, int row) throws XLWrapException, XLWrapEOFException{
        int col = Utils.alphaToIndex(dataColumn) + 1;
        int actualRow = row - 1;
        //ystem.out.println (dataColumn + col + "  " + actualRow);
        return getZeroBasedCellValue (col, actualRow);
    }
}


