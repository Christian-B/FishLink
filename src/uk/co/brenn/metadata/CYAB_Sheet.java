/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.metadata;

import java.util.Date;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;

/**
 *
 * @author Christian
 */
public class CYAB_Sheet {

    Sheet poiSheet;

    CYAB_Sheet(Sheet sheet) {
        poiSheet = sheet;
    }

    /**
     * 
     * @param row Excel based
     * @return
     */
    private Row getRow(int row){
        Row poiRow = poiSheet.getRow(row -1);
        if (poiRow == null){
            poiRow = poiSheet.createRow(row - 1);
        }
        return poiRow;
    }

    private Cell getCell(String column, int row){
       Row poiRow = getRow(row);
        int columnNumber = POI_Utils.alphaToIndex(column);
        Cell cell = poiRow.getCell(columnNumber);
        if (cell == null){
           cell = poiRow.createCell(columnNumber);
        }
        return cell;
    }

    private Cell getCellZeroBased(int column, int row){
        Row poiRow = poiSheet.getRow(row);
        if (poiRow == null){
            poiRow = poiSheet.createRow(row);
        }
        Cell cell = poiRow.getCell(column);
        if (cell == null){
           cell = poiRow.createCell(column);
        }
        return cell;
    }

    /**
     * Sets the value of the cell to the integer.
     * 
     * Creates the row and or cell if required.
     * 
     * @param row Using Excel counting
     * @param column Using Excel names
     * @param value Integer value
     */
    void setValue (String column, int row, int value){
        Cell cell = getCell(column, row);
        cell.setCellValue(value);
    }

    /**
     * Sets the value of the cell to the double.
     *
     * Creates the row and or cell if required.
     *
     * @param row Using Excel counting
     * @param column Using Excel names
     * @param value Integer value
     */
    void setValue (String column, int row, double value){
        Cell cell = getCell(column, row);
        cell.setCellValue(value);
    }

    /**
     * Sets the value of the cell to the String.
     *
     * Creates the row and or cell if required.
     *
     * @param row Zero based
     * @param column zero based
     * @param value Integer value
     */
    public void setValueZeroBased (int column, int row, String value){
        Cell cell = getCellZeroBased(column, row);
        cell.setCellValue(value);
    }

    /**
     * Sets the value of the cell to the String.
     *
     * Creates the row and or cell if required.
     *
     * @param row Using Excel counting
     * @param column Using Excel names
     * @param value Integer value
     */
    public void setValue (String column, int row, String value){
        Cell cell = getCell(column, row);
        cell.setCellValue(value);
    }

    /**
     * Sets the value of the cell to the boolean.
     *
     * Creates the row and or cell if required.
     *
     * @param row Using Excel counting
     * @param column Using Excel names
     * @param value Integer value
     */
    public void setValue (String column, int row, boolean value){
        Cell cell = getCell(column, row);
        cell.setCellValue(value);
    }

    /**
     * Sets the value of the cell to the Date.
     *
     * Creates the row and or cell if required.
     *
     * @param row Using Excel counting
     * @param column Using Excel names
     * @param value Integer value
     */
    public void setValue (String column, int row, Date value){
        Cell cell = getCell(column, row);
        cell.setCellValue(value);
    }

    public String findFreeColumn (){
       Row poiRow = getRow(1);
       int columnNumber = 0;
       Cell cell;
       do {
          //ystem.out.println(columnNumber);
          cell = poiRow.getCell(columnNumber);
          columnNumber++;
       } while (cell != null);
       if (columnNumber != 26){
           char first = (char)( columnNumber+ 64);
           return first + "";
       }
       return "none";
   }

    void addValidation (String column, int row, String rule, String popupTitle, String popupMessage,
            int errorStyle, String errorTitle, String errorMessage)
            throws JavaToExcelException {
        //ystem.out.println ("list at " + column + row);
        System.out.println(column + ": " + rule);
        int columnNumber = POI_Utils.alphaToIndex(column);
        CellRangeAddressList addressList = new CellRangeAddressList(row - 1, row -1, columnNumber,  columnNumber);
        DataValidationHelper dataValidationHelper = poiSheet.getDataValidationHelper();
        DataValidationConstraint constraint;
        if (rule.isEmpty()){
            constraint = dataValidationHelper.createCustomConstraint("true");
        } else {
            constraint = dataValidationHelper.createFormulaListConstraint(rule);
        }
        DataValidation dataValidation = dataValidationHelper.createValidation(constraint, addressList);
        dataValidation.setSuppressDropDownArrow(false);
        if (!popupTitle.isEmpty()){
            dataValidation.createPromptBox(popupTitle, popupMessage);
            dataValidation.setShowPromptBox(true);
        }
        dataValidation.setErrorStyle(errorStyle);
        if (errorTitle.isEmpty()){
            dataValidation.createErrorBox(errorTitle, errorMessage);
        }
        poiSheet.addValidationData(dataValidation);
    }

    /* Pass through methods */
    
    public void addValidationData(HSSFDataValidation dataValidation) {
        poiSheet.addValidationData(dataValidation);
    }

    String getSheetName(){
        return poiSheet.getSheetName();
    }

    void autoSizeColumn(int column) {
        poiSheet.autoSizeColumn(column);
    }

    int getLastRowNum(){
       return poiSheet.getLastRowNum();
    }

    int getLastCellNum(){
        int maxCell = 0;
        for (Iterator<Row> rit = poiSheet.rowIterator(); rit.hasNext(); ) {
            Row row = rit.next();
            if (row.getLastCellNum() > maxCell){
                maxCell = row.getLastCellNum();
            }
        }
        return maxCell;
    }

}
