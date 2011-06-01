/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.metadata;

import java.util.Iterator;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
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
     * Sets the value of the cell to the integer.
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

    void addListValidation (String column, int row, String list, String popupTitle, String popupMessage)
            throws JavaToExcelException {
        //ystem.out.println ("list at " + column + row);
        int columnNumber = POI_Utils.alphaToIndex(column);
        CellRangeAddressList addressList = new CellRangeAddressList(row - 1, row -1, columnNumber,  columnNumber);
        DataValidationConstraint constraint;
        DataValidation dataValidation;
        if (poiSheet instanceof HSSFSheet) {
            constraint = DVConstraint.createFormulaListConstraint(list);
            dataValidation = new HSSFDataValidation(addressList, constraint);
        } else {
            throw new JavaToExcelException("Unable to add contraint to sheet of type " + poiSheet.getClass());
        }
        dataValidation.setSuppressDropDownArrow(false);
        dataValidation.createPromptBox(popupTitle, popupMessage);
        dataValidation.setShowPromptBox(true);
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
