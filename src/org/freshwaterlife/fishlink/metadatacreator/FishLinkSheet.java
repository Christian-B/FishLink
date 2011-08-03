package org.freshwaterlife.fishlink.metadatacreator;

import java.util.Calendar;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.freshwaterlife.fishlink.FishLinkUtils;
import org.freshwaterlife.fishlink.FishLinkException;

/**
 * Wrapper for an Apache Poi SS Sheet.
 * 
 * Main use is to allow Excel style Letter Column references and row numbers starting with 1.
 * @author Christian
 */
public class FishLinkSheet {

    /**
     * Apache poi sheet which does the actual work
     */
    private Sheet poiSheet;

    /**
     * Wraps a Apache Poi sheet with support methods.
     * 
     * Lets Columns and Row be referred to Using Excel referrences.
     * @param sheet The Apache poi method
     */
    FishLinkSheet(Sheet sheet) {
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

    /**
     *
     * @param row Using Excel counting
     * @param column Using Excel names
     * @return 
     */
    private Cell getCell(String column, int row){
       Row poiRow = getRow(row);
        int columnNumber = FishLinkUtils.alphaToIndex(column);
        Cell cell = poiRow.getCell(columnNumber);
        if (cell == null){
           cell = poiRow.createCell(columnNumber);
        }
        return cell;
    }

    /**
     * 
     * @param row Zero based
     * @param column zero based
     * @return 
     */
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
    void setValueZeroBased (int column, int row, String value){
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
    void setValue (String column, int row, String value){
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
    void setValue (String column, int row, boolean value){
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
    void setValue (String column, int row, Date value){
        Cell cell = getCell(column, row);
        cell.setCellValue(value);
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(value);
        int hours = calendar.get(Calendar.HOUR_OF_DAY);
        int minutes = calendar.get(Calendar.MINUTE);
        Workbook wb = poiSheet.getWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        if (hours == 0 && minutes == 0){
            //TODO: check for seconds
            CellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
            cell.setCellStyle(dateStyle);
            CellStyle test = cell.getCellStyle();
        } else {
            //TODO: Check if just a time
            CellStyle dateAndTimeStyle = wb.createCellStyle();
            dateAndTimeStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
            cell.setCellStyle(dateAndTimeStyle);
        }
    }

    /**
     * Sets the value of the cell to the Date.
     *
     * Creates the row and or cell if required.
     *
     * @param row Using Excel counting
     * @param column Using Excel names
     * @param value Integer value
     * @param format String format of the datacell;
     */
    void setValue (String column, int row, Date value, String format){
        Cell cell = getCell(column, row);
        cell.setCellValue(value);
        Workbook wb = poiSheet.getWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(format));
        cell.setCellStyle(cellStyle);
    }

    /**
     * Sets the background of the cell to Aqua.
     *
     * Creates the row and or cell if required.
     *
     * @param row Using Excel counting
     * @param column Using Excel names
     */
    void setForegroundAqua (String column, int row){
        Cell cell = getCell(column, row);
        // Aqua background
        CellStyle style = poiSheet.getWorkbook().createCellStyle();
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        Font font = poiSheet.getWorkbook().createFont();
        font.setBoldweight(font.BOLDWEIGHT_BOLD);
        style.setFont(font);
        cell.setCellStyle(style);
    }

    /**
     * Adds data validation to a cell.
     * 
     * Replaces and reference to a replaceable Columns with the actual Column in the rule.
     * For Example if the Rule is "=indirect($B)" and the replace column is "B" and the actual colum is "R",
     * It Changes the Rule to "=indirect($R)".
     * <p> Then calls {@link #addValidation(java.lang.String, int, java.lang.String, java.lang.String, java.lang.String, 
     *     int, java.lang.String, java.lang.String)} which in turn calls {@link #addCheckedValidation(java.lang.String, 
     * int, java.lang.String, java.lang.String, java.lang.String, int, java.lang.String, java.lang.String) }
     * @param replaceColum Excel name (excluding $) of the Column used in the preset rule
     * @param column Using Excel names
     * @param row Using Excel counting
     * @param rule String representation of the Rule.
     * @param popupTitle The Text for the title of the popup
     * @param popupMessage The Text for the message of the popup
     * @param errorStyle The Style of the Error based on POI DataValidation.ErrorStyle
     * @param errorTitle The Text for the title of the error popup
     * @param errorMessage The Text for the message of the error popup
     * @throws FishLinkException Any Exceptions Caught and wrapped
     */
    void addValidation (String replaceColum, String column, int row, String rule, String popupTitle, String popupMessage,
            int errorStyle, String errorTitle, String errorMessage) throws FishLinkException {
        rule = rule.replaceAll("\\$"+replaceColum, "\\$" + column);
        //ystem.out.println(rule);
        addValidation (column, row, rule, popupTitle, popupMessage, errorStyle, errorTitle, errorMessage);
    }
    
    /**
     * Adds data validation to a cell.
     * 
     * Validates the lengths of the titles and input messages.
     * <p> Then calls {@link #addCheckedValidation(java.lang.String, 
     * int, java.lang.String, java.lang.String, java.lang.String, int, java.lang.String, java.lang.String) }
     * @param column Using Excel names
     * @param row Using Excel counting
     * @param rule String representation of the Rule.
     * @param popupTitle The Text for the title of the popup
     * @param popupMessage The Text for the message of the popup
     * @param errorStyle The Style of the Error based on POI DataValidation.ErrorStyle
     * @param errorTitle The Text for the title of the error popup
     * @param errorMessage The Text for the message of the error popup
     * @throws FishLinkException If any of the Titles or Messages are too long. 
     */
    private void addValidation (String column, int row, String rule, String popupTitle, String popupMessage,
            int errorStyle, String errorTitle, String errorMessage) throws FishLinkException {
        if (rule.length() > 255) {
            throw new FishLinkException("Validation rule can not be longer than 255 for " + column + row);
        }
        if (popupTitle.length() > 31) {
            System.err.println(popupTitle);
            System.err.println(popupTitle.length());
            throw new FishLinkException("Validation popupTitle can not be longer than 31 for " + column + row);
        }
        if (popupMessage.length() > 255) {
            System.err.println(popupMessage);
            System.err.println(popupMessage.length());
            throw new FishLinkException("Validation popupMessage can not be longer than 255 for " + column + row);
        }
        if (errorTitle.length() > 31) {
            System.err.println(errorTitle);
            System.err.println(errorTitle.length());
            throw new FishLinkException("Validation errorTitle can not be longer than 31 for " + column + row);
        }
        if (errorMessage.length() > 255) {
            System.err.println(errorMessage);
            System.err.println(errorMessage.length());
            throw new FishLinkException("Validation errorMessage can not be longer than 255 for " + column + row);
        }
        addCheckedValidation (column, row, rule, popupTitle, popupMessage, errorStyle, errorTitle, errorMessage);
    }
    
    /**
     * Adds data validation to a cell.
     * 
     * Type of Data Validation used depends on the rule.
     * <ul>
     *    <li>An Empty rule can be used to Set a DataValidation which does no checking but does provide a popup.
     *    <li>Rules that start with '{' will be assumed to be an Explicit List, with the values comma separated 
     *        and an Optional closing '}'   
     *    <li>The Rule "BLANK" will create a DataValidation which forces the cell to be left Blank or empty.
     *    <li>Any other rule taken as a Formula List.
     * </ul>
     * @param column Using Excel names
     * @param row Using Excel counting
     * @param rule String representation of the Rule.
     * @param popupTitle The Text for the title of the popup
     * @param popupMessage The Text for the message of the popup
     * @param errorStyle The Style of the Error based on POI DataValidation.ErrorStyle
     * @param errorTitle The Text for the title of the error popup
     * @param errorMessage The Text for the message of the error popup
     */
    private void addCheckedValidation (String column, int row, String rule, String popupTitle, String popupMessage,
            int errorStyle, String errorTitle, String errorMessage) {
        int columnNumber = FishLinkUtils.alphaToIndex(column);
        CellRangeAddressList addressList = new CellRangeAddressList(row - 1, row -1, columnNumber,  columnNumber);
        DataValidationHelper dataValidationHelper = poiSheet.getDataValidationHelper();
        DataValidationConstraint constraint;
        if (rule.isEmpty()){
            if (popupTitle.isEmpty() && popupMessage.isEmpty()){
                return; //No rule or message so ignore.
            }
            //Create an any rule. Numberic method is only one that support Validation Type.
            constraint = dataValidationHelper.createNumericConstraint(DataValidationConstraint.ValidationType.ANY, 
                    DataValidationConstraint.OperatorType.IGNORED, null, null); 
        } else if (rule.startsWith("{")){
            rule = rule.replace("{", "");
            rule = rule.replace("}", "");
            String[] values = rule.split(",");
            constraint = dataValidationHelper.createExplicitListConstraint(values);
        } else if (rule.equals("BLANK")){
            constraint = dataValidationHelper.createTextLengthConstraint(
                    DataValidationConstraint.OperatorType.LESS_THAN, "1", "1");
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
        if (!errorTitle.isEmpty()){
            dataValidation.createErrorBox(errorTitle, errorMessage);
        }
        poiSheet.addValidationData(dataValidation);
    }

    /**
     * Sets the background of the cell to Aqua.
     *
     * Creates the row and or cell if required.
     *
     * @param row Using Excel counting
     * @param column Using Excel names
     */
    void createFreezePane (String column, int row){
        int columnNumber = FishLinkUtils.alphaToIndex(column);
        poiSheet.createFreezePane(columnNumber, row, columnNumber, row);
    }

    /* Pass through methods */
    
    /**
     * Sets the Column Width to the maximum Width of any cell in that Column.
     * @param column Zero Based Column Integer
     */
    void autoSizeColumn(int column) {
        poiSheet.autoSizeColumn(column);
    }

}
