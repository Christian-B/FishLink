package org.freshwaterlife.fishlink.metadatacreator;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.TypeAnnotation;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.File;
import java.util.Date;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Name;
import org.freshwaterlife.fishlink.FishLinkPaths;
import org.freshwaterlife.fishlink.FishLinkUtils;
import org.freshwaterlife.fishlink.FishLinkConstants;
import org.freshwaterlife.fishlink.FishLinkException;
import org.freshwaterlife.fishlink.xlwrap.expr.func.FishLinkToXlWrapRegister;

/**
 * Main class for copy data into a Specially prepared Annotation Workbook.
 * 
 * @author Christian
 */
public class MetaDataCreator {
       
    /**
     * Apache poi class for a=making sure there is ever only on Object for each Workbook and each sheet.
     */
    private ExecutionContext context;

    /**
     * Wrapper around the List Sheet from the MetaMaster
     */
    private Sheet masterListSheet;
    
    /**
     * Wrapper around the dropDowns sheet from the MetaMaster
     */
    private Sheet masterDropdownSheet;
    
    /**
     * Constructs a new MetaDataCreator class.
     * 
     * Obtains and wraps the sheets from the MetaMaster.
     * 
     * Creates an XLWrap ExecutionContext
     */
    public MetaDataCreator(String masterUrl) throws FishLinkException{
        context = new ExecutionContext();
        try {
            masterListSheet = context.getSheet(masterUrl, FishLinkConstants.LIST_SHEET);
        } catch (XLWrapException ex) {
            throw new FishLinkException("Error opening the vocabulary sheet " + FishLinkConstants.LIST_SHEET + 
                    " in ExcelSheet " + masterUrl, ex);
        }
        try {
            masterDropdownSheet = context.getSheet(masterUrl, FishLinkConstants.DROP_DOWN_SHEET);
        } catch (XLWrapException ex) {
            throw new FishLinkException("Error opening the dropdown sheet " + FishLinkConstants.DROP_DOWN_SHEET+ 
                    " in ExcelSheet " + masterUrl, ex);
        }
    }
        
    /**
     * Creates a Range found in the MetaMaster in the Annotatable Worksheet, and creates a named range on this data.
     * 
     * Works by reading one Column in the List sheet of the MetaMaster.
     * The top row ("1" by Excel counting 0 by Poi's) is assumed to contain the range name.
     * All the following cells directly below this (same Column) are assumed to be values in the range.
     * The first blank cell is assumed to be one below the end of the range.
     * 
     * @param annotationWorkbook Wrapped annotationWorkbook being created.
     * @param zeroColumn  column zero based index

     * @return
     * @throws FishLinkException 
     */
    private String createNamedRange (FishLinkWorkbook annotationWorkbook, int zeroColumn) throws FishLinkException {
        FishLinkSheet metaSheet = annotationWorkbook.getSheet(FishLinkConstants.LIST_SHEET);
        String rangeName =  FishLinkUtils.getTextZeroBased(masterListSheet, zeroColumn, 0);
        if (rangeName == null || rangeName.isEmpty()){
            return "";
        }
        String columnName = FishLinkUtils.indexToAlpha(zeroColumn);
        metaSheet.setValueZeroBased(zeroColumn, 0, rangeName);
        int zeroRow = 1;
        String fieldName = FishLinkUtils.getTextZeroBased(masterListSheet, zeroColumn, zeroRow);
        while (!fieldName.isEmpty()) {
            metaSheet.setValueZeroBased(zeroColumn, zeroRow, fieldName);
            zeroRow++;
            fieldName = FishLinkUtils.getTextZeroBased(masterListSheet, zeroColumn, zeroRow);
        } while (!fieldName.isEmpty());
        Name range = annotationWorkbook.createName();
        range.setNameName(rangeName);
        String rangeDef = "'" + FishLinkConstants.LIST_SHEET + "'!$" + columnName + "$2:$" + columnName + "$" + zeroRow;
        range.setRefersToFormula(rangeDef);
        return rangeName;
    }

    /**
     * Copies all the ranges from the MetaMaster List sheet and creates named ranges
     * 
     * The ranges are assumed to be on the List sheet,
     * Starting at Column A (or 0) for the categories and their fields.
     * Then a blank column.
     * Then the Subcategories and another balank column.
     * Finally the other ranges.
     * @param annotationWorkbook Wrapped annotationWorkbook being created.
     * @throws FishLinkException 
     */
    private void createNamedRanges (FishLinkWorkbook annotationWorkbook) throws FishLinkException {
        FishLinkSheet metaSheet = annotationWorkbook.getSheet(FishLinkConstants.LIST_SHEET);
        //Get the categories
        int zeroColumn = 0;
        String rangeName =  createNamedRange(annotationWorkbook, zeroColumn);
        while (!rangeName.isEmpty()) {
            zeroColumn++;
            rangeName = createNamedRange(annotationWorkbook, zeroColumn);
        }
        //first space splits the categories from the rest
        Name range = annotationWorkbook.createName();
        range.setNameName("Category");
        String columnName = FishLinkUtils.indexToAlpha(zeroColumn -1);
        range.setRefersToFormula("'" + FishLinkConstants.LIST_SHEET + "'!$A1:$" + columnName + "$1");
        //Now get the subTypes
        do {
            zeroColumn++;
            rangeName = createNamedRange(annotationWorkbook, zeroColumn);
        } while (!rangeName.isEmpty());
        //Now the renaming ranges
        do {
            zeroColumn++;
            rangeName = createNamedRange(annotationWorkbook, zeroColumn);
        } while (!rangeName.isEmpty());
        //Hack to avoid last column getting lost
        metaSheet.setValueZeroBased(zeroColumn + 1, 0, "");
    }

    /**
     * Copies the text in Column A from the MeteMaster dropdown sheet to the Annotation Sheet
     * @param annotationSheet Wrapped Sheet in the annotation Workbook being created.
     * @return The First Header Row in the Annotated sheet.
     * @throws FishLinkException 
     */
    private int prepareColumnA(FishLinkSheet annotationSheet) throws FishLinkException {
        int zeroRow = 0;
        String value;
        do {
            value = FishLinkUtils.getTextZeroBased(masterListSheet, 0, zeroRow);
            annotationSheet.setValue("A",zeroRow + 1, value);
            zeroRow++;
        } while (!value.isEmpty());
        annotationSheet.autoSizeColumn(0);
        zeroRow++; //leave a blank row
        annotationSheet.setValue("A",zeroRow, "Header");
        return zeroRow;
    }

    /**
     * Prepares the dropdown boxed in the Annotation rows for this Column.
     * 
     * @param lastMetaRow Last row which requires a drop down.
     * @param annotationSheet Wrapped Sheet in the annotation Workbook being created.
     * @param zeroColumn  column zero based index
     * @throws FishLinkException 
     */
    private void prepareDropDowns(FishLinkSheet annotationSheet, int lastMetaRow, String zeroColumn) 
            throws FishLinkException {
        for (int zeroRow = 0; zeroRow < lastMetaRow; zeroRow ++) {
            String list = FishLinkUtils.getTextZeroBased(masterDropdownSheet, 2, zeroRow);
            String popupTitle = FishLinkUtils.getTextZeroBased(masterDropdownSheet, 3, zeroRow);
            String popupMessage =  FishLinkUtils.getTextZeroBased(masterDropdownSheet, 4, zeroRow);
            String errorStyleString =  FishLinkUtils.getTextZeroBased(masterDropdownSheet, 5, zeroRow);
            int errorStyle = DataValidation.ErrorStyle.STOP;
            if (errorStyleString.startsWith("W")){
                errorStyle = DataValidation.ErrorStyle.WARNING;
            }
            if (errorStyleString.startsWith("I")){
                errorStyle = DataValidation.ErrorStyle.INFO;
            }
            String errorTitle =  FishLinkUtils.getTextZeroBased(masterDropdownSheet, 6, zeroRow);
            String errorMessage =  FishLinkUtils.getTextZeroBased(masterDropdownSheet, 7, zeroRow);
            annotationSheet.addValidation("B", zeroColumn, zeroRow + 1, list, popupTitle, popupMessage,
                    errorStyle, errorTitle, errorMessage);
        }
    }
    
    /**
     * Copies already filled in Annotations, from a previous annotation sheet into this one.
     * 
     * For each Annotation row with the same text in column A, 
     *    the value found in the copy sheet is copied to the new Annotation sheet.
     * No checking is done to see if the value is still valid,
     *     as values no longer valid may help determine the correct new value.
     * 
     * @param annotationSheet Wrapped Sheet in the annotation Workbook being created.
     * @param copySheet Wrapped Sheet in the Workbook being copied. 
     *                  Similar to dataSheet in other methods except that here is known to have old annotation rows.
     * @param lastMetaRow Last row in the annotationSheet the contains dropdowns.
     * @throws FishLinkException 
     */
    private void copyPreviousAnnotations(FishLinkSheet annotationSheet, Sheet copySheet, int lastMetaRow) 
            throws FishLinkException{
         for (int zeroRow = 0; zeroRow < lastMetaRow; zeroRow ++) {
            String masterColumnA = FishLinkUtils.getTextZeroBased(masterDropdownSheet, 0, zeroRow);
            int copyRow = -1;
            for (int tryRow = 0; tryRow <= lastMetaRow; tryRow ++) {
                 String tryColumnA = FishLinkUtils.getTextZeroBased(copySheet, 0, tryRow);
                 if (tryColumnA.equalsIgnoreCase(masterColumnA)){
                     copyRow = tryRow;
                 }
            }
            if (copyRow >= 0){
                for ( int zeroColumn = 0;  zeroColumn <= copySheet.getColumns(); zeroColumn++){
                    String copyMeta = FishLinkUtils.getTextZeroBased(copySheet, zeroColumn, copyRow);
                    annotationSheet.setValueZeroBased(zeroColumn, zeroRow, copyMeta);
                }
            }
        }
    }
    
    /**
     * Copies the data for one column in one sheet into the annotationSheet.
     * 
     * Copies all the data from one column in a (data)sheet in the input data into anAnnotationSheet in the workbook being created.
     * Does not copy any data in Annotation rows. This is copied by another method.
     * 
     * @param annotationSheet Wrapped Sheet in the annotation Workbook being created.
     * @param dataSheet Wrapper around the sheet the data comes from
     * @param headerRow Row in the annotation Sheet that contains the first header line. 
     *    (Not including Annotation/Dropdown) rows
     * @param annotationColumn Column being written to in Excel Indexing
     * @param zeroDataColumn Column being read from in zero based indexing
     * @param ignoreRows Number of rows at the top of the dataSheet that should be ignored.
     *    These are ignored because they contain annotations and not data. 
     * @throws FishLinkException 
     */
    private void copyData(FishLinkSheet annotationSheet, Sheet dataSheet, int headerRow, String annotationColumn, 
            int zeroDataColumn, int ignoreRows) throws FishLinkException {
        for (int zeroRow = 0; zeroRow < dataSheet.getRows() - ignoreRows; zeroRow++){
            Cell cell;
            try {
                cell = dataSheet.getCell(zeroDataColumn, zeroRow + ignoreRows);
            } catch (XLWrapException ex) {
                throw new FishLinkException("Error getting Cell", ex);
            } catch (XLWrapEOFException ex) {
                throw new FishLinkException("Error getting Cell", ex);
            }
            TypeAnnotation typeAnnotation;
            try {
                typeAnnotation = cell.getType();
            } catch (XLWrapException ex) {
                throw new FishLinkException("Error getting annotation type", ex);
            }
            switch (typeAnnotation){
                case BOOLEAN:
                    boolean booleanValue;
                    try {
                        booleanValue = cell.getBoolean();
                    } catch (XLWrapException ex) {
                        throw new FishLinkException("Error getting boolean value", ex);
                    }
                    annotationSheet.setValue(annotationColumn, headerRow + zeroRow, booleanValue);
                    break;
                case NUMBER:
                    double doubleValue;
                    try {
                        doubleValue = cell.getNumber();
                    } catch (XLWrapException ex) {
                        throw new FishLinkException("Error getting double value", ex);
                    }
                    annotationSheet.setValue(annotationColumn, headerRow + zeroRow, doubleValue);
                    break;
                case TEXT:
                    String textValue;
                    try {
                        textValue = cell.getText();
                    } catch (XLWrapException ex) {
                        throw new FishLinkException("Error getting text value", ex);
                    }
                    annotationSheet.setValue(annotationColumn, headerRow + zeroRow, textValue);
                    break;
                case DATE:
                    Date dateValue;
                    try {
                        dateValue = cell.getDate();
                    } catch (XLWrapException ex) {
                        throw new FishLinkException("Error getting date text value", ex);
                    }
                    String format;
                    try {
                        format = cell.getDateFormat();
                    } catch (XLWrapException ex) {
                        throw new FishLinkException("Error getting date format", ex);
                    }
                    annotationSheet.setValue(annotationColumn, headerRow + zeroRow, dateValue, format);
                    break;
                case NULL:
                    break;
                default:
                    throw new FishLinkException("Unexpected Cell Type");
            }
        }
        annotationSheet.setForegroundAqua(annotationColumn, headerRow);
        annotationSheet.autoSizeColumn(zeroDataColumn);
    }
    
    /**
     * Determines the number of annotation rows if any in the data being copied in.
     * 
     * This method check if column A has the expected annotation labels in the first 2 rows.
     *    These currently are "Category" and "Field"
     * If these are not found this method returns zero indicating this is raw(unannotated data)
     * <p>
     * It the first few rows indicate that the data Sheet being copied already has annotation rows 
     *    the number of such rows is returned.
     * The assumption is that the first blank cell in Column "A" indicates the end of the Annotation Rows.
     * @param dataSheet Wrapper around the sheet the data comes from
     * @return the number of Annotation rows or zero is not an Annotation sheet.
     * @throws FishLinkException 
     */
    private int findNumberOfAnnotationRows(Sheet dataSheet) throws FishLinkException{
        String columnA = FishLinkUtils.getTextZeroBased(dataSheet, 0, 0);
        if (!columnA.equalsIgnoreCase(FishLinkConstants.CATEGORY_LABEL)){
            return 0;
        }
        columnA = FishLinkUtils.getTextZeroBased(dataSheet, 0, 1);
        if (!columnA.equalsIgnoreCase(FishLinkConstants.FIELD_LABEL)){
            return 0;
        }
        //Ok assume it is old metaData
        int zeroRow = 2;
        String value;
        do {
            value = FishLinkUtils.getTextZeroBased(dataSheet, 0, zeroRow);
            zeroRow++;
        } while (!value.isEmpty());
        return zeroRow;  
    }

    /**
     * Adds dropdowns and copies data into one sheet of the workbook being created.
     * 
     * This method (via sub methods) does a number of things.
     * <ul>
     *     <li> Writes the labels in ColumnA
     *     <li> It determine if the data sheet has previous annotation rows.
     *         <ul> <li> If so it copied them. </ul>
     *     <li>It create Freeze Panes to keep the annotations at the top.
     *     <li>It adds the dropdowns to the annotation rows.
     *     <li>It copies the data into the annotation sheet.
     * </ul>
     * 
     * @param annotationSheet Wrapped Sheet in the annotation Workbook being created.
     * @param dataSheet Wrapper around the sheet the data comes from
     * @throws FishLinkException 
     */
    private void prepareSheet(FishLinkSheet annotationSheet, Sheet dataSheet) throws FishLinkException {
        int ignoreRows = findNumberOfAnnotationRows(dataSheet);
        int columnShift;
        int headerRow = prepareColumnA(annotationSheet);
        int lastColumn = dataSheet.getColumns();
        if (ignoreRows == 0){
            //Plus one as metaColumn one over from DetaColumn
            columnShift = 1;
        } else {
            //dataSheet is a annotationSheet so already has a none data column A
            columnShift = 0;
            copyPreviousAnnotations(annotationSheet, dataSheet, headerRow -2); 
        }
        annotationSheet.createFreezePane("B", headerRow);
        for ( int zeroDataColumn = 0;  zeroDataColumn < lastColumn; zeroDataColumn++){
            String metaColumn = FishLinkUtils.indexToAlpha(zeroDataColumn + columnShift); 
            prepareDropDowns(annotationSheet, headerRow -2, metaColumn);
            copyData(annotationSheet, dataSheet, headerRow, metaColumn, zeroDataColumn, ignoreRows);
        }
        //Hack to avoid last column getting lost
        annotationSheet.setValueZeroBased(lastColumn + 1, headerRow, "");
    }

    /**
     * Check to see if a sheet is empty.
     * @param dataSheet Wrapper around the sheet the data comes from
     * @return True if and only if the sheet has noCell beyond A1, otherwise False
     */
    private boolean containsData (Sheet dataSheet){
        if (dataSheet.getColumns() < 1) return false;
        if (dataSheet.getRows() < 1) return false;
        return true;
    }

    /**
     * Prepares all the data sheets in the annotation workbook based on the annotationWorkbook.
     * 
     * Ignores the List/Vocuabulary sheet as that is prepared elsewhere
     *     ignores any blank sheets.
     * For details see 
     * {@link #prepareSheet(org.freshwaterlife.fishlink.metadatacreator.FishLinkSheet, at.jku.xlwrap.spreadsheet.Sheet) }
     * @param annotationWorkbook Wrapped annotationWorkbook being created.
     * @param annotationWorkbook Wrapper around the data being copied in.
     * @throws FishLinkException 
     */
    private void prepareSheets(FishLinkWorkbook annotationWorkbook, Workbook dataWorkbook) 
            throws FishLinkException{
        String[] dataSheets = dataWorkbook.getSheetNames();
        for (int i = 0; i  < dataSheets.length; i++){
            if (!dataSheets[i].equalsIgnoreCase(FishLinkConstants.LIST_SHEET)){
                Sheet  dataSheet;
                try {
                    dataSheet = dataWorkbook.getSheet(dataSheets[i]);
                } catch (XLWrapException ex) {
                    throw new FishLinkException("Unable to get sheet " + dataSheets[i], ex);
                }
                if (containsData(dataSheet)){
                    FishLinkSheet metaSheet = annotationWorkbook.getSheet(dataSheets[i]);
                    prepareSheet(metaSheet, dataSheet);
                } else  {
                    FishLinkUtils.report("Skipping empty " + dataSheet.getSheetInfo());
                }
            }
        }
    }

    /**
     * Writes a new annotation workbook to the file based on the data at the url.
     * 
     * @param dataUrl XLWrap format URL to a data (or previous annotation sheet)
     * @param output File to be written to.
     * @throws FishLinkException 
     */
    private void writeAnnotationWorkbook(String dataUrl, File output) throws FishLinkException{
        FishLinkUtils.report("Creating metaData sheet for " + dataUrl);
        Workbook annotationWorkbook;
        try {
            annotationWorkbook = context.getWorkbook(dataUrl);
        } catch (XLWrapException ex) {
            throw new FishLinkException("Error opening the dataset " + dataUrl, ex);
        }
        FishLinkWorkbook metaWorkbook = new FishLinkWorkbook();
        createNamedRanges(metaWorkbook);
        //TODO copy
        prepareSheets(metaWorkbook, annotationWorkbook);
        metaWorkbook.write(output);
        FishLinkUtils.report("Wrote to  " + output.getAbsolutePath());
    }

    /**
     * Creates a annotated workbook based on the dataFile and the masterFile.
     * 
     * Opens an XLWrap Workbook based on dataUrl, and an XLWrap Workbook based on masterUrl.
     * Creates a new Excel workbook which contains the vocabulary lists from the master and all the sheets from data,
     *     annotated with the dropDowns found in the Master.
     * <p>Both the data pointed to by dataUrl and the masterFile pointed to by masterUrl 
     *     must be in a format XLWrap can handle.
     * <p>File formats XLWrap can handle include
     * <ul>
     *     <li>Excel xls files
     *     <li>Excel 2001 xlsx files (Assuming the FishLink extended XlWrap is used.
     *     <li>csv Files
     *     <li>Open Office files. (untested)
     * </ul>
     * <p>Know protocols it can handle included
     * <ul>
     *     <li>File: 
     *     <li>http:
     * </ul>
     * <p>If the data already contains annotated header rows, this tool will copy over the notations are much as possible.
     *     Even if they are no longer correct. The vocabulary lists and dropdown rules are taken from the new MetaMaster.
     * @param dataUrl Url to the raw data or a previous annotated data workbook to be updated. 
     * @param masterUrl URl to the MetaMaster file
     * @return File Where the annotated sheet is saved.
     * @throws FishLinkException Any Exceptions that might have occurred. Including that where thrown by XLWrap. 
     */
    public File createMetaData(String dataUrl) throws FishLinkException{
        String[] parts = dataUrl.split("[\\\\/.]");// Split on a forawrd slash, back slash and a full stop
        System.out.println("=======");
        String fileName;
        if (parts[parts.length-2].endsWith("MetaData")){
            fileName = parts[parts.length-2] + ".xls";
        } else {
            fileName = parts[parts.length-2] + "MetaData.xls";
        }
        File output =  new File (FishLinkPaths.META_FILE_ROOT, fileName);
        writeAnnotationWorkbook(dataUrl, output);
        return output;
    }
   
    /**
     * Explains what the main args are expected to be.
     */
    private static void usage(){
        FishLinkUtils.report("Creates a Meta Data collection file based on input.");
        FishLinkUtils.report("Requires two paramters.");
        FishLinkUtils.report("First is the raw data to be described.");
        FishLinkUtils.report("Second if the location (path and file name) of the MetaData file.");
    }

    /**
     * Main method for creating MetaData.
     * 
     * Converts the first argument into DataUrl, the second into MasterUrl and 
     * then calls {@link #createMetaData(java.lang.String, java.lang.String)}
     * @param args DataUrl and MasterUrl
     * @throws FishLinkException 
     */
    public static void main(String[] args) throws FishLinkException{
        if (args.length != 2){
            usage();
            System.exit(2);
        }
        FishLinkToXlWrapRegister.register();
        MetaDataCreator metaDataCreator = new MetaDataCreator(args[1]);
        File output = metaDataCreator.createMetaData(args[0]);
        FishLinkUtils.report("Mapping file can be found at: "+output.getAbsolutePath());
    }

}
