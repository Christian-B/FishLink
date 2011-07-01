package org.freshwaterlife.fishlink.metadatacreator;

import at.jku.xlwrap.common.XLWrapException;
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
import org.freshwaterlife.fishlink.MasterFactory;
import org.freshwaterlife.fishlink.POI_Utils;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

/**
 *
 * @author Christian
 */
public class MetaDataCreator {

    //Limits the number of rows copied in
    static private int MAX_DATA_ROW = Integer.MAX_VALUE;
    
//    static private int CATEGORY_ROW = 1;
 //   static private int FIELD_ROW = 2;
   // private Sheet Sheet;

    public MetaDataCreator(){
    }
    
    private void addMetaDataSheet(CYAB_Workbook dataWorkbook, String dataFile, String doi){
        CYAB_Sheet sheet = dataWorkbook.getSheet("MetaData");
        sheet.setValue("A", 1, "File");
        sheet.setValue("B", 1, dataFile);
        sheet.setValue("A", 2, "Doi");
        sheet.setValue("B", 2, doi);
    }

    private void writeMeta(CYAB_Workbook metaWorkbook, String dataName) throws XLWrapMapException {
        String fileFront = dataName.substring(0, dataName.lastIndexOf("."));
        metaWorkbook.write(FishLinkPaths.META_DIR + fileFront + "MetaData.xls");
    }

    private String createNamedRange (CYAB_Workbook metaWorkbook, int zeroColumn) throws XLWrapMapException {
        Sheet masterSheet = MasterFactory.getMasterListSheet();
        CYAB_Sheet metaSheet = metaWorkbook.getSheet(MasterFactory.LIST_SHEET);
        String rangeName =  MasterFactory.getTextZeroBased(masterSheet, zeroColumn, 0);
        if (rangeName == null || rangeName.isEmpty()){
            return "";
        }
        String columnName = POI_Utils.indexToAlpha(zeroColumn);
        metaSheet.setValueZeroBased(zeroColumn, 0, rangeName);
        int zeroRow = 1;
        String fieldName = MasterFactory.getTextZeroBased(masterSheet, zeroColumn, zeroRow);
        while (!fieldName.isEmpty()) {
            metaSheet.setValueZeroBased(zeroColumn, zeroRow, fieldName);
            zeroRow++;
            fieldName = MasterFactory.getTextZeroBased(masterSheet, zeroColumn, zeroRow);
        } while (!fieldName.isEmpty());
        Name range = metaWorkbook.createName();
        range.setNameName(rangeName);
        String rangeDef = "'" + MasterFactory.LIST_SHEET + "'!$" + columnName + "$2:$" + columnName + "$" + zeroRow;
        range.setRefersToFormula(rangeDef);
        return rangeName;
    }

    private void createNamedRanges (CYAB_Workbook metaWorkbook) throws XLWrapMapException {
        Sheet masterSheet =  MasterFactory.getMasterListSheet();
        CYAB_Sheet metaSheet = metaWorkbook.getSheet(MasterFactory.LIST_SHEET);
        //Get the categories
        int zeroColumn = 0;
        String rangeName =  createNamedRange(metaWorkbook, zeroColumn);
        while (!rangeName.isEmpty()) {
            zeroColumn++;
            rangeName = createNamedRange(metaWorkbook, zeroColumn);
        }
        //first space splits the categories from the rest
        Name range = metaWorkbook.createName();
        range.setNameName("Category");
        String columnName = POI_Utils.indexToAlpha(zeroColumn -1);
        range.setRefersToFormula("'" + MasterFactory.LIST_SHEET + "'!$A1:$" + columnName + "$1");
        //Now get the subTypes
        do {
            zeroColumn++;
            rangeName = createNamedRange(metaWorkbook, zeroColumn);
        } while (!rangeName.isEmpty());
        //Now the renaming ranges
        do {
            zeroColumn++;
            rangeName = createNamedRange(metaWorkbook, zeroColumn);
        } while (!rangeName.isEmpty());
    }

    private int prepareColumnA(Sheet masterSheet, CYAB_Sheet metaSheet, Sheet dataSheet) throws XLWrapMapException {
        int zeroRow = 0;
        String value;
        do {
            value = MasterFactory.getTextZeroBased(masterSheet, 0, zeroRow);
            metaSheet.setValue("A",zeroRow + 1, value);
            zeroRow++;
        } while (!value.isEmpty());
        metaSheet.autoSizeColumn(0);
        int maxRow = dataSheet.getRows();
        if (maxRow > 100){
            maxRow = 100;
        }
        metaSheet.setValue("A",zeroRow + 1, "Column in Data File");
        metaSheet.setValue("A",zeroRow + 2, "Header");
        for (int row = 2; row <= maxRow; row++){
           metaSheet.setValue("A",zeroRow + row + 1, row);
        }
        return zeroRow - 1;
    }

    private void prepareDropDowns(Sheet masterSheet, int lastMetaRow, CYAB_Sheet metaSheet, String column) 
            throws XLWrapMapException {
        for (int zeroRow = 0; zeroRow < lastMetaRow; zeroRow ++) {
            String list = MasterFactory.getTextZeroBased(masterSheet, 2, zeroRow);
            String popupTitle = MasterFactory.getTextZeroBased(masterSheet, 3, zeroRow);
            String popupMessage =  MasterFactory.getTextZeroBased(masterSheet, 4, zeroRow);
            String errorStyleString =  MasterFactory.getTextZeroBased(masterSheet, 5, zeroRow);
            int errorStyle = DataValidation.ErrorStyle.STOP;
            if (errorStyleString.startsWith("W")){
                errorStyle = DataValidation.ErrorStyle.WARNING;
            }
            if (errorStyleString.startsWith("I")){
                errorStyle = DataValidation.ErrorStyle.INFO;
            }
            String errorTitle =  MasterFactory.getTextZeroBased(masterSheet, 6, zeroRow);
            String errorMessage =  MasterFactory.getTextZeroBased(masterSheet, 7, zeroRow);
            metaSheet.addValidation("B", column, zeroRow + 1, list, popupTitle, popupMessage,
                    errorStyle, errorTitle, errorMessage);
        }
    }

    private void copyMetaData(Sheet masterSheet, CYAB_Sheet metaSheet, Sheet copySheet, int lastMetaRow, int lastColumn) 
            throws XLWrapMapException{
        System.out.println(copySheet.getSheetInfo());
        for (int zeroRow = 0; zeroRow < lastMetaRow; zeroRow ++) {
            String masterColumnA = MasterFactory.getTextZeroBased(masterSheet, 0, zeroRow);
            int copyRow = -1;
            for (int tryRow = 0; tryRow <= lastMetaRow; tryRow ++) {
                 String tryColumnA = MasterFactory.getTextZeroBased(copySheet, 0, tryRow);
                 if (tryColumnA.equalsIgnoreCase(masterColumnA)){
                     copyRow = tryRow;
                 }
            }
            System.out.println(masterColumnA + ": " + copyRow);
            if (copyRow >= 0){
                for ( int zeroColumn = 0;  zeroColumn <= lastColumn; zeroColumn++){
                    String copyMeta = MasterFactory.getTextZeroBased(copySheet, zeroColumn, copyRow);
                    metaSheet.setValueZeroBased(zeroColumn, zeroRow, copyMeta);
                }
            }
        }
    }
            
    private void copyData(int letterRow, CYAB_Sheet metaSheet, Sheet dataSheet, String metaColumn) throws XLWrapMapException {
        int zeroColumn = POI_Utils.alphaToIndex(metaColumn) - 1; //-1 as Column A of Data goes in B of Meta
        String dataColumn = POI_Utils.indexToAlpha(zeroColumn);
        metaSheet.setValue(metaColumn, letterRow, dataColumn);
        metaSheet.setForegroundAqua(metaColumn, letterRow);
        int maxRow = dataSheet.getRows();
        if (maxRow > MAX_DATA_ROW){
            maxRow = MAX_DATA_ROW;
        }
        for (int zeroRow = 0; zeroRow < maxRow; zeroRow++){
            Cell cell;
            try {
                cell = dataSheet.getCell(zeroColumn, zeroRow);
            } catch (XLWrapException ex) {
                throw new XLWrapMapException("Error getting Cell", ex);
            } catch (XLWrapEOFException ex) {
                throw new XLWrapMapException("Error getting Cell", ex);
            }
            TypeAnnotation typeAnnotation;
            try {
                typeAnnotation = cell.getType();
            } catch (XLWrapException ex) {
                throw new XLWrapMapException("Error getting annotation type", ex);
            }
            switch (typeAnnotation){
                case BOOLEAN:
                    boolean booleanValue;
                    try {
                        booleanValue = cell.getBoolean();
                    } catch (XLWrapException ex) {
                        throw new XLWrapMapException("Error getting boolean value", ex);
                    }
                    metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, booleanValue);
                    break;
                case NUMBER:
                    double doubleValue;
                    try {
                        doubleValue = cell.getNumber();
                    } catch (XLWrapException ex) {
                        throw new XLWrapMapException("Error getting double value", ex);
                    }
                    metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, doubleValue);
                    break;
                case TEXT:
                    String textValue;
                    try {
                        textValue = cell.getText();
                    } catch (XLWrapException ex) {
                        throw new XLWrapMapException("Error getting text value", ex);
                    }
                    metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, textValue);
                    break;
                case DATE:
                    Date dateValue;
                    try {
                        dateValue = cell.getDate();
                    } catch (XLWrapException ex) {
                        throw new XLWrapMapException("Error getting date text value", ex);
                    }
                    String format;
                    try {
                        format = cell.getDateFormat();
                    } catch (XLWrapException ex) {
                        throw new XLWrapMapException("Error getting date format", ex);
                    }
                    metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, dateValue, format);
                    break;
                case NULL:
                    break;
                default:
                    throw new XLWrapMapException("Unexpected Cell Type");
            }
        }
        metaSheet.autoSizeColumn(zeroColumn);
    }

    private void prepareSheet(Sheet masterSheet, CYAB_Sheet metaSheet, Sheet dataSheet, Sheet copySheet) throws XLWrapMapException {
        int lastMetaRow = prepareColumnA(masterSheet, metaSheet, dataSheet);
        int lastColumn = dataSheet.getColumns();
        metaSheet.createFreezePane("B", lastMetaRow + 2);
        for ( int zeroDataColumn = 0;  zeroDataColumn < lastColumn; zeroDataColumn++){
            String metaColumn = POI_Utils.indexToAlpha(zeroDataColumn + 1); //Plus one as metaColumn on over from DetaColumn
            prepareDropDowns(masterSheet, lastMetaRow, metaSheet, metaColumn);
            copyData(lastMetaRow + 2, metaSheet, dataSheet, metaColumn);
        }
        System.out.println(dataSheet.getSheetInfo());
        System.out.println(copySheet);
        if (copySheet != null){
            copyMetaData(masterSheet, metaSheet, copySheet, lastMetaRow + 2, lastColumn); 
        }
    }

    private boolean containsData (Sheet dataSheet){
        if (dataSheet.getColumns() < 1) return false;
        if (dataSheet.getRows() < 1) return false;
        return true;
    }

    private void prepareSheets(CYAB_Workbook metaWorkbook, Workbook dataWorkbook, Workbook copyWorkbook) throws XLWrapMapException{
        Sheet masterSheet = MasterFactory.getMasterDropdownSheet();
        String[] dataSheets = dataWorkbook.getSheetNames();
        for (int i = 0; i  < dataSheets.length; i++){
            Sheet  dataSheet;
            try {
                dataSheet = dataWorkbook.getSheet(dataSheets[i]);
            } catch (XLWrapException ex) {
                throw new XLWrapMapException("Unable to get sheet " + dataSheets[i], ex);
            }
            Sheet copySheet = null;
            if (copyWorkbook != null){
                try {
                    copySheet = copyWorkbook.getSheet(dataSheet.getName());
                } catch (XLWrapException ex) {
                    System.err.println(ex);
                    copySheet = null;
                }
            }
            if (containsData(dataSheet)){
                CYAB_Sheet metaSheet = metaWorkbook.getSheet(dataSheets[i]);
                prepareSheet(masterSheet, metaSheet, dataSheet, copySheet);
            } else  {
                System.out.println("Skipping empty " + dataSheet.getSheetInfo());
            }
        }
    }

    public void prepareMetaDataOnDoi(String dataFile, String doi) throws XLWrapMapException{
        System.out.println("Preparing meta data collector for " + dataFile);
        Workbook dataWorkbook;
        try {
            dataWorkbook = MasterFactory.getExecutionContext().getWorkbook("file:" + FishLinkPaths.RAW_DIR +dataFile);
        } catch (XLWrapException ex) {
            throw new XLWrapMapException("Unable to open file " + dataFile, ex);
        }
        Workbook copyWorkbook;
        try {
           String fileFront = dataFile.substring(0, dataFile.lastIndexOf("."));
           String copyName = FishLinkPaths.OLD_META_DIR + fileFront + "MetaData.xls";
           copyWorkbook = MasterFactory.getExecutionContext().getWorkbook("file:" + copyName);
        } catch (XLWrapException ex) {
            System.err.println(ex);
            System.out.println("No old data for " + FishLinkPaths.OLD_META_DIR + dataFile);
            copyWorkbook = null;
        }
        CYAB_Workbook metaWorkbook = new CYAB_Workbook();
        addMetaDataSheet(metaWorkbook, dataFile, doi);
        createNamedRanges(metaWorkbook);
        prepareSheets(metaWorkbook,dataWorkbook, copyWorkbook);
        writeMeta(metaWorkbook, dataFile);
    }

    public void prepareMetaDataOnTarget(String dataPath, String targetPath) throws XLWrapMapException{
        Workbook dataWorkbook;
        try {
            dataWorkbook = MasterFactory.getExecutionContext().getWorkbook(dataPath);
        } catch (Exception e){
            try {
                //assume the "file:" bit is missing
                dataWorkbook = MasterFactory.getExecutionContext().getWorkbook("file:" + dataPath);
            } catch (XLWrapException ex) {
                throw new XLWrapMapException("Unable to create file " + dataPath, ex);
            }
        }
        CYAB_Workbook metaWorkbook = new CYAB_Workbook();
        createNamedRanges(metaWorkbook);
        prepareSheets(metaWorkbook,dataWorkbook, null);
        metaWorkbook.write(targetPath);
    }

    public static void main(String[] args) throws XLWrapMapException{
        MetaDataCreator creator = new MetaDataCreator();

        //creator.prepareMetaDataOnDoi ("CumbriaTarnsPart1.xls", "CTP1");
        //creator.prepareMetaDataOnDoi ("FBA_Tarns.xls", "FBA345");
        //creator.prepareMetaDataOnDoi ("Records.xls", "rec12564");
        //creator.prepareMetaDataOnDoi ("Species.xls", "spec564");
        //creator.prepareMetaDataOnDoi ("Stokoe.xls", "stokoe32433232");
        //creator.prepareMetaDataOnDoi ("Tarns.xls", "tarns33exdw2");
        //creator.prepareMetaDataOnDoi ("TarnschemFinal.xls", "TSF1234");
        creator.prepareMetaDataOnDoi ("WillbyGroups.xls", "wbgROUPS8734");
    }
}
