/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.metadata;

import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *
 * @author Christian
 */
public class MetaDataCreator {

    static private String MAIN_ROOT = "c:Dropbox/FishLink XLWrap data/";

    private String metaRoot;
    private String dataRoot;
    
    static private int CATEGORY_ROW = 1;
    static private int FIELD_ROW = 2;

    public MetaDataCreator(String metaDir, String dataDir){
        metaRoot = metaDir;
        dataRoot = dataDir;
    }

    private void addMetaDataSheet(CYAB_Workbook dataWorkbook, String dataFile, String doi){
        CYAB_Sheet sheet = dataWorkbook.getSheet("MetaData");
        sheet.setValue("A", 1, "File");
        sheet.setValue("A", 2, dataFile);
        sheet.setValue("B", 1, "Doi");
        sheet.setValue("B", 2, doi);
    }

    private void writeMeta(CYAB_Workbook metaWorkbook, String dataName) throws FileNotFoundException, IOException{
        String fileFront = dataName.substring(0, dataName.lastIndexOf("."));
        metaWorkbook.write(metaRoot + fileFront + "MetaData.xls");
    }

    private void prepareColumnA(CYAB_Sheet metaSheet, CYAB_Sheet dataSheet){
        metaSheet.setValue("A", CATEGORY_ROW, "Category");
        metaSheet.setValue("A", FIELD_ROW, "Field");
        metaSheet.setValue("A", 3, "Id/Value Column (or \"Row\")");
        metaSheet.setValue("A", 4, "'-- Links --");
        metaSheet.setValue("A", 5, "Site");
        metaSheet.setValue("A", 6, "*Site Link Type*");
        metaSheet.setValue("A", 7, "Location");
        metaSheet.setValue("A", 8, "Survey");
        metaSheet.setValue("A", 9, "'-- Constants --");
        metaSheet.setValue("A", 10, "Factor");
        metaSheet.setValue("A", 11, "Unit");
        metaSheet.setValue("A", 12, "DerivationFactor");
        metaSheet.setValue("A", 13, "Date/Start Date");
        metaSheet.setValue("A", 14, "");
        metaSheet.setValue("A", 15, "End Date (if Applicable)");
        int headerRow = 17;
        metaSheet.setValue("A", headerRow, "Column in Data File ");
        metaSheet.autoSizeColumn(0);
        int maxRow = dataSheet.getLastRowNum();
        System.out.println(maxRow);
        if (maxRow > 100){
            maxRow = 100;
        }
        for (int row = 1; row <= maxRow; row++){
           metaSheet.setValue("A",headerRow + row, row);
        }
    }

    private void prepareDropDowns(CYAB_Sheet metaSheet, String column) throws JavaToExcelException{
        metaSheet.addListValidation(column, CATEGORY_ROW, ListWriter.CATEGORY_NAME, "Type of data",
                "Please select the catogory that data in this column belongs to");
        metaSheet.addListValidation(column, FIELD_ROW,  "INDIRECT(SUBSTITUTE($" + column + "$1,\" \",\"_\"))",
                "Type of feild", "Please select the feild that data in this column belongs to");
    }

    private void prepareSheet(CYAB_Sheet metaSheet, CYAB_Sheet dataSheet) throws JavaToExcelException{
        prepareColumnA(metaSheet, dataSheet);
        prepareDropDowns(metaSheet, "B");
    }

    private void prepareSheets(CYAB_Workbook metaWorkbook, CYAB_Workbook dataWorkbook) throws JavaToExcelException{
        for (int sheetNumber = 0; sheetNumber < dataWorkbook.getNumberOfSheets(); sheetNumber++){
            CYAB_Sheet dataSheet = dataWorkbook.getSheetAt(sheetNumber);
            CYAB_Sheet metaSheet = metaWorkbook.getSheet(dataSheet.getSheetName());
            prepareSheet(metaSheet, dataSheet);
        }
    }

    public void prepareMetaData(String dataFile, String doi) throws FileNotFoundException, IOException, InvalidFormatException, JavaToExcelException{
        CYAB_Workbook dataWorkbook = new CYAB_Workbook(dataRoot + dataFile);
        CYAB_Workbook metaWorkbook = new CYAB_Workbook();
        addMetaDataSheet(metaWorkbook, dataFile, doi);
        ListWriter.writeLists(metaWorkbook);
        prepareSheets(metaWorkbook,dataWorkbook);
        writeMeta(metaWorkbook, dataFile);
    }

    public static void main(String[] args) throws IOException, FileNotFoundException, InvalidFormatException, JavaToExcelException{
        MetaDataCreator creator = new MetaDataCreator(MAIN_ROOT + "Meta Data/",MAIN_ROOT + "Raw Data/");
        creator.prepareMetaData ("Records.xls", "rec12564");
    }
}
