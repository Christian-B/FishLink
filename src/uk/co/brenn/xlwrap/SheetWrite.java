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
public class SheetWrite {

    private static String RDF_BASE_URL = "http://rdf.fba.org.uk/";

    private ExecutionContext context;
    private Sheet metaSheet;
    private String sheetInURI;
    private int sheetNumber;

    private String dataPath;
    private String doi;

    int firstData;

    //Excell columns and Rows used here are Excell based and not 0 based as XLWrap uses internally
    private int categoryRow = -1;
    private int fieldRow = -1;
    private int idTypeRow = -1;
    private int externalSheetRow = -1;
    //private int externalColumnRow = -1;
    private int ignoreZerosRow = -1;
    private int firstLink = -1;
    private int lastLink = -1;
    private int firstConstant = -1;
    private int lastConstant = -1;
    private String LAST_DATA_COLUMN;

    //, String mapFileName, String rdfFileName
    public SheetWrite (ExecutionContext context, Workbook metaWorkbook, String dataURL, String doi, String sheetName)
            throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        this.doi = doi;
        dataPath = dataURL;
        Workbook dataWorkbook = context.getWorkbook(dataPath);
        String[] sheetNames = dataWorkbook.getSheetNames();
        for (int i = 0; i< sheetNames.length; i++ ){
            if (sheetNames[i].equalsIgnoreCase(sheetName)){
                sheetNumber = i;
                Sheet dataSheet = dataWorkbook.getSheet(i);
                sheetInURI = dataSheet.getName() + "/";
                System.out.println("Successfully opened RawData: " + dataSheet.getName());
            }
        }
        if (sheetInURI == null){
            throw new XLWrapMapException ("Unable to find sheet " + sheetName);
        }
        String[] metaNames = metaWorkbook.getSheetNames();
        for (int i = 0; i< metaNames.length; i++ ){
            System.out.println("£" + metaNames[i]+"£");
            System.out.println(metaNames[i].equals(sheetName));
        }
        metaSheet = metaWorkbook.getSheet(sheetName);
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

    private void findMetaSplits() throws XLWrapException, XLWrapEOFException{
        int row = 1;
        SplitType splitType = SplitType.NONE;
        do {
            String columnA = getCellValue("A",row);
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
                //} else if (columnALower.equals("external column")){
                //    externalColumnRow = row;
                } else if (columnALower.contains("links")){
                    firstLink = row + 1;
                    endMataSplit(row, splitType);
                    splitType = SplitType.LINKS;
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

    private void findAndCheckMetaSplits() throws XLWrapException, XLWrapEOFException{
        LAST_DATA_COLUMN = POI_Utils.indexToAlpha(metaSheet.getColumns() -1);
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

    protected void writeMapping(BufferedWriter writer) throws IOException, XLWrapMapException{
        writer.write("	xl:offline \"false\"^^xsd:boolean ;");
        writer.newLine();
        writer.write("	xl:template [");
        writer.newLine();
        writer.write("		xl:fileName \"" + dataPath + "\" ;");
        writer.newLine();
        writer.write("		xl:sheetNumber \""+ sheetNumber +"\" ;");
        writer.newLine();
        writer.write("		xl:templateGraph :");
        writer.write(template());
        writer.write(" ;");
        writer.newLine();
        writer.write("		xl:transform [");
        writer.newLine();
        writer.write("			a rdf:Seq ;");
        writer.newLine();
        writer.write("			rdf:_1 [");
        writer.newLine();
        writer.write("				a xl:RowShift ;");
        writer.newLine();
        writer.write("				xl:restriction \"A" + firstData + ":" + LAST_DATA_COLUMN + firstData + "\" ;");
        writer.newLine();
        writer.write("				xl:steps \"1\" ;");
        writer.newLine();
        writer.write("			] ;");
        writer.newLine();
        writer.write("		]");
        writer.newLine();
        writer.write("	] ;");
        writer.newLine();
    }

    private String getCellValue (String column, int row) throws XLWrapException, XLWrapEOFException{
        //String sheetName = null; //null is first (0) sheet.
        int col = Utils.alphaToIndex(column);
        int actualRow = row - 1;
        //CellRange cellRange = new CellRange("File:" +xlsPath, sheetName, col, actualRow);
        //Cell cell = context.getCell(cellRange);
        Cell cell = metaSheet.getCell(col, actualRow);
        XLExprValue<?> value = Utils.getXLExprValue(cell);
        if (value == null){
            return null;
        }
        //remove the quotes that get added and we don't want here.
        return value.toString().replace("\"","");
    }

    private void writeURI(BufferedWriter writer, String externalUri, String category, String field, String idType, String dataColumn,
            boolean ignoreZeros) throws IOException, XLWrapMapException {
        String localUri = RDF_BASE_URL + category + "/" + doi + "/" + sheetInURI;
        System.out.println("writeURI " + externalUri + " " + category + " " + field + " " + idType + " " + dataColumn);
        if (field.toLowerCase().equals("id")){
            writer.write("[ xl:uri \"ID_URI('" + externalUri + "', " + dataColumn + firstData + ","
                    + ignoreZeros + ")\"^^xl:Expr ] ");
            return;
        }
        if(field.toLowerCase().equals("value")) {
            writer.write("[ xl:uri \"CELL_URI('" + localUri + "', " + dataColumn + firstData + ","
                    + ignoreZeros + ")\"^^xl:Expr ] ");
            return;
        }
        if ((idType == null) || (idType.isEmpty())){
            throw new XLWrapMapException("Column " + dataColumn + " in Sheet " + metaSheet.getName() +
                    " has an unexpected emtpy IDType");
        }
        if (idType.equalsIgnoreCase("n/a")) {
            throw new XLWrapMapException("Column " + dataColumn + " in Sheet " + metaSheet.getName() +
                    " has an unexpected n/a IDType with field " + field);
        }
        if(category.toLowerCase().equals("observation")) {
            writer.write("[ xl:uri \"OTHER_CELL_URI('" + localUri + "', " + idType + firstData + ","
                    + dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
            return;
        }
        if (idType.equalsIgnoreCase("ROW")){
            writer.write("[ xl:uri \"ROW_URI('" + localUri + "row', " + dataColumn + firstData + "," +
                    ignoreZeros + ")\"^^xl:Expr ] ");
        } else {
            writer.write("[ xl:uri \"OTHER_ID_URI('" + localUri + "', " + idType + firstData + "," +
                    dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
        }
    }

    private void writeLink (BufferedWriter writer, String metaColumn, String dataColumn,
            int row, boolean ignoreZeros)
            throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException{
        String feild = getCellValue ("A", row);
        if (feild == null){
            return;
        }
        if (feild.contains("*")){
            //Link from previous row.
            return;
        }
        String nextFeild = getCellValue ("A", row + 1);
        String type = feild;
        if (nextFeild.contains("*") && nextFeild.contains(feild)){
            type = getCellValue (metaColumn, row + 1);
            if (type == null){
                type = feild;
            }
        }
        String link = getCellValue (metaColumn, row);
        System.out.println("writeLink " + metaColumn + "\t" + feild + "\t" + link);
        if (link == null){
           //ystem.out.println("Skippig column " + metaColumn + " " + feild + "row as it is blank");
            return;
        }
        if (link.equalsIgnoreCase("n/a")){
            return;
        }
        writer.write("	vocab:has" + type + "\t");
        //TODO checkthis
        writeURI(writer, null, "link", feild, link, dataColumn, ignoreZeros);
        writer.write(" ;");
        writer.newLine();
    }

    private void writeConstant (BufferedWriter writer, String metaColumn, int row)
            throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException{
        String feild = getCellValue ("A", row);
        if (feild == null){
            return;
        }
        String value = getCellValue (metaColumn, row);
        //ystem.out.println(metaColumn + "\t" + feild + "\t" + value);
        if (value == null){
            //ystem.out.println("Skippig column " + metaColumn + " " + feild + "row as it is blank");
            return;
        }
        if (value.equalsIgnoreCase("n/a")){
            return;
        }
        if (value.contains("!")){
            if (value.startsWith("!")){
                value = value.substring(1) + firstData;
            } else {
                throw new XLWrapMapException("Unexpected \"!\" in cell "+ metaSheet.getName() + "!" + metaColumn + row);
            }
        } else {
            value = "'" + value + "'";
        }
        if (feild.contains("/")){
            String nextValue = getCellValue (metaColumn, row+1);
            if (nextValue == null){
                writer.write("	vocab:has" + feild.substring(0, feild.indexOf('/')) + "\t");
            } else {
                writer.write("	vocab:has" + feild.substring(feild.indexOf('/')+1) + "\t");
            }
        } else {
            writer.write("	vocab:has" + feild + "\t");
        }
        writer.write("[ xl:uri \"'" + RDF_BASE_URL + "resource/' & URLENCODE(" + value + ")\"^^xl:Expr ] ;");
        //writer.write ("\"" + value + "\" ;");
        writer.newLine();
    }

    private boolean getIgnoreZeros (String metaColumn) throws XLWrapException, XLWrapEOFException{
        if (ignoreZerosRow >= 0){
            String ignoreZeroString = getCellValue (metaColumn, ignoreZerosRow);
            if (ignoreZeroString == null){
                return false;
            } else  if (ignoreZeroString.isEmpty()){
                return false;
            } else {
                return Boolean.parseBoolean(ignoreZeroString);
            }
        } else{
            return false;
        }
    }

    private void writeVocab(BufferedWriter writer, String metaColumn, String field, String dataColumn)
            throws IOException, XLWrapException, XLWrapEOFException{
        System.out.println ("write vocab " + field);
        String noExternal = "\"" + dataColumn + firstData + "\"^^xl:Expr ;";
        writer.write("	vocab:has" + field + "\t");
        String externalDoi;
        String externalSheet;
        if (externalSheetRow >= 1){
            String externalSheetName = getCellValue (metaColumn, externalSheetRow);
            if (externalSheetName == null || externalSheetName.isEmpty()) {
                //data is a value
                //ystem.out.println ("name is null");
                writer.write(noExternal);
            } else {
                if (externalSheetName.startsWith("[")){
                    externalDoi = externalSheetName.substring(1, externalSheetName.indexOf(']'));
                    externalSheet = externalSheetName.substring( externalSheetName.indexOf(']')+1);
                } else {
                    externalDoi = doi;
                    externalSheet = externalSheetName;
                }
                //dataColumn = getCellValue (metaColumn, externalColumnRow);
                //data is an external link
                System.out.println ("external");
                writer.write("[ xl:uri \"ID_URI('" + RDF_BASE_URL + field + "/" + externalDoi + "/" + externalSheet +
                    "/', " + dataColumn + firstData + ")\"^^xl:Expr ];");
            }
        } else {
            //ystem.out.println ("sheetrow = null");
            //data is a value
            writer.write(noExternal);
        }
        writer.newLine();
    }

    private String getExternalUri(String metaColumn, String field) throws XLWrapException, XLWrapEOFException{
        if (externalSheetRow < 1){
            return  RDF_BASE_URL + field + "/" + doi + "/" + sheetInURI;
        }
        String externalSheetName = getCellValue (metaColumn, externalSheetRow);
        if (externalSheetName == null || externalSheetName.isEmpty()) {
            return  RDF_BASE_URL + field + "/" + doi + "/" + sheetInURI;
        }
        String externalDoi;
        String externalSheet;
        if (externalSheetName.startsWith("[")){
            externalDoi = externalSheetName.substring(1, externalSheetName.indexOf(']'));
            externalSheet = externalSheetName.substring( externalSheetName.indexOf(']')+1);
        } else {
            externalDoi = doi;
            externalSheet = externalSheetName;
        }
        return  RDF_BASE_URL + field + "/" + externalDoi + "/" + externalSheet + "/";
    }

        //externalColumnRow
    private void writeTemplateColumn(BufferedWriter writer, String metaColumn, String dataColumn)
            throws IOException, XLWrapException, XLWrapEOFException, XLWrapMapException{
        String category = getCellValue (metaColumn, categoryRow);
        if (category == null || category.toLowerCase().equals("undefined")) {
            //ystem.out.println("Skippig column " + metaColumn + " as no Category provided");
            return;
        }
        String field = getCellValue (metaColumn, fieldRow);
        if (field == null){
            //ystem.out.println("Skippig column " + metaColumn + " as no Feild provided");
            return;
        }
        String idType = getCellValue (metaColumn, idTypeRow);
        boolean ignoreZeros  = getIgnoreZeros(metaColumn);
        String externalUri = getExternalUri(metaColumn, field);
        writeURI(writer, externalUri, category, field, idType, dataColumn,ignoreZeros);
        writer.write (" a ex:");
        writer.write (category);
        writer.write (" ;");
        writer.newLine();
        writeVocab(writer, metaColumn, field, dataColumn);
        for (int row = firstLink; row <= lastLink; row++){
            writeLink(writer, metaColumn, dataColumn, row, ignoreZeros);
        }
        for (int row = firstConstant; row <= lastConstant; row++){
            writeConstant(writer, metaColumn, row);
        }

        writer.write("	rdf:type [ xl:uri \"'" + RDF_BASE_URL + "resource/" + doi + category+ "'\"^^xl:Expr ] ;");
        writer.newLine();

        writer.write(".");
        writer.newLine();
        writer.newLine();
    }

    protected void writeTemplate(BufferedWriter writer) throws IOException, XLWrapException, XLWrapEOFException, XLWrapMapException{
        writer.write(":");
        writer.write(template());
        writer.write(" {");
        writer.newLine();
        int maxColumn = POI_Utils.alphaToIndex(LAST_DATA_COLUMN);
        for (int i = 0; i < maxColumn; i++){
            String data = POI_Utils.indexToAlpha(i);
            String meta = POI_Utils.indexToAlpha(i+1);
            writeTemplateColumn(writer, meta, data);
        }
        writer.write("}");
        writer.newLine();
    }

}


