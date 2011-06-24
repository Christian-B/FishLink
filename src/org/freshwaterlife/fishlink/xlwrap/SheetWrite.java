package org.freshwaterlife.fishlink.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.BufferedWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import org.freshwaterlife.fishlink.MasterFactory;
import org.freshwaterlife.fishlink.POI_Utils;

/**
 *
 * @author Christian
 */
public class SheetWrite extends AbstractSheet{

    private static String RDF_BASE_URL = "http://rdf.freshwaterlife.org/";

    private String sheet;
    private int sheetNumber;

    private String dataPath;
    private String doi;

    private String sheetInfo;

    private static NameChecker masterNameChecker;

    private HashMap<String,String> idColumns;
    private HashMap<String,String> categoryUris;
    private ArrayList<String> allColumns;

    public SheetWrite (Workbook metaWorkbook, String dataURL, String doi, String sheetName)
            throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        super(metaWorkbook.getSheet(sheetName));
        this.doi = doi;
        dataPath = dataURL;
        Workbook dataWorkbook = MasterFactory.getExecutionContext().getWorkbook(dataPath);
        String[] sheetNames = dataWorkbook.getSheetNames();
        for (int i = 0; i< sheetNames.length; i++ ){
            if (sheetNames[i].equalsIgnoreCase(sheetName)){
                sheetNumber = i;
                Sheet dataSheet = dataWorkbook.getSheet(i);
                sheetInfo = dataSheet.getSheetInfo();
                sheet = dataSheet.getName();
                //ystem.out.println("Successfully opened RawData: " + dataSheet.getName());
            }
        }
        if (sheet == null){
            throw new XLWrapMapException ("Unable to find sheet " + sheetName);
        }
        idColumns = new HashMap<String,String>();
        categoryUris = new HashMap<String,String>();
        allColumns = new ArrayList<String>();
    }

    String getSheetInfo(){
        return sheetInfo;
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
        writer.write("				xl:restriction \"A" + firstData + ":" + lastDataColumn + firstData + "\" ;");
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

    private void writeUri(BufferedWriter writer, String category, String field, String idType,
            String dataColumn, boolean ignoreZeros)
            throws IOException, XLWrapMapException, XLWrapException, XLWrapEOFException {
        if(category.toLowerCase().equals("observation")) {
            writelUriCell(writer, category, field, idType, dataColumn, ignoreZeros);
        } else if (idType == null || idType.isEmpty()  || idType.equalsIgnoreCase("n/a") ||
                idType.equalsIgnoreCase("automatic")){
            writeUri(writer, category, field, dataColumn, ignoreZeros);
        } else {
            //There is an Id reference to another column.
            if (idType.equalsIgnoreCase("row")){
                throw new XLWrapMapException("IDType row not longer supported. Foun in sheet " + metaSheet.getName());
            } else if (idType.equalsIgnoreCase("all")){
                throw new XLWrapMapException("Unexpected IDType all");
            } else {
                String idCategory = getMetaCellValueOnDataColumn(idType, categoryRow);
                String idColumn = idColumns.get(idCategory);
                writeUri(writer, idCategory, field, idColumn, ignoreZeros);
            }
        }
    }
    
    private void writeUri(BufferedWriter writer, String category, String field, String dataColumn, boolean ignoreZeros)
            throws IOException, XLWrapMapException {
        String uri = categoryUris.get(category);
        if (field.toLowerCase().equals("id")){
            writer.write("[ xl:uri \"ID_URI('" + uri + "', " + dataColumn + firstData + ","
                    + ignoreZeros + ")\"^^xl:Expr ] ");
            return;
        }
        String idColumn = idColumns.get(category);
        writeUriOther(writer, uri, idColumn, dataColumn, ignoreZeros);
    }

    private void writelUriCell(BufferedWriter writer, String category, String field, String idColumn, String dataColumn,
            boolean ignoreZeros) throws IOException, XLWrapMapException, XLWrapException, XLWrapEOFException {
        String uri = categoryUris.get(category);
        if(field.toLowerCase().equals("value")) {
            writer.write("[ xl:uri \"CELL_URI('" + uri + "', " + dataColumn + firstData + ","
                    + ignoreZeros + ")\"^^xl:Expr ] ");
            return;
        }
        if (idColumn == null || idColumn.isEmpty()){
            throw new XLWrapMapException("Data Column " + dataColumn + " with category " + category + " and field " +
                    field + " needs an id Type");
        }
//        writer.write("[ xl:uri \"OTHER_CELL_URI('" + uri + "', " + idColumn + firstData + ","
//            + dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
        writer.write("[ xl:uri \"CELL_URI('" + uri + "', " + idColumn + firstData + ","
            + dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
    }

    private void writeUriOther(BufferedWriter writer, String uri, String idColumn, String dataColumn, boolean ignoreZeros)
            throws IOException {
        if (idColumn.equalsIgnoreCase("row")){
            writer.write("[ xl:uri \"ROW_URI('" + uri + "', " + dataColumn + firstData + "," +
                    ignoreZeros + ")\"^^xl:Expr ] ");
        } else {
//            writer.write("[ xl:uri \"OTHER_ID_URI('" + uri + "', " + idColumn + firstData + "," +
//                    dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
            writer.write("[ xl:uri \"ID_URI('" + uri + "', " + idColumn + firstData + "," +
                    dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
        }
    }

    private void writeVocab (BufferedWriter writer, String vocab) throws IOException{
        if (vocab.startsWith("is") || vocab.startsWith("has")){
            writer.write("	vocab:" + vocab + "\t");
        } else {
            writer.write("	vocab:has" + vocab + "\t");
        }
    }

    private void writeValue(BufferedWriter writer, String value) throws IOException{
        if (!value.startsWith("'")){
            value = "'" + value ;
        }
        if (!value.endsWith("'")){
            value = value + "'";
        }
        writer.write("[ xl:uri \"'" + RDF_BASE_URL + "resource/' & URLENCODE(" + value + ")\"^^xl:Expr ] ;");
    }

    private void writeConstant (BufferedWriter writer, String metaColumn, int row)
            throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException{
        String feild = getCellValue ("A", row);
        if (feild == null){
            return;
        }
        String value = getCellValue (metaColumn, row);
        if (value == null){
            return;
        }
        if (value.equalsIgnoreCase("n/a")){
            return;
        }
        writeVocab (writer,	feild);
        writeValue(writer, value);
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

    private String refersToCategory(String field) throws XLWrapEOFException, XLWrapException{
       if (this.isCategory(field)) {
           return field;
       }
       return Constants.refersToCategory(field);
    }

    private void writeData(BufferedWriter writer, String metaColumn)
            throws IOException, XLWrapException, XLWrapEOFException{
        String field = getCellValue (metaColumn, fieldRow);
        writeVocab(writer, field);
        String category = refersToCategory(field);
        String dataColumn = metaToDataColumn (metaColumn);
        if (category  == null) {
            writer.write("\"" + dataColumn + firstData + "\"^^xl:Expr ;");
        } else {
            String uri = getUri(metaColumn, category);
            writer.write("[ xl:uri \"ID_URI('" + uri + "'," + dataColumn + firstData + ", false)\"^^xl:Expr ];");
        }
        writer.newLine();
    }

    private void checkName(String categery, String field)
            throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        if (masterNameChecker == null){
            masterNameChecker = new NameChecker();
        }
        masterNameChecker.checkName(sheetInfo, categery, field);
    }

    private boolean isCategory(String field) throws XLWrapException, XLWrapEOFException {
        if (masterNameChecker == null){
            masterNameChecker = new NameChecker();
        }
        return masterNameChecker.isCategory(field);
    }

    private void writeAutoRelated(BufferedWriter writer, String category, String dataColumn, boolean ignoreZeros)
            throws IOException{
        String related = Constants.autoRelatedCategory(category);
        if (related == null){
            return;
        }
        String uri = categoryUris.get(related);
        if (uri == null){
            return;
        }
        writeVocab(writer, related);
        String idColumn = idColumns.get(category);
        writeUriOther(writer, uri, idColumn, dataColumn, ignoreZeros);
        writer.write(";");
        writer.newLine();
    }

    private void writeAllRelated(BufferedWriter writer, String category, String dataColumn, boolean ignoreZeros)
            throws IOException, XLWrapException, XLWrapEOFException{
        if (!category.equalsIgnoreCase(Constants.OBSERVATION_LABEL)){
            return;
        }
        String uri = categoryUris.get(category);
        String idColumn = idColumns.get(category);
        for (String column : allColumns){
            writeData(writer, column);
        }
    }

    private boolean writeTemplateColumn(BufferedWriter writer, String metaColumn, String dataColumn)
            throws IOException, XLWrapException, XLWrapEOFException, XLWrapMapException{
        String category = getCellValue (metaColumn, categoryRow);
        String field = getCellValue (metaColumn, fieldRow);
        String idType = getCellValue (metaColumn, idTypeRow);
        String external = getExternal(metaColumn);
        if (category == null || category.toLowerCase().equals("undefined")) {
            System.out.println("Skippig column " + metaColumn + " as no Category provided");
            return false;
        }
        if (field == null){
            System.out.println("Skippig column " + metaColumn + " as no Feild provided");
            return false;
        }
        checkName(category, field);
        if (field.equalsIgnoreCase("id") && !external.isEmpty()){
            System.out.println("Skipping column " + metaColumn + " as it is an external id");
            return false;
        }
        if (idType != null && idType.equals(Constants.ALL_LABEL)){
            System.out.println("Skipping column " + metaColumn + " as it is an all column.");
            return false;
        }
        boolean ignoreZeros  = getIgnoreZeros(metaColumn);
        writeUri(writer, category, field, idType, dataColumn, ignoreZeros);
        writer.write (" a ex:");
        writer.write (category);
        writer.write (" ;");
        writer.newLine();
        writeData(writer, metaColumn);
        writeAutoRelated(writer, category, dataColumn, ignoreZeros);
        writeAllRelated(writer, category, dataColumn, ignoreZeros);
        for (int row = firstConstant; row <= lastConstant; row++){
            writeConstant(writer, metaColumn, row);
        }

        if (external.isEmpty()){
            writer.write("	rdf:type [ xl:uri \"'" + RDF_BASE_URL + "resource/" + doi + "/" + sheet + "/" 
                    + category + "/'\"^^xl:Expr ] ;");
            writer.newLine();
        }

        writer.write(".");
        writer.newLine();
        writer.newLine();
        return true;
    }

    private String getExternal(String metaColumn) throws XLWrapException, XLWrapEOFException{
        if (externalSheetRow < 1){
            return "";
        }
        String externalFeild = getCellValue (metaColumn, externalSheetRow);
        if (externalFeild == null){
            return "";
        }
        return externalFeild;
   }

   private String getCatgerogyUri(String category){
       return  RDF_BASE_URL + category + "/" + doi + "/" + sheet + "/";
   }

   private String getUri(String metaColumn , String category)
            throws XLWrapException, XLWrapEOFException{
        String externalField = getExternal(metaColumn);
        if (externalField.isEmpty()){
            return getCatgerogyUri(category);
        }
        String externalDoi;
        String externalSheet;
        if (externalField.startsWith("[")){
            externalDoi = externalField.substring(1, externalField.indexOf(']'));
            externalSheet = externalField.substring( externalField.indexOf(']')+1);
        } else {
            externalDoi = doi;
            externalSheet = externalField;
        }
        return  RDF_BASE_URL + category + "/" + externalDoi + "/" + externalSheet + "/";
    }

   private String metaToDataColumn(String metaColumn){
       return POI_Utils.indexToAlpha(POI_Utils.alphaToIndex(metaColumn)-1);
   }

   private void findId(String category, String metaColumn)
           throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        String field = getCellValue (metaColumn, fieldRow);
        String id = idColumns.get(category);
        if (field.equalsIgnoreCase("id")){
            if (id == null || id.equalsIgnoreCase("row")){
                String dataColumn = metaToDataColumn(metaColumn);
                idColumns.put(category, dataColumn);
                categoryUris.put(category, getUri(metaColumn, category));
            } else {
                throw new XLWrapMapException("Found two different id columns of type " + category);
            }
        } else {
            if (id == null){
                idColumns.put(category, "row");
            }//else leave the column or "row" already there.
            if (categoryUris.get(category) == null){
                categoryUris.put(category, getCatgerogyUri(category));
            }
        }
    }

   private void findAll(String category, String metaColumn)
           throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        String idColumn = getCellValue (metaColumn, idTypeRow);
        if (idColumn == null || !idColumn.equalsIgnoreCase("all")){
            return;
        }
        if (category.equalsIgnoreCase(Constants.OBSERVATION_LABEL)){
            String field = getCellValue (metaColumn, fieldRow);
            if (field == null || field.isEmpty()){
                throw new XLWrapException ("All id.Value Column " + metaColumn + " missing a field value");
            }
            allColumns.add(metaColumn);
        } else {
            throw new XLWrapException ("All id.Value Column only supported for Categeroy " +
                    Constants.OBSERVATION_LABEL);
        }
    }

   private void findIds() throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        int maxColumn = POI_Utils.alphaToIndex(lastDataColumn);
        for (int i = 0; i < maxColumn; i++){
            String metaColumn = POI_Utils.indexToAlpha(i);
            String category = getCellValue (metaColumn, categoryRow);
            if (category == null || category.isEmpty()){
                //do nothing
            } else {
                findId(category, metaColumn);
                findAll(category, metaColumn);
            }
        }
     }

    protected void writeTemplate(BufferedWriter writer) throws IOException, XLWrapException, XLWrapEOFException, XLWrapMapException{
        System.out.println("Writing template for "+sheet);
        findIds();
        writer.write(":");
        writer.write(template());
        writer.write(" {");
        writer.newLine();
        int maxColumn = POI_Utils.alphaToIndex(lastDataColumn);
        boolean foundColumn = false;
        for (int i = 0; i < maxColumn; i++){
            String data = POI_Utils.indexToAlpha(i);
            String meta = POI_Utils.indexToAlpha(i+1);
            if (writeTemplateColumn(writer, meta, data)){
                foundColumn = true;
            }
        }
        if (!foundColumn){
            throw new XLWrapMapException("No mappable columns found in sheet " +  metaSheet.getName());
        }
        writer.write("}");
        writer.newLine();
    }

}


