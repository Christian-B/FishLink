package org.freshwaterlife.fishlink.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import java.io.BufferedWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import org.freshwaterlife.fishlink.FishLinkUtils;

/**
 *
 * @author Christian
 */
public class SheetWrite extends AbstractSheet{

    private int sheetNumber;

    private String dataPath;
    private String pid;

    private NameChecker masterNameChecker;

    private HashMap<String,String> idColumns;
    private HashMap<String,String> categoryUris;
    private ArrayList<String> allColumns;

    public SheetWrite (NameChecker nameChecker, Sheet annotatedSheet, int sheetNumber, String url, String pid)
            throws XLWrapMapException{
        super(annotatedSheet);
        this.pid = pid;
        this.sheetNumber = sheetNumber;
        this.dataPath = url;
        this.masterNameChecker = nameChecker;
        idColumns = new HashMap<String,String>();
        categoryUris = new HashMap<String,String>();
        allColumns = new ArrayList<String>();
    }

    String getSheetInfo(){
        return sheet.getSheetInfo();
    }

    private String template(){
        return sheet.getName() + "template";
    }

    protected void writeMapping(BufferedWriter writer) throws XLWrapMapException{
        try {
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
            writer.write("				xl:restriction \"B" + firstData + ":" + lastDataColumn + firstData + "\" ;");
            writer.newLine();
            writer.write("				xl:steps \"1\" ;");
            writer.newLine();
            writer.write("			] ;");
            writer.newLine();
            writer.write("		]");
            writer.newLine();
            writer.write("	] ;");
            writer.newLine();
        } catch (IOException ex) {
            throw new XLWrapMapException("Unable to write mapping", ex);
        }
    }

    private void writeUri(BufferedWriter writer, String category, String field, String idType,
            String dataColumn, boolean ignoreZeros) throws XLWrapMapException {
        if(category.equalsIgnoreCase(Constants.OBSERVATION_LABEL)) {
            writelUriCell(writer, category, field, idType, dataColumn, ignoreZeros);
        } else if (idType == null || idType.isEmpty()  || idType.equalsIgnoreCase("n/a") ||
                idType.equalsIgnoreCase(Constants.AUTOMATIC_LABEL)){
            writeUri(writer, category, field, dataColumn, ignoreZeros);
        } else {
            //There is an Id reference to another dataColumn.
            if (idType.equalsIgnoreCase("row")){
                throw new XLWrapMapException("IDType row not longer supported. Found in sheet " + sheet.getSheetInfo());
            } else if (idType.equalsIgnoreCase(Constants.ALL_LABEL)){
                throw new XLWrapMapException("Unexpected IDType " + Constants.ALL_LABEL);
            } else {
                String idCategory = this.getCellValue(idType, categoryRow);
                String idColumn = idColumns.get(idCategory);
                writeUri(writer, idCategory, field, idColumn, ignoreZeros);
            }
        }
    }
    
    private void writeUri(BufferedWriter writer, String category, String field, String dataColumn, boolean ignoreZeros)
            throws XLWrapMapException {
        String uri = categoryUris.get(category);
        if (field.toLowerCase().equals(Constants.ID_LABEL)){
            try {
                writer.write("[ xl:uri \"ID_URI('" + uri + "', " + dataColumn + firstData + ","
                        + ignoreZeros + ")\"^^xl:Expr ] ");
            } catch (IOException ex) {
                throw new XLWrapMapException("Unable to write uri", ex);
            }
            return;
        }
        String idColumn = idColumns.get(category);
        writeUriOther(writer, uri, idColumn, dataColumn, ignoreZeros);
    }

    private void writelUriCell(BufferedWriter writer, String category, String field, String idColumn, String dataColumn,
            boolean ignoreZeros) throws XLWrapMapException {
        String uri = categoryUris.get(category);
        if(field.equalsIgnoreCase(Constants.VALUE_LABEL)) {
            try {
                writer.write("[ xl:uri \"CELL_URI('" + uri + "', " + dataColumn + firstData + ","
                        + ignoreZeros + ")\"^^xl:Expr ] ");
            } catch (IOException ex) {
                throw new XLWrapMapException("Unable to write uri", ex);
            }                
            return;
        }
        if (idColumn == null || idColumn.isEmpty()){
            throw new XLWrapMapException(sheet.getSheetInfo() + " Data Column " + dataColumn + " with category " +
                    category + " and field " + field + " needs an id Type");
        }
//        writer.write("[ xl:uri \"OTHER_CELL_URI('" + uri + "', " + idColumn + firstData + ","
//            + dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
        try {
            writer.write("[ xl:uri \"CELL_URI('" + uri + "', " + idColumn + firstData + ","
                + dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
        } catch (IOException ex) {
            throw new XLWrapMapException("Unable to write uri", ex);
        }
    }

    private void writeUriOther(BufferedWriter writer, String uri, String idColumn, String dataColumn, boolean ignoreZeros)
            throws XLWrapMapException {
        //"row" is added for automatic ones where no id column found.
        try {
            if (idColumn.equalsIgnoreCase("row")){
                writer.write("[ xl:uri \"ROW_URI('" + uri + "', " + dataColumn + firstData + "," +
                        ignoreZeros + ")\"^^xl:Expr ] ");
            } else {
//            writer.write("[ xl:uri \"OTHER_ID_URI('" + uri + "', " + idColumn + firstData + "," +
//                    dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
                writer.write("[ xl:uri \"ID_URI('" + uri + "', " + idColumn + firstData + "," +
                        dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
            }
        } catch (IOException ex) {
            throw new XLWrapMapException("Unable to write uri", ex);
        }
    }

    private void writeVocab (BufferedWriter writer, String vocab) throws XLWrapMapException {
        try {
            if (vocab.startsWith("is") || vocab.startsWith("has")){
                writer.write("	vocab:" + vocab + "\t");
            } else {
                writer.write("	vocab:has" + vocab + "\t");
            }
        } catch (IOException ex) {
            throw new XLWrapMapException("Unable to write vocab", ex);
        }            
    }

    private void writeRdfType (BufferedWriter writer, String type) throws XLWrapMapException {
        try {
           writer.write ("	rdf:type");
           writeType(writer, type);
        } catch (IOException ex) {
            throw new XLWrapMapException("Unable to write vocab", ex);
        }            
    }

    private void writeType (BufferedWriter writer, String type) throws XLWrapMapException {
        try {
           writer.write (" type:" + type);
           writer.write (" ;");
           writer.newLine();
        } catch (IOException ex) {
            throw new XLWrapMapException("Unable to write vocab", ex);
        }            
    }

    private void writeValue(BufferedWriter writer, String value) throws XLWrapMapException{
        if (!value.startsWith("'")){
            value = "'" + value ;
        }
        if (!value.endsWith("'")){
            value = value + "'";
        }
        try {
            writer.write("[ xl:uri \"'" + Constants.RDF_BASE_URL + "constant/' & URLENCODE(" + value + ")\"^^xl:Expr ] ;");
            writer.newLine();
        } catch (IOException ex) {
            throw new XLWrapMapException("Unable to write vocab", ex);
        }            
    }

    private void writeConstant (BufferedWriter writer, String category, String column, int row) 
            throws XLWrapMapException{
        String field = getCellValue ("A", row);
        if (field == null){
            return;
        }
        String value = getCellValue (column, row);
        if (value == null){
            return;
        }
        if (value.equalsIgnoreCase("n/a")){
            return;
        }
        if (Constants.isRdfTypeField(field)){
            masterNameChecker.checkSubType(sheet.getSheetInfo(), category, value);
            writeRdfType(writer, value);
        } else {
            masterNameChecker.checkConstant(sheet.getSheetInfo(), field, value);
            writeVocab (writer, field);
            writeValue(writer, value);
        }
    }

    private boolean getIgnoreZeros (String column) throws XLWrapMapException {
        if (ignoreZerosRow >= 0){
            String ignoreZeroString = getCellValue (column, ignoreZerosRow);
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

    private String refersToCategory(String field) throws XLWrapMapException{
       if ( masterNameChecker.isCategory(field)) {
           return field;
       }
       return Constants.refersToCategory(field);
    }

    private void writeData(BufferedWriter writer, String column) throws XLWrapMapException {
        String field = getCellValue (column, fieldRow);
        writeVocab(writer, field);
        String category = refersToCategory(field);
        try {
            if (category  == null) {
                writer.write("\"" + column + firstData + "\"^^xl:Expr ;");
            } else {
                String uri = getUri(column, category);
                writer.write("[ xl:uri \"ID_URI('" + uri + "'," + column + firstData + ", false)\"^^xl:Expr ];");
            }
            writer.newLine();
            }  catch (IOException ex) {
                throw new XLWrapMapException("Unable to write data", ex);
            }            
    }

    private void writeAutoRelated(BufferedWriter writer, String category, String column, boolean ignoreZeros)
            throws XLWrapMapException {
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
        writeUriOther(writer, uri, idColumn, column, ignoreZeros);
        try {
            writer.write(";");
            writer.newLine();
        }  catch (IOException ex) {
            throw new XLWrapMapException("Unable to write ", ex);
        }            
    }

    private void writeAllRelated(BufferedWriter writer, String category, boolean ignoreZeros)
            throws XLWrapMapException {
        if (!category.equalsIgnoreCase(Constants.OBSERVATION_LABEL)){
            return;
        }
        String uri = categoryUris.get(category);
        String idColumn = idColumns.get(category);
        for (String column : allColumns){
            writeData(writer, column);
        }
    }

    private boolean writeTemplateColumn(BufferedWriter writer, String column)
            throws XLWrapMapException{
        String category = getCellValue (column, categoryRow);
        //ystem.out.println(category + " " + metaColumn + " " + categoryRow);
        String field = getCellValue (column, fieldRow);
        String idType = getCellValue (column, idTypeRow);
        String external = getExternal(column);
        if (category == null || category.toLowerCase().equals("undefined")) {
            FishLinkUtils.report("Skippig column " + column + " as no Category provided");
        }
        if (field == null){
            FishLinkUtils.report("Skippig column " + column + " as no Feild provided");
            return false;
        }
        masterNameChecker.checkName(sheet.getSheetInfo(), category, field);
        if (field.equalsIgnoreCase("id") && !external.isEmpty()){
            FishLinkUtils.report("Skipping column " + column + " as it is an external id");
            return false;
        }
        if (idType != null && idType.equals(Constants.ALL_LABEL)){
            FishLinkUtils.report("Skipping column " + column + " as it is an all column.");
            return false;
        }
        boolean ignoreZeros  = getIgnoreZeros(column);
        writeUri(writer, category, field, idType, column, ignoreZeros);
        try {
            writer.write (" a ");
            writeType (writer, category);
        }  catch (IOException ex) {
            throw new XLWrapMapException("Unable to write a type", ex);
        }            
        writeData(writer, column);
        writeAutoRelated(writer, category, column, ignoreZeros);
        writeAllRelated(writer, category, ignoreZeros);
        for (int row = firstConstant; row <= lastConstant; row++){
            writeConstant(writer,  category, column, row);
        }

        try {
            if (external.isEmpty()){
               writeRdfType (writer, pid + "_" + sheet.getName() + "_" + category);
             }

            writer.write(".");
            writer.newLine();
            writer.newLine();
        }  catch (IOException ex) {
            throw new XLWrapMapException("Unable to other type ", ex);
        }            
        return true;
    }

    private String getExternal(String metaColumn) throws XLWrapMapException {
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
       return  Constants.RDF_BASE_URL + "resource/" + category + "_" + pid + "_" + sheet.getName() + "/";
   }

   private String getUri(String metaColumn , String category) throws XLWrapMapException {
        String externalField = getExternal(metaColumn);
        if (externalField.isEmpty()){
            return getCatgerogyUri(category);
        }
        String externalPid;
        String externalSheet;
        if (externalField.startsWith("[")){
            externalPid = externalField.substring(1, externalField.indexOf(']'));
            externalSheet = externalField.substring( externalField.indexOf(']')+1);
        } else {
            externalPid = pid;
            externalSheet = externalField;
        }
        return  Constants.RDF_BASE_URL + "resource/" + category + "_" + pid + "_" + externalSheet + "/";
    }

   private void findId(String category, String column) throws XLWrapMapException{
        String field = getCellValue (column, fieldRow);
        String id = idColumns.get(category);
        if (field.equalsIgnoreCase("id")){
            if (id == null || id.equalsIgnoreCase("row")){
                System.out.println (column + " " +category + "    " + column);
                idColumns.put(category, column);
                categoryUris.put(category, getUri(column, category));
            } else {
                throw new XLWrapMapException("Found two different id columns of type " + category);
            }
        } else {
            if (id == null){
                System.out.println (column + " " + category + "    row");
                idColumns.put(category, "row");
            }//else leave the column or "row" already there.
            if (categoryUris.get(category) == null){
                categoryUris.put(category, getCatgerogyUri(category));
            }
        }
    }

   private void findAll(String category, String metaColumn) throws XLWrapMapException{
        String idColumn = getCellValue (metaColumn, idTypeRow);
        if (idColumn == null || !idColumn.equalsIgnoreCase("all")){
            return;
        }
        if (category.equalsIgnoreCase(Constants.OBSERVATION_LABEL)){
            String field = getCellValue (metaColumn, fieldRow);
            if (field == null || field.isEmpty()){
                throw new XLWrapMapException ("All id.Value Column " + metaColumn + " missing a field value");
            }
            allColumns.add(metaColumn);
        } else {
            throw new XLWrapMapException ("All id.Value Column only supported for Categeroy " +
                    Constants.OBSERVATION_LABEL);
        }
    }

   private void findIds() throws XLWrapMapException{
        int maxColumn = FishLinkUtils.alphaToIndex(lastDataColumn);
        for (int i = 1; i < maxColumn; i++){
            String metaColumn = FishLinkUtils.indexToAlpha(i);
            String category = getCellValue (metaColumn, categoryRow);
            if (category == null || category.isEmpty()){
                //do nothing
            } else {
                findId(category, metaColumn);
                findAll(category, metaColumn);
            }
        }
     }

    protected void writeTemplate(BufferedWriter writer) throws XLWrapMapException{
        FishLinkUtils.report("Writing template for "+sheet.getSheetInfo());
        findIds();
        try {
            writer.write(":");
            writer.write(template());
            writer.write(" {");
            writer.newLine();
        }  catch (IOException ex) {
            throw new XLWrapMapException("Unable to write template ", ex);
        }            
        int maxColumn = FishLinkUtils.alphaToIndex(lastDataColumn);
        boolean foundColumn = false;
        //Start at 1 to ignore label column;
        for (int i = 1; i < maxColumn; i++){
            String column = FishLinkUtils.indexToAlpha(i);
            if (writeTemplateColumn(writer, column)){
                foundColumn = true;
            }
        }
        if (!foundColumn){
            throw new XLWrapMapException("No mappable columns found in sheet " +  sheet.getSheetInfo());
        }
        try {
            writer.write("}");
            writer.newLine();
        }  catch (IOException ex) {
            throw new XLWrapMapException("Unable to write template end ", ex);
        }            
    }

}


