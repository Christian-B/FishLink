/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.BufferedWriter;
import java.io.IOException;
import java.util.HashMap;
import uk.co.brenn.metadata.MetaDataCreator;
import uk.co.brenn.metadata.POI_Utils;

/**
 *
 * @author Christian
 */
public class SheetWrite extends AbstractSheet{

    private static String RDF_BASE_URL = "http://rdf.fba.org.uk/";

    private String sheet;
    private int sheetNumber;

    private String dataPath;
    private String doi;

    private String sheetInfo;

    private static NameChecker masterNameChecker;

    private HashMap<String,String> idColumns;
    private HashMap<String,String> categoryUris;

    //, String mapFileName, String rdfFileName
    public SheetWrite (Workbook metaWorkbook, String dataURL, String doi, String sheetName)
            throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        super(metaWorkbook.getSheet(sheetName));
        //ystem.out.println(this.lastDataColumn);
        this.doi = doi;
        dataPath = dataURL;
        Workbook dataWorkbook = MetaDataCreator.getExecutionContext().getWorkbook(dataPath);
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

    /*
    private void XwriteURI(BufferedWriter writer, String externalUri, String category, String field, String idType, String dataColumn,
            boolean ignoreZeros) throws IOException, XLWrapMapException {
        String uri;
        if (externalUri.isEmpty()){
            uri = RDF_BASE_URL + category + "/" + doi + "/" + sheetInURI;
        } else {
            uri = externalUri;
        }
        //ystem.out.println("writeURI " + uri + " " + category + " " + field + " " + idType + " " + dataColumn);
        if (field.toLowerCase().equals("id")){
            writer.write("[ xl:uri \"ID_URI('" + uri + "', " + dataColumn + firstData + ","
                    + ignoreZeros + ")\"^^xl:Expr ] ");
            return;
        }
        if(field.toLowerCase().equals("value")) {
            writer.write("[ xl:uri \"CELL_URI('" + uri + "', " + dataColumn + firstData + ","
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
            writer.write("[ xl:uri \"OTHER_CELL_URI('" + uri + "', " + idType + firstData + ","
                    + dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
            return;
        }
        if (idType.equalsIgnoreCase("ROW")){
            writer.write("[ xl:uri \"ROW_URI('" + uri + "row', " + dataColumn + firstData + "," +
                    ignoreZeros + ")\"^^xl:Expr ] ");
        } else {
            writer.write("[ xl:uri \"OTHER_ID_URI('" + uri + "', " + idType + firstData + "," +
                    dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
        }
    }
 */
    private void writeURI(BufferedWriter writer, String category, String field, String idType, String dataColumn,
            boolean ignoreZeros) throws IOException, XLWrapMapException, XLWrapException, XLWrapEOFException {
        if (idType == null || idType.isEmpty() || idType.equalsIgnoreCase("n/a")){
            writeURI(writer, category, field, dataColumn, ignoreZeros);
        } else {
            //There is an Id reference to another column.
            if (idType.equalsIgnoreCase("row")){
                throw new XLWrapMapException("IDType row not longer supported. Foun in sheet " + metaSheet.getName());
            } else if (idType.equalsIgnoreCase("all")){
                throw new XLWrapMapException("Unexpected IDType all");
            } else {
                String idCategory = getMetaCellValueOnDataColumn(idType, categoryRow);
                String idColumn = idColumns.get(idCategory);
                System.out.println (idType + " " + idCategory + "/"+ idColumn);
                writeURI(writer, idCategory, field, idColumn, ignoreZeros);
            }
        }
    }
    
    private void writeURI(BufferedWriter writer, String category, String field, String dataColumn, boolean ignoreZeros)
            throws IOException, XLWrapMapException {
        String uri = categoryUris.get(category);
        System.out.println("writeURI " + uri + " " + category + " " + field + " " + dataColumn);
        if (field.toLowerCase().equals("id")){
            writer.write("[ xl:uri \"ID_URI('" + uri + "', " + dataColumn + firstData + ","
                    + ignoreZeros + ")\"^^xl:Expr ] ");
            return;
        }
        if(category.toLowerCase().equals("observation")) {
            if(field.toLowerCase().equals("value")) {
                writer.write("[ xl:uri \"CELL_URI('" + uri + "', " + dataColumn + firstData + ","
                        + ignoreZeros + ")\"^^xl:Expr ] ");
                return;
            } else {
                throw new XLWrapMapException ("Obeservation column with field " + field +
                        " must have and value column specified.");
            }
        }
        String idType = idColumns.get(category);
        if (idType.equalsIgnoreCase("row")){
            writer.write("[ xl:uri \"ROW_URI('" + uri + "row', " + dataColumn + firstData + "," +
                    ignoreZeros + ")\"^^xl:Expr ] ");
        } else {
            writer.write("[ xl:uri \"OTHER_ID_URI('" + uri + "', " + idType + firstData + "," +
                    dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
        }
    }

    private void writeVocab (BufferedWriter writer, String vocab) throws IOException{
        //if (vocab.contains("/")){
        //    int error = 1/0;
        //}
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

    /*private void writeLink (BufferedWriter writer, String metaColumn, String dataColumn,
            int row, boolean ignoreZeros)
            throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException{
        String feild = getCellValue ("A", row);
        if (feild == null){
            return;
        }
        String nextFeild = getCellValue ("A", row + 1);
        String link = getCellValue (metaColumn, row);
        //ystem.out.println("writeLink " + metaColumn + "\t" + feild + "\t" + link);
        if (link == null){
           //ystem.out.println("Skippig column " + metaColumn + " " + feild + "row as it is blank");
            return;
        }
        if (link.equalsIgnoreCase("n/a")){
            return;
        }
        if (feild.contains("/")){
            String nextValue = getCellValue (metaColumn, row+1);
            if (nextValue == null){
                feild = feild.substring(0, feild.indexOf('/'));
            } else {
                feild = feild.substring(feild.indexOf('/')+1);
            }
        }
        writeVocab(writer, feild);
        if (link.startsWith("=")){
            link = link.replaceAll("=", "");
            writeValue(writer, link); //Which includes the ;
        } else {
            String exernalUri;
            if (link.equals("row")){
                exernalUri = "";
            } else {
                exernalUri = getOtherExternalId(link, feild);
            }
            writeURI(writer, exernalUri, "link", feild, link, dataColumn, ignoreZeros);
            writer.write(" ;");
        }
        writer.newLine();
    }
*/
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
        writeVocab (writer,	feild);
        writeValue(writer, value);
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

    private void writeData(BufferedWriter writer, String category, String field, String metaColumn, String dataColumn)
            throws IOException, XLWrapException, XLWrapEOFException{
        //ystem.out.println ("write vocab " + field);
        writeVocab(writer, field);
        String uri = getExternalUri(metaColumn, category);
        System.out.println(metaColumn + ": " + uri);
        if (uri == null) {
            writer.write("\"" + dataColumn + firstData + "\"^^xl:Expr ;");
        } else {
            writer.write("[ xl:uri \"ID_URI('" + uri + "'," + dataColumn + firstData + ", false)\"^^xl:Expr ];");
        }
        writer.newLine();
    }

    /*private String getOtherExternalId(String otherColumn, String category) throws XLWrapException, XLWrapEOFException{
        String metaColumn = POI_Utils.indexToAlpha(POI_Utils.alphaToIndex(otherColumn)+1);
        return getExternalUri1(metaColumn, category);
    }*/

    /*private String getExternalUri(String metaColumn, String category)
            throws XLWrapException, XLWrapEOFException{
        if (externalSheetRow < 1){
            return "";
        }
        String externalFeild = getCellValue (metaColumn, externalSheetRow);
        if (externalFeild == null){
            return "";
        }
        if (externalFeild.isEmpty()) {
            return "";
        }
        String externalDoi;
        String externalSheet;
        if (externalFeild.startsWith("[")){
            externalDoi = externalFeild.substring(1, externalFeild.indexOf(']'));
            externalSheet = externalFeild.substring( externalFeild.indexOf(']')+1);
        } else {
            externalDoi = doi;
            externalSheet = externalFeild;
        }
        return  RDF_BASE_URL + category + "/" + externalDoi + "/" + externalSheet + "/";
    }*/

    private void checkName(String categery, String field)
            throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        if (masterNameChecker == null){
            masterNameChecker = new NameChecker();
        }
        masterNameChecker.checkName(sheetInfo, categery, field);
    }
        //externalColumnRow
    private boolean writeTemplateColumn(BufferedWriter writer, String metaColumn, String dataColumn)
            throws IOException, XLWrapException, XLWrapEOFException, XLWrapMapException{
        String category = getCellValue (metaColumn, categoryRow);
        if (category == null || category.toLowerCase().equals("undefined")) {
            System.out.println("Skippig column " + metaColumn + " as no Category provided");
            return false;
        }
        String field = getCellValue (metaColumn, fieldRow);
        if (field == null){
            System.out.println("Skippig column " + metaColumn + " as no Feild provided");
            return false;
        }
        String external = getExternal(metaColumn);
        if (field.equalsIgnoreCase("id") && !external.isEmpty()){
            System.out.println("Skipping column " + metaColumn + " as it is an external id");
            return false;
        }
        checkName(category, field);
        String idType = getCellValue (metaColumn, idTypeRow);
        boolean ignoreZeros  = getIgnoreZeros(metaColumn);
        //String externalUri = getExternalUri(metaColumn, category);
        writeURI(writer, category, field, idType, dataColumn,ignoreZeros);
        writer.write (" a ex:");
        writer.write (category);
        writer.write (" ;");
        writer.newLine();
        writeData(writer, category, field, metaColumn, dataColumn);
        //if (firstLink >= 0){
        //    for (int row = firstLink; row <= lastLink; row++){
        //        writeLink(writer, metaColumn, dataColumn, row, ignoreZeros);
        //    }
        //}
        for (int row = firstConstant; row <= lastConstant; row++){
            writeConstant(writer, metaColumn, row);
        }

        if (external.isEmpty()){
            writer.write("	rdf:type [ xl:uri \"'" + RDF_BASE_URL + "resource/" + doi + category+ "'\"^^xl:Expr ] ;");
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

   private String getExternalUri(String metaColumn, String category)
            throws XLWrapException, XLWrapEOFException{
        String externalField = getExternal(metaColumn);
        if (externalField.isEmpty()){
            return null;
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

   private String getIdUri(String metaColumn, String category)
            throws XLWrapException, XLWrapEOFException{
        String uri = getExternalUri(metaColumn, category);
        if (uri == null) {
            return getCatgerogyUri(category);
        }
        return uri;
    }

   private void findIds() throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        int maxColumn = POI_Utils.alphaToIndex(lastDataColumn);
        for (int i = 0; i < maxColumn; i++){
            //ystem.out.println(i);
            String metaColumn = POI_Utils.indexToAlpha(i);
            String category = getCellValue (metaColumn, categoryRow);
            if (category == null || category.isEmpty()){
                //do nothing
            } else {
                String field = getCellValue (metaColumn, fieldRow);
                String id = idColumns.get(category);
                if (field.equalsIgnoreCase("id")){
                    if (id == null || id.equalsIgnoreCase("row")){
                        String dataColumn = POI_Utils.indexToAlpha(i-1);
                        idColumns.put(category, dataColumn);
                        categoryUris.put(category, getIdUri(metaColumn, category));
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
        }
        //for (String key : idColumns.keySet()) {
            //ystem.out.println("Key: " + key + ", idColumn: " + idColumns.get(key) +
        //            ", uri: " + categoryUris.get(key));
        //}
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
            //ystem.out.println("writeTemplate " + i + lastDataColumn);
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


