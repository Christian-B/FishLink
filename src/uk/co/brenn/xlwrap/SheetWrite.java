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

    //Excell columns and Rows used here are Excell based and not 0 based as XLWrap uses internally
    private String LAST_DATA_COLUMN = "X";
    private int CATEGORY_ROW = 1;
    private int FIELD_ROW = 2;
    private int ID_TYPE_ROW = 3;
    private int ignoreZerosRow = -1;
    private int firstLink;
    private int lastLink;
    private int firstConstant;
    private int lastConstant;
    private int firstData;

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
        metaSheet = metaWorkbook.getSheet(sheetName);
 
        String columnA;
        int row = 4;
        columnA = getCellValue("A",row);
        //ystem.out.println("Read from ColumnA " + columnA);

        if (columnA.toLowerCase().contains("ignore zeros")){
            ignoreZerosRow = row;
            row++;
            columnA = getCellValue("A",row);
        }
        if (columnA.toLowerCase().contains("links")){
            firstLink = row + 1;
        } else {
            throw new XLWrapMapException ("Ignoring sheet " + sheetName + " did not find \"links\" seperator in expected place");
        }
        do{
            row++;
            columnA = getCellValue("A",row);
        } while (!columnA.toLowerCase().contains("constant"));
        lastLink = row -1;
        firstConstant = row + 1;
        do{
            row++;
            columnA = getCellValue("A",row);
        } while (columnA !=null && !columnA.isEmpty());
        lastConstant = row - 1;
        do{
            row++;
            columnA = getCellValue("A",row);
        } while (columnA == null || !columnA.toLowerCase().contains("datastart"));
        firstData = 2;
    }

    /*
    private void writePrefix (BufferedWriter writer) throws IOException{
        writer.write("@prefix rdfs:   <http://www.w3.org/2000/01/rdf-schema#> .");
        writer.newLine();
        writer.write("@prefix rdf:    <http://www.w3.org/1999/02/22-rdf-syntax-ns#> .");
        writer.newLine();
        writer.write("@prefix xsd:    <http://www.w3.org/2001/XMLSchema#> .");
        writer.newLine();
        writer.write("@prefix owl:    <http://www.w3.org/2002/07/owl#> .");
        writer.newLine();
        writer.write("@prefix foaf:	<http://xmlns.com/foaf/0.1/> .");
        writer.newLine();
        writer.write("@prefix ex:	<" + RDF_BASE_URL + "resource/> .");
        writer.newLine();
        writer.write("@prefix vocab:	<" + RDF_BASE_URL + "vocab/resource/> .");
        writer.newLine();
        writer.write("@prefix dc:     <http://purl.org/dc/elements/1.1/> .");
        writer.newLine();
        writer.write("@prefix xl:	<http://purl.org/NET/xlwrap#> .");
        writer.newLine();
        writer.write("@prefix scv:	<http://purl.org/NET/scovo#> .");
        writer.newLine();
        writer.write("@prefix :       <http://myApplication/configuration#> .");
        writer.newLine();
        writer.newLine();
    }
  */
    private String template(){
        return metaSheet.getName() + "template";
//        return "template";
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

    private void writeExternalId (BufferedWriter writer, String type, String idType, String dataColumn,
            boolean ignoreZeros) throws IOException{
        String externalDOI;
        if (idType.startsWith("[")){
            externalDOI = idType.substring(1, idType.indexOf(']'));
            //ystem.out.println ("externalDOI = " + externalDOI);
            idType = idType.substring(idType.indexOf(']')+1);
        } else {
            externalDOI = doi;
        }
        //ystem.out.println(idType);
        //ystem.out.println(idType.indexOf('!'));
        String externalSheet = idType.substring(0, idType.indexOf('!'));
        String externalDataColumn = idType.substring(idType.indexOf('!')+1);
        writer.write("[ xl:uri \"ID_URI('" + RDF_BASE_URL + type + "/" + externalDOI + "/" + externalSheet + "/', " +
                dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
        return;
    }

    private void writeOtherID(BufferedWriter writer, String type, String idType, String dataColumn,
            boolean ignoreZeros) throws IOException{
        String externalDOI;
        if (idType.startsWith("[")){
            externalDOI = idType.substring(1, idType.indexOf(']'));
            //ystem.out.println ("externalDOI = " + externalDOI);
            idType = idType.substring(idType.indexOf(']')+1);
        } else {
            externalDOI = doi;
        }
        String externalSheet = sheetInURI;
        String externalDataColumn = dataColumn;
        if (idType.contains("!")){
            externalSheet = idType.substring(0, idType.indexOf('!')) + "/";
            externalDataColumn = idType.substring(idType.indexOf('!')+1);
        }
        //ystem.out.println(idType);
        //ystem.out.println(idType.indexOf('!'));
        writer.write("[ xl:uri \"OTHER_ID_URI('" + RDF_BASE_URL + type + "/" + externalDOI + "/" + externalSheet + "', " +
            externalDataColumn + firstData + "," + dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
    }

    private void writeURI(BufferedWriter writer, String type, String feild, String idType, String dataColumn,
            boolean ignoreZeros) throws IOException, XLWrapMapException {
        System.out.println("writeURI " + type + " " + feild + " " + idType + " " + dataColumn);
        if (feild.toLowerCase().equals("external_id")){
            writeExternalId (writer, type, idType, dataColumn, ignoreZeros);
            return;
        }
        if (feild.toLowerCase().equals("id")){
            writer.write("[ xl:uri \"ID_URI('" + RDF_BASE_URL + type + "/" + doi + "/" + sheetInURI + "', " +
                    dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
            return;
        } 
        if(feild.toLowerCase().equals("value")) {
            writer.write("[ xl:uri \"CELL_URI('" + RDF_BASE_URL + type + "/" + doi + "/" + sheetInURI + "', " +
                dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
            return;
        } 
        if ((idType == null) || (idType.isEmpty())){
            throw new XLWrapMapException("Column " + dataColumn + " in Sheet " + metaSheet.getName() +
                    " has an unexpected emtpy IDType");
        }
        if (idType.equalsIgnoreCase("n/a")) {
            throw new XLWrapMapException("Column " + dataColumn + " in Sheet " + metaSheet.getName() +
                    " has an unexpected n/a IDType");
        }
        if (idType == null){
            //ystem.out.println("Skippig column " + dataColumn + " as no idType provided");
        } else if (idType.equalsIgnoreCase("ROW")){
            writer.write("[ xl:uri \"ROW_URI('" + RDF_BASE_URL + type + "/" + doi + "/" + sheetInURI + "row', " +
                    dataColumn + firstData + "," + ignoreZeros + ")\"^^xl:Expr ] ");
        } else {
            writeOtherID(writer, type, idType, dataColumn, ignoreZeros);
        }
    }

    private void writeLink (BufferedWriter writer, String metaColumn, String dataColumn, int row, boolean ignoreZeros)
            throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException{
        String feild = getCellValue ("A", row);
        if (feild == null){
            return;
        }
        String link = getCellValue (metaColumn, row);
        System.out.println("writeLink " + metaColumn + "\t" + feild + "\t" + link);
        if (link == null){
           // System.out.println("Skippig column " + metaColumn + " " + feild + "row as it is blank");
            return;
        }
        if (link.equalsIgnoreCase("n/a")){
            return;
        }
        writer.write("	vocab:has" + feild + "\t");
        writeURI(writer, feild, feild, link, dataColumn, ignoreZeros);
        writer.write(" ;");
        writer.newLine();
    }

    private void writeConstant (BufferedWriter writer, String metaColumn, int row)
            throws XLWrapException, XLWrapEOFException, IOException{
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
        writer.write("	vocab:has" + feild + "\t");
        writer.write("[ xl:uri \"'" + RDF_BASE_URL + "resource/' & URLENCODE('" + value + "')\"^^xl:Expr ] ;");
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

    private void writeVocab(BufferedWriter writer, String field, String dataColumn) throws IOException{
        if (!field.toLowerCase().startsWith("external")){
            writer.write("	vocab:has" + field + "\t\"" + dataColumn + firstData + "\"^^xl:Expr ;");
            writer.newLine();
        }
    }

    private void writeTemplateColumn(BufferedWriter writer, String metaColumn, String dataColumn)
            throws IOException, XLWrapException, XLWrapEOFException, XLWrapMapException{
        String category = getCellValue (metaColumn, CATEGORY_ROW);
        if (category == null){
            //ystem.out.println("Skippig column " + metaColumn + " as no Category provided");
            return;
        }
        String field = getCellValue (metaColumn, FIELD_ROW);
        if (field == null){
            //ystem.out.println("Skippig column " + metaColumn + " as no Feild provided");
            return;
        }
        String idType = getCellValue (metaColumn, ID_TYPE_ROW);
        boolean ignoreZeros  = getIgnoreZeros(metaColumn);
        writeURI(writer, category, field, idType, dataColumn,ignoreZeros);
        writer.write (" a ex:");
        writer.write (category);
        writer.write (" ;");
        writer.newLine();
        writeVocab(writer, field, dataColumn);
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
        for (char meta = 'B'; meta < 'Y'; meta++){
            int charValue = Character.valueOf(meta);
            String data = String.valueOf( (char) (charValue - 1));
            //ystem.out.println(meta + "  " + data);
            writeTemplateColumn(writer, "" + meta, data);
        }
        writer.write("}");
        writer.newLine();
    }

  /*  public void writeMap(String mapFileName) throws IOException, XLWrapMapException, XLWrapException, XLWrapEOFException{
        File mapFile = new File(MAP_FILE_ROOT + mapFileName);
        BufferedWriter mapWriter = new BufferedWriter(new FileWriter(mapFile));
        writePrefix(mapWriter);
        writeMapping(mapWriter);
        writeTemplate(mapWriter);
        mapWriter.close();
        System.out.println("Done writing map file");
    }

    public void runMap(String mapFileName, String rdfFileName) throws XLWrapException, IOException{
        XLWrapMapping map = MappingParser.parse(MAP_FILE_ROOT + mapFileName);

        XLWrapMaterializer mat = new XLWrapMaterializer();
        Model m = mat.generateModel(map);
        m.setNsPrefix("ex", RDF_BASE_URL);

        File out = new File (RDF_FILE_ROOT + rdfFileName);
        FileWriter writer = new FileWriter(out);
                //"RDF/XML", "RDF/XML-ABBREV", "N-TRIPLE", "TURTLE", (and "TTL") and "N3"
        //m.write(writer, "RDF/XML", RDF_BASE_URL);
        m.write(writer, "RDF/XML");
        System.out.println("Done writing rdf file to "+ out.getAbsolutePath());
    }
*/
}
