/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap;

import uk.co.brenn.xlwrap.expr.func.BrennRegister;
import at.jku.xlwrap.common.Utils;
import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.exec.XLWrapMaterializer;
import at.jku.xlwrap.map.MappingParser;
import at.jku.xlwrap.map.XLWrapMapping;
import at.jku.xlwrap.map.expr.func.FunctionRegistry;
import at.jku.xlwrap.map.expr.val.XLExprValue;
import at.jku.xlwrap.map.range.CellRange;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import com.hp.hpl.jena.rdf.model.Model;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;

/**
 *
 * @author Christian
 */
public class MapWrite {

    private String MAP_FILE_NAME = "output/mappings/test2.trig";

    private String XLS_FILE_NAME = "c:/Dropbox/FISH.Link_code/tarns/TarnschemFinalOntology.xls";

    private String RDF_FILE_NAME = "output/rdf/testOutput.xml";

    //Urgently needs changing to an FBA address
    private String RDF_BASE_URL = "http://rpc466.cs.man.ac.uk:2020/";

    private String PREFIX = "TarnsSchema";

    //Excell columns and Rows used here are Excell based and not 0 based as XLWrap uses internally

    private String LAST_DATA_COLUMN = "X";

    private int CATEGORY_ROW = 1;

    private int FIELD_ROW = 2;

    private int ID_TYPE_ROW = 3;

    private int firstLink;
    private int lastLink;
    private int firstConstant;
    private int lastConstant;
    private int firstData;

    ExecutionContext context;

    public MapWrite () throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        context  = new ExecutionContext();
        String columnA;
        int row = 4;
        columnA = getCellValue("A",row);
        System.out.println(columnA);
        if (columnA.toLowerCase().contains("links")){
            firstLink = row + 1;
        } else {
            throw new XLWrapMapException ("Did not find \"links\" seperator in expected place");
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
        firstData = row;
    }

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

    private void writeMapping(BufferedWriter writer, File xlsFile) throws IOException, XLWrapMapException{
        if (!xlsFile.exists()){
            throw new XLWrapMapException("File " + xlsFile + " does not exist");
        }
        String fileName = xlsFile.getAbsolutePath();
        fileName = fileName.replace('\\','/');
        writer.write("# mapping");
        writer.newLine();
        writer.write("{ [] a xl:Mapping ;");
        writer.newLine();
        writer.write("	xl:offline \"false\"^^xsd:boolean ;");
        writer.newLine();
        writer.write("	xl:template [");
        writer.newLine();
        writer.write("		xl:fileName \"File:" + fileName + "\" ;");
        writer.newLine();
        writer.write("		xl:sheetNumber \"0\" ;");
        writer.newLine();
        writer.write("		xl:templateGraph :template ;");
        writer.newLine();
        writer.write("		xl:transform [");
        writer.newLine();
        writer.write("			a rdf:Seq ;");
        writer.newLine();
        writer.write("			rdf:_1 [");
        writer.newLine();
        writer.write("				a xl:RowShift ;");
        writer.newLine();
        writer.write("				xl:restriction \"B" + firstData + ":" + LAST_DATA_COLUMN + firstData + "\" ;");
        writer.newLine();
        writer.write("				xl:steps \"1\" ;");
        writer.newLine();
        writer.write("			] ;");
        writer.newLine();
        writer.write("		]");
        writer.newLine();
        writer.write("	] .");
        writer.newLine();
        writer.write("}");
        writer.newLine();
        writer.newLine();
    }

    private String getCellValue (String column, int row) throws XLWrapException, XLWrapEOFException{
        String sheetName = null; //null is first (0) sheet.
        int col = Utils.alphaToIndex(column);
        int actualRow = row - 1;
        CellRange cellRange = new CellRange("File:" +XLS_FILE_NAME, sheetName, col, actualRow);
        Cell cell = context.getCell(cellRange);
        XLExprValue<?> value = Utils.getXLExprValue(cell);
        if (value == null){
            return null;
        }
        return value.toString().replace("\"","");
    }

    private void writeLink (BufferedWriter writer, String column, int row)
            throws XLWrapException, XLWrapEOFException, IOException{
        String feild = getCellValue ("A", row);
        if (feild == null){
            return;
        }
        String link = getCellValue (column, row);
        System.out.println(column + "\t" + feild + "\t" + link);
        if (link == null){
            System.out.println("Skippig column " + column + " " + feild + "row as it is blank");
            return;
        }
        if (link.equalsIgnoreCase("n/a")){
            return;
        }
        writer.write("	vocab:has" + feild + "\t");
        boolean rowBased = link.equalsIgnoreCase("ROW");
        if (rowBased){
            writer.write("[ xl:uri \"ROW_URI('" + RDF_BASE_URL + feild + "/', " + column + firstData +
                    ")\"^^xl:Expr ] ;");
        } else {
            writer.write("[ xl:uri \"CELL_URI('" + RDF_BASE_URL + feild + "/', " + link + firstData +
                    ")\"^^xl:Expr ] ;");
        }
        writer.newLine();
    }

    private void writeConstant (BufferedWriter writer, String column, int row)
            throws XLWrapException, XLWrapEOFException, IOException{
        String feild = getCellValue ("A", row);
        if (feild == null){
            return;
        }
        String value = getCellValue (column, row);
        System.out.println(column + "\t" + feild + "\t" + value);
        if (value == null){
            System.out.println("Skippig column " + column + " " + feild + "row as it is blank");
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

    private void writeTemplateColumn(BufferedWriter writer, String column)
            throws IOException, XLWrapException, XLWrapEOFException{
        String category = getCellValue (column, CATEGORY_ROW);
        if (category == null){
            System.out.println("Skippig column " + column + " as no Category provided");
            return;
        }
        String feild = getCellValue (column, FIELD_ROW);
        if (feild == null){
            System.out.println("Skippig column " + column + " as no Feild provided");
            return;
        }
        String idType = getCellValue (column, ID_TYPE_ROW);
        if (idType == null){
            System.out.println("Skippig column " + column + " as no idType provided");
            return;
        }
        boolean rowBased = idType.equalsIgnoreCase("ROW");
        if (rowBased){
            writer.write("	[ xl:uri \"ROW_URI('" + RDF_BASE_URL + category + "/', " + column + firstData +
                    ")\"^^xl:Expr ] a ex:" + category + " ;");
        } else {
            writer.write("	[ xl:uri \"CELL_URI('" + RDF_BASE_URL + category + "/', " + column + firstData +
                    ")\"^^xl:Expr ] a ex:" + category + " ;");
        }
        writer.newLine();
        writer.write("	vocab:has" + feild + "\t\"" + column + firstData + "\"^^xl:Expr ;");
        writer.newLine();
        for (int row = firstLink; row <= lastLink; row++){
            writeLink(writer, column, row);
        }
        for (int row = firstConstant; row <= lastConstant; row++){
            writeConstant(writer, column, row);
        }

        writer.write("	rdf:type [ xl:uri \"'" + RDF_BASE_URL + "resource/" + PREFIX + category+ "'\"^^xl:Expr ] ;");
        writer.newLine();

        writer.write(".");
        writer.newLine();
        writer.newLine();
    }

    private void writeTemplate(BufferedWriter writer) throws IOException, XLWrapException, XLWrapEOFException{
        writer.write(":template {");
        writer.newLine();
        for (char x = 'B'; x < 'Y'; x++){
            writeTemplateColumn(writer, ""+x);
        }
        writer.write("}");
        writer.newLine();
    }

    private void writeMap() throws IOException, XLWrapMapException, XLWrapException, XLWrapEOFException{
        File mapFile = new File(MAP_FILE_NAME);
        File xlsFile = new File(XLS_FILE_NAME);
        BufferedWriter mapWriter = new BufferedWriter(new FileWriter(mapFile));
        writePrefix(mapWriter);
        writeMapping(mapWriter, xlsFile);
        writeTemplate(mapWriter);
        mapWriter.close();
        System.out.println("Done writing map file");
    }

    private void runMap() throws XLWrapException, IOException{
        XLWrapMapping map = MappingParser.parse(MAP_FILE_NAME);

        XLWrapMaterializer mat = new XLWrapMaterializer();
        Model m = mat.generateModel(map);
        m.setNsPrefix("ex", RDF_BASE_URL);

        File out = new File (RDF_FILE_NAME);
        FileWriter writer = new FileWriter(out);
                //"RDF/XML", "RDF/XML-ABBREV", "N-TRIPLE", "TURTLE", (and "TTL") and "N3"
        //m.write(writer, "RDF/XML", RDF_BASE_URL);
        m.write(writer, "RDF/XML");
        System.out.println("Done writing rdf file");
    }

    public static void main(String[] args) throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException {
        BrennRegister.register();
        MapWrite mapWrite = new MapWrite();
        mapWrite.writeMap();
        mapWrite.runMap();
    }
}
