/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap;

import at.jku.xlwrap.common.Utils;
import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.exec.XLWrapMaterializer;
import at.jku.xlwrap.map.MappingParser;
import at.jku.xlwrap.map.XLWrapMapping;
import at.jku.xlwrap.map.expr.val.XLExprValue;
import at.jku.xlwrap.map.range.CellRange;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import com.hp.hpl.jena.rdf.model.Model;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

/**
 *
 * @author Christian
 */
public class WorkbookWrite {

    private static String MAP_FILE_ROOT = "output/mappings/";

    private static String RDF_FILE_ROOT = "output/rdf/";

    private static String RDF_BASE_URL = "http://rdf.fba.org.uk/";

    private ExecutionContext context;
 
    private String workbookPath;
    private String doi;

    SheetWrite[] sheetWrites;

    //, String mapFileName, String rdfFileName
    public WorkbookWrite (ExecutionContext context, String workbookPath, String doi) throws XLWrapException, XLWrapEOFException, XLWrapMapException{
        this.doi = doi;
        Workbook workbook = context.getWorkbook(workbookPath);
        String[] sheetNames = workbook.getSheetNames();
        sheetWrites = new SheetWrite[sheetNames.length];
        for (int i = 0; i< sheetNames.length; i++ ){
            sheetWrites[i] = new SheetWrite (context, workbookPath, doi, i);
        }
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

     public void writeMap(String mapFileName) throws IOException, XLWrapMapException, XLWrapException, XLWrapEOFException{
        File mapFile = new File(MAP_FILE_ROOT + mapFileName);
        BufferedWriter mapWriter = new BufferedWriter(new FileWriter(mapFile));
        writePrefix(mapWriter);

        mapWriter.write("# mapping");
        mapWriter.newLine();
        mapWriter.write("{ [] a xl:Mapping ;");
        mapWriter.newLine();
        for (int i = 0; i < sheetWrites.length; i++){
            System.out.println(i);
            sheetWrites[i].writeMapping(mapWriter);
        }
        mapWriter.write("}");
        mapWriter.newLine();
        mapWriter.newLine();

        for (int i = 0; i < sheetWrites.length; i++){
            sheetWrites[i].writeTemplate(mapWriter);
         }
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

}
