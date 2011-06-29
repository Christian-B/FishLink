package org.freshwaterlife.fishlink.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import org.freshwaterlife.fishlink.FishLinkPaths;
import org.freshwaterlife.fishlink.MasterFactory;

/**
 *
 * @author Christian
 */
public class WorkbookWrite {

    private String doi;

    private SheetWrite[] sheetWrites;

    private static String dataRoot;

    private static String metaRoot;

    public static void setRoots(String metaRoot, String dataRoot){
        WorkbookWrite.dataRoot = dataRoot;
        WorkbookWrite.metaRoot = metaRoot;
    }

    public WorkbookWrite (String metaFileName) throws XLWrapMapException{
        Workbook workbook;
        Sheet metaData;
        try {
            workbook = MasterFactory.getExecutionContext().getWorkbook(metaRoot + metaFileName);
            metaData = workbook.getSheet("MetaData");
            Cell cell;
            try{
                cell = metaData.getCell(1, 0);
            } catch (NullPointerException e) {
                throw new XLWrapMapException("Workbook: " + metaFileName + " does not have a \"MetaData\" sheet.");
            }
            String dataFileName = cell.getText();
            Workbook dataWorkbook = MasterFactory.getExecutionContext().getWorkbook(dataRoot + dataFileName);
            cell = metaData.getCell(1, 1);
            doi = cell.getText();
            String[] sheetNames = workbook.getSheetNames();
            sheetWrites = new SheetWrite[sheetNames.length - 2];
            int j = 0;
            for (int i = 0; i< sheetNames.length; i++ ){
                if (sheetNames[i].equals("MetaData") || sheetNames[i].equals("Lists")) {
                    //do nothing
                } else {
                    sheetWrites[j] = new SheetWrite(workbook, dataRoot + dataFileName, doi, sheetNames[i]);
                    MasterReader masterReader = new MasterReader ();
                    masterReader.check(sheetWrites[j]);
                    j++;
                }
            }
        } catch (XLWrapException ex) {
            throw new XLWrapMapException ("Unable to create workbook", ex);
        } catch (XLWrapEOFException ex) {
            throw new XLWrapMapException ("Unable to create workbook", ex);
        }
    }

    private void writePrefix (BufferedWriter writer) throws XLWrapMapException {
        try {
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
            writer.write("@prefix constant:	<" + FishLinkPaths.RDF_BASE_URL + "constant/> .");
            writer.newLine();
            writer.write("@prefix type:	<" + FishLinkPaths.RDF_BASE_URL + "type/> .");
            writer.newLine();
            writer.write("@prefix resource:	<" + FishLinkPaths.RDF_BASE_URL + "resource/> .");
            writer.newLine();
            writer.write("@prefix vocab:	<" + FishLinkPaths.RDF_BASE_URL + "vocab/> .");
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
        } catch (IOException ex) {
            throw new XLWrapMapException ("Unable to write prefix", ex);
        }
    }

     public String writeMap() throws XLWrapMapException {
        System.out.println("write map");
        File mapFile = new File(FishLinkPaths.MAP_FILE_ROOT);
        if (!mapFile.exists()){
            throw new XLWrapMapException("Unable to find MAP_FILE_ROOT. " + FishLinkPaths.MAP_FILE_ROOT + " Please create it.");
        }
        mapFile = new File(FishLinkPaths.MAP_FILE_ROOT + doi + ".trig");
        BufferedWriter mapWriter;
        try {
            mapWriter = new BufferedWriter(new FileWriter(mapFile));
            writePrefix(mapWriter);

            mapWriter.write("# mapping");
            mapWriter.newLine();
            mapWriter.write("{ [] a xl:Mapping ;");
            mapWriter.newLine();
            for (int i = 0; i < sheetWrites.length; i++){
                sheetWrites[i].writeMapping(mapWriter);
            }
            mapWriter.write("}");
            mapWriter.newLine();
            mapWriter.newLine();

            for (int i = 0; i < sheetWrites.length; i++){
                sheetWrites[i].writeTemplate(mapWriter);
            }
            mapWriter.close();
        } catch (IOException ex) {
            throw new XLWrapMapException ("Unable to write mapping file.", ex);
        }
        System.out.println("Done writing map file");
        return doi;
    }

}
