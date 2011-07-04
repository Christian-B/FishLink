package org.freshwaterlife.fishlink.xlwrap;

import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import org.freshwaterlife.fishlink.FishLinkPaths;
import org.freshwaterlife.fishlink.POI_Utils;

/**
 *
 * @author Christian
 */
public class WorkbookWrite {

    private String pid;

    private SheetWrite[] sheetWrites;

    public WorkbookWrite (String metaPid, String dataPid) throws XLWrapMapException{
        Workbook workbook = POI_Utils.getWorkbookOnPid(metaPid);
        Workbook dataWorkbook = POI_Utils.getWorkbookOnPid(dataPid);
        pid = dataPid;
        String[] sheetNames = workbook.getSheetNames();
        sheetWrites = new SheetWrite[sheetNames.length - 1];
        int j = 0;
        for (int i = 0; i< sheetNames.length; i++ ){
            if (sheetNames[i].equals("MetaData") || sheetNames[i].equals("Lists")) {
                //do nothing
            } else {
                //ystem.out.println("  " + i + " " + sheetNames.length);
                sheetWrites[j] = new SheetWrite(workbook, dataPid, sheetNames[i]);
                MasterReader masterReader = new MasterReader ();
                masterReader.check(sheetWrites[j]);
                j++;
            }
        }
    }

    /*public WorkbookWrite (String metaFileName) throws XLWrapMapException{
        Workbook workbook;
        Sheet metaData;
        workbook = POI_Utils.getWorkbook("file:" + metaRoot + metaFileName);
        metaData = POI_Utils.getSheet(workbook, "MetaData");
        Cell cell = POI_Utils.getCell (metaData, 1, 0);
        String dataFileName = POI_Utils.getText(cell);
        Workbook dataWorkbook = POI_Utils.getWorkbook("file:" + dataRoot + dataFileName);
        cell = POI_Utils.getCell(metaData, 1, 1);
        doi = POI_Utils.getText(cell);
        String[] sheetNames = workbook.getSheetNames();
        sheetWrites = new SheetWrite[sheetNames.length - 2];
        int j = 0;
        for (int i = 0; i< sheetNames.length; i++ ){
            if (sheetNames[i].equals("MetaData") || sheetNames[i].equals("Lists")) {
                //do nothing
            } else {
                sheetWrites[j] = new SheetWrite(workbook, "file:" + dataRoot + dataFileName, doi, sheetNames[i]);
                MasterReader masterReader = new MasterReader ();
                masterReader.check(sheetWrites[j]);
                j++;
            }
        }
    }*/

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

     public void writeMap() throws XLWrapMapException {
        System.out.println("write map");
        File mapFile = new File(FishLinkPaths.MAP_FILE_ROOT);
        if (!mapFile.exists()){
            throw new XLWrapMapException("Unable to find MAP_FILE_ROOT. " + FishLinkPaths.MAP_FILE_ROOT + " Please create it.");
        }
        mapFile = new File(FishLinkPaths.MAP_FILE_ROOT + pid + ".trig");
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
    }

}
