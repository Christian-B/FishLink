package org.freshwaterlife.fishlink.xlwrap;

import org.freshwaterlife.fishlink.FishLinkException;
import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import org.freshwaterlife.fishlink.FishLinkConstants;
import org.freshwaterlife.fishlink.FishLinkPaths;
import org.freshwaterlife.fishlink.FishLinkUtils;

/**
 * Write the mapping file for a whole workbook.
 * 
 * Creates the Workbook and the classes based on the MetaMaster.
 * Does the workbook level writing and then calls the Sheets to write themselves.
 * 
 * @author Christian
 */
public class WorkbookWrite {

    /**
     * Workbook PID used to create a unique map file name
     */
    private String pid;

    /**
     * The individual sheet (wrappers) which do most of the actually writing.
     */
    private SheetWrite[] sheetWrites;

    /**
     * Constructor of the workbook wrapper as well as the wrappers around the sheets in this workbook and the MetaMaster. 
     * 
     * Also checks the data workbook against the MetaMaster.
     * 
     * @param dataUrl URL (in xlwrap format) to the Workbook that holds the Annotated Data
     * @param pid Unique identifier to this data. Used in URIs and file naming.
     * @param masterUrl URL (in xlwrap format) to the MetaMaster Workbook
     * @throws FishLinkException 
     */
    public WorkbookWrite(String dataUrl, String pid, String masterUrl) throws FishLinkException{
        ExecutionContext context = new ExecutionContext();
        this.pid = pid;
        //Wrap the MetaMaster
        Sheet masterListSheet;
        try {
            masterListSheet = context.getSheet(masterUrl, FishLinkConstants.LIST_SHEET);
        } catch (XLWrapException ex) {
            throw new FishLinkException("Error opening the vocabulary sheet " + FishLinkConstants.LIST_SHEET + 
                    " in ExcelSheet " + masterUrl, ex);
        }
        NameChecker nameChecker = new NameChecker(masterListSheet);
        Sheet masterDropdownSheet;
        try {
            masterDropdownSheet = context.getSheet(masterUrl, FishLinkConstants.DROP_DOWN_SHEET);
        } catch (XLWrapException ex) {
            throw new FishLinkException("Error opening the dropdown sheet " + FishLinkConstants.DROP_DOWN_SHEET+ 
                    " in ExcelSheet " + masterUrl, ex);
        }
        MasterReader masterReader = new MasterReader(masterDropdownSheet);
        //Wrap the workbook
        Workbook annotatedWorkbook;
        try {
            annotatedWorkbook = context.getWorkbook(dataUrl);
        } catch (XLWrapException ex) {
            throw new FishLinkException("Error opening the workbook " + dataUrl, ex);
        }
        String[] sheetNames = annotatedWorkbook.getSheetNames();
        sheetWrites = new SheetWrite[sheetNames.length - 1];
        int j = 0;
        for (int i = 0; i< sheetNames.length; i++ ){
            Sheet sheet;
            try {
                sheet = annotatedWorkbook.getSheet(i);
            } catch (XLWrapException ex) {
                throw new FishLinkException ("Unable to open sheet " + i + " in " + dataUrl, ex);
            }
            if (sheet.getName().equals("MetaData") || sheet.getName().equals("Lists")) {
                //do nothing
            } else {
                sheetWrites[j] = new SheetWrite(nameChecker, sheet, i, dataUrl, pid );
                masterReader.check(sheetWrites[j]);
                j++;
            }
        }
    }

    /**
     * Writes the data independent part of the mapping file.
     * @param writer
     * @throws FishLinkException 
     */
    private void writePrefix (BufferedWriter writer) throws FishLinkException {
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
            writer.write("@prefix constant:	<" + FishLinkConstants.RDF_BASE_URL + "constant/> .");
            writer.newLine();
            writer.write("@prefix type:	<" + FishLinkConstants.RDF_BASE_URL + "type/> .");
            writer.newLine();
            writer.write("@prefix resource:	<" + FishLinkConstants.RDF_BASE_URL + "resource/> .");
            writer.newLine();
            writer.write("@prefix vocab:	<" + FishLinkConstants.RDF_BASE_URL + "vocab/> .");
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
            throw new FishLinkException ("Unable to write prefix", ex);
        }
    }

    /**
     * Writes the mapping file.
     * 
     * Creates a mapping file based on the PID.
     * Writes the prefix and then asks each wrapped sheet to write its part of the mapping file.
     * @return The Mapping file.
     * @throws FishLinkException 
     */
     public File writeMap() throws FishLinkException {
        FishLinkUtils.report("write map");
        File mapFile = new File(FishLinkPaths.MAP_FILE_ROOT);
        if (!mapFile.exists()){
            throw new FishLinkException("Unable to find MAP_FILE_ROOT. " + FishLinkPaths.MAP_FILE_ROOT + " Please create it.");
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
            throw new FishLinkException ("Unable to write mapping file.", ex);
        }
        FishLinkUtils.report("Done writing map file");
        return mapFile;
    }

}
