package org.freshwaterlife.fishlink.metadatacreator;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.freshwaterlife.fishlink.FishLinkException;

/**
 * Wrapper around a Poi Workbook (
 * 
 * Currently uses HSSF but could be changed to XSSF without issue.
 * Main role is to create FishLinkSheet, Create names and catch Exceptions in writing
 * @author Christian
 */
public class FishLinkWorkbook {

    /**
     * Apache poi workbook which does the actual work
     */
    private Workbook poiWorkbook;

    /**
     * Creates a Workbook wrapping an Apache Poi HSSF workbook
     */
    FishLinkWorkbook(){
        poiWorkbook = new HSSFWorkbook();
    }
    
    /**
     * Gets or creates a Poi Sheet with this name and wraps it in a FISHLinkSheet.
     * @param name Name of the workbook
     * @return A Sheet with the given name.
     */
    FishLinkSheet getSheet(String name) {
        Sheet sheet = poiWorkbook.getSheet(name);
        if (sheet == null) {
            sheet = poiWorkbook.createSheet(name);
        }
        return new FishLinkSheet(sheet);
    }

    /**
     * Writes the workbook to file.
     * 
     * Will always be *.xls format regardless of the extention of the File.
     * @param file A file with Ideally a *.xls Extension.
     * @throws FishLinkException 
     */
    void write(File file) throws FishLinkException {
        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream(file);
        } catch (IOException ex) {
            throw new FishLinkException("Unable to open "+ file.getAbsolutePath(), ex);
        }
        try {
            poiWorkbook.write(fileOut);
        } catch (IOException ex) {
            throw new FishLinkException("Unable to write "+ file.getAbsolutePath(), ex);
        }
        try {
            fileOut.close();
        } catch (IOException ex) {
            throw new FishLinkException("Unable to close "+ file.getAbsolutePath(), ex);
        }

    }

    /**
     * Create a blank Name (Excel Named Range)
     * @return a blank Name (Excel Named Range)
     */
    Name createName() {
        return poiWorkbook.createName();
    }

}
