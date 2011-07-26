package org.freshwaterlife.fishlink.metadatacreator;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.freshwaterlife.fishlink.FishLinkException;

/**
 *
 * @author Christian
 */
public class FishLinkWorkbook {

    Workbook poiWorkbook;

    public FishLinkWorkbook(){
        poiWorkbook = new HSSFWorkbook();
    }
    
    FishLinkSheet getSheet(String name) {
        Sheet sheet = poiWorkbook.getSheet(name);
        if (sheet == null) {
            sheet = poiWorkbook.createSheet(name);
        }
        return new FishLinkSheet(sheet);
    }

    void write(String filePath) throws FishLinkException {
        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream(filePath);
        } catch (IOException ex) {
            throw new FishLinkException("Unable to open "+ filePath, ex);
        }
        try {
            poiWorkbook.write(fileOut);
        } catch (IOException ex) {
            throw new FishLinkException("Unable to write "+ filePath, ex);
        }
        try {
            fileOut.close();
        } catch (IOException ex) {
            throw new FishLinkException("Unable to close "+ filePath, ex);
        }

    }

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

    Name createName() {
        return poiWorkbook.createName();
    }

}
