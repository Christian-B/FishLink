package uk.co.brenn.metadata;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author Christian
 */
public class CYAB_Workbook {

    Workbook poiWorkbook;

    public CYAB_Workbook(){
        poiWorkbook = new HSSFWorkbook();
    }
    
    CYAB_Sheet getSheet(String name) {
        Sheet sheet = poiWorkbook.getSheet(name);
        if (sheet == null) {
            sheet = poiWorkbook.createSheet(name);
        }
        return new CYAB_Sheet(sheet);
    }

    void write(String filePath) throws FileNotFoundException, IOException {
        FileOutputStream fileOut = new FileOutputStream(filePath);
        poiWorkbook.write(fileOut);
        fileOut.close();
    }

    Name createName() {
        return poiWorkbook.createName();
    }

}
