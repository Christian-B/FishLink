/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.metadata;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author Christian
 */
public class CYAB_Workbook {

    Workbook poiWorkbook;

    //HashMap<String,Sheet> sheets = new HashMap<String,Sheet>();

    public CYAB_Workbook(){
        poiWorkbook = new HSSFWorkbook();
    }
    
    public CYAB_Workbook(String path) throws FileNotFoundException, IOException, InvalidFormatException{
        InputStream inputStream = new FileInputStream(path);
        poiWorkbook = WorkbookFactory.create(inputStream);
    }

    public CYAB_Sheet getSheet(String name) {
        Sheet sheet = poiWorkbook.getSheet(name);
        if (sheet == null) {
            sheet = poiWorkbook.createSheet(name);
        }
        return new CYAB_Sheet(sheet);
    }

    public CYAB_Sheet getSheetAt(int column) {
        Sheet sheet = poiWorkbook.getSheetAt(column);
        if (sheet == null) {
            return null;
        }
        return new CYAB_Sheet(sheet);
    }

    public void write(String filePath) throws FileNotFoundException, IOException {
        FileOutputStream fileOut = new FileOutputStream(filePath);
        poiWorkbook.write(fileOut);
        fileOut.close();
    }

    public Name createName() {
        return poiWorkbook.createName();
    }

    int getNumberOfSheets() {
        return poiWorkbook.getNumberOfSheets();
    }

}
