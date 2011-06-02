package uk.co.brenn.test;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFName;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import uk.co.brenn.metadata.CYAB_Sheet;
import uk.co.brenn.metadata.CYAB_Workbook;
import uk.co.brenn.test.NamedRange;

/**
 *
 * @author Christian
 */
public class HSSF_TEST {

    public static void main1(String[] args) throws IOException{
        Workbook wb = new HSSFWorkbook();
        //Workbook wb = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet listSheet = wb.createSheet("lists");
        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet1.createRow((short)0);
        // Create a cell and put a value in it.
        Cell cell = row.createCell(0);
        cell.setCellValue(1);

        CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
        
        Name rangeName = wb.createName();
        rangeName.setNameName("list1");
        rangeName.setRefersToFormula("'lists'!$B$1:$B$3");
        DVConstraint dvConstraint = DVConstraint.createFormulaListConstraint("list1");
        HSSFDataValidation dataValidation = new HSSFDataValidation(addressList, dvConstraint);
        dataValidation.setSuppressDropDownArrow(false);
        dataValidation.createPromptBox("Prompt Title", "This is what you must do");
        dataValidation.setShowPromptBox(true);
        sheet1.addValidationData(dataValidation);
        wb.write(fileOut);
        fileOut.close();
    }

    public static void main(String[] args) throws IOException{
        CYAB_Workbook workbook = new CYAB_Workbook();
        //Workbook wb = new XSSFWorkbook();
        CYAB_Sheet sheet1 = workbook.getSheet("data");

        sheet1.setValue ("A", 1, "Assay");

        CellRangeAddressList addressList = new CellRangeAddressList(1, 1, 0, 0);

        ListWriter.writeLists(workbook);
        Name rangeName = workbook.createName();
        rangeName.setNameName("ROW_A");
        rangeName.setRefersToFormula("'data'!$A$1:$A$1");

//        DVConstraint dvConstraint = DVConstraint.createFormulaListConstraint("INDIRECT(SUBSTITUTE(ROW_A,\" \",\"_\"))");
        DVConstraint dvConstraint = DVConstraint.createFormulaListConstraint("INDIRECT(SUBSTITUTE($a$1,\" \",\"_\"))");
        HSSFDataValidation dataValidation = new HSSFDataValidation(addressList, dvConstraint);
        dataValidation.setSuppressDropDownArrow(false);
        dataValidation.createPromptBox("Prompt Title", "This is what you must do");
        dataValidation.setShowPromptBox(true);
        sheet1.addValidationData(dataValidation);

        //CellRangeAddressList addressList2 = new CellRangeAddressList(2, 2, 0, 0);
        //dvConstraint = DVConstraint.createFormulaListConstraint("INDIRECT(Assay)");
        //dataValidation = new HSSFDataValidation(addressList2, dvConstraint);
        //dataValidation.setSuppressDropDownArrow(false);
        //dataValidation.createPromptBox("Prompt Title", "This is what you must do");
        //dataValidation.setShowPromptBox(true);
        //sheet1.addValidationData(dataValidation);

        workbook.write("workbook.xls");
     }

}
