/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap.OldStuff;

import at.jku.xlwrap.common.Utils;
import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.map.expr.val.XLExprValue;
import at.jku.xlwrap.map.range.CellRange;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;

/**
 *
 * @author Christian
 */
public class XLRead {

    public static void main(String[] args) throws XLWrapException, XLWrapEOFException {
        ExecutionContext context = new ExecutionContext();

        String fileName = "File:c:/Dropbox/FISH.Link_code/tarns/TarnschemFinalOntology.xls";
        String sheetName = null; //null is first (0) sheet.
        int col = Utils.alphaToIndex("D");
        int row = 8 - 1;
        CellRange cellRange = new CellRange(fileName, sheetName, col, row);
        Cell cell = context.getCell(cellRange);
        System.out.println (cell);
        XLExprValue<?> value = Utils.getXLExprValue(cell);
        System.out.println(value);
    }
}
