/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.IOException;
import uk.co.brenn.xlwrap.expr.func.BrennRegister;

/**
 *
 * @author christian
 */
public class ChristianDemo {
    private static String XLS_FILE_PATH = "c:/Dropbox/FISH.Link/data/TarnschemFinalOntology.xls";

    private static String MAP_FILE_NAME = "test.trig";

    private static String RDF_FILE_NAME = "testOutput.rdf";

    private static String doi = "doi12345";

    public static void main(String[] args) throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException {
        BrennRegister.register();
        MapWrite mapWrite = new MapWrite(XLS_FILE_PATH, doi);
        mapWrite.writeMap(MAP_FILE_NAME);
        mapWrite.runMap(MAP_FILE_NAME, RDF_FILE_NAME);
    }

}
