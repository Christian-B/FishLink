package org.freshwaterlife.fishlink.xlwrap.run;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.XLWrapMaterializer;
import at.jku.xlwrap.map.MappingParser;
import at.jku.xlwrap.map.XLWrapMapping;
import com.hp.hpl.jena.rdf.model.Model;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import org.freshwaterlife.fishlink.FishLinkPaths;
import org.freshwaterlife.fishlink.FishLinkUtils;
import org.freshwaterlife.fishlink.FishLinkException;
import org.freshwaterlife.fishlink.FishLinkConstants;

/**
 * This class just holds static methods creating RDF based on a mapping file.
 * 
 * @author Christian
 */
public class MapRun {

    
    /**
     * Convenience method on top of {@link #runMap(java.lang.String, java.io.File) }
     * 
     * <ul>
     *    <li>Converts pid into an XlWrap mapping file, 
     *         assuming the mapping file is in the default place and has the default file name based on the pid.
     *    <li>Creates an rdf file based on the pid in the default place.
     *    <li>Calls {@link #runMap(java.lang.String, java.io.File) }
     * </ul>
     * @param mappingFile File that holds the mapping file
     * @param pid pid on which to base the name of the RDF file on.
     * @throws FishLinkException 
     */
    public static void runMap(String pid) throws FishLinkException{
        String mappingUrl = FishLinkPaths.MAP_FILE_ROOT + pid + ".trig";
        File rdfFile = new File (FishLinkPaths.RDF_FILE_ROOT + pid + ".rdf");
        runMap(mappingUrl, rdfFile);
    }

    /**
     * Convenience method on top of {@link #runMap(java.lang.String, java.io.File) }
     * 
     * <ul>
     *    <li>Converts the mappingFile to an XLWrap URI.
     *    <li>Calls {@link #runMap(java.lang.String, java.io.File) }
     * </ul>
     * @param mappingFile File that holds the mapping file
     * @param rdfFile File to write the rdf to.
     * @throws FishLinkException 
     */
    public static void runMap(File mappingFile, File rdfFile) throws FishLinkException{
        String mappingUrl = mappingFile.getAbsolutePath();
        runMap(mappingUrl, rdfFile);        
    }
    
    /**
     * Convenience method on top of {@link #runMap(java.lang.String, java.io.File) }
     * 
     * <ul>
     *    <li>Converts the mappingFile to an XLWrap URI.
     *    <li>Creates a file based on the pid in the default place.
     *    <li>Calls {@link #runMap(java.lang.String, java.io.File) }
     * </ul>
     * @param mappingFile File that holds the mapping file
     * @param pid pid on which to base the name of the RDF file on.
     * @throws FishLinkException 
     */
    public static void runMap(File mappingFile, String pid) throws FishLinkException{
        String mappingUrl = "File:" + mappingFile.getAbsolutePath();
        File rdfFile = new File (FishLinkPaths.RDF_FILE_ROOT + pid + ".rdf");
        runMap(mappingUrl, rdfFile);        
    }

    /**
     * Creates RDF based on a XL wrap URL to a mapping file, and saves it in the specified file.
     * 
     * <ul>
     *     <li>Creates an XLWrapMapping object by parsing the mapping URL. 
     *     <li>Creates a Jena.rdf.model based on the XLWarpMapping Object
     *     <li>Adds the Fish,Link specific prefixes to the model.
     *     <li>Writes the model to the rdfFile.
     * </ul>
     * @param mappingUrl XlWrap format URL to the mapping file.
     * @param rdfFile File to write the rdf to.
     * @throws FishLinkException 
     */
    public static void runMap(String mappingUrl, File rdfFile) throws FishLinkException{
        FishLinkUtils.report("Running map");     
        XLWrapMapping map;
        try {
            map = MappingParser.parse(mappingUrl);
        } catch (XLWrapException ex) {
            throw new FishLinkException ("Error parsing "+ mappingUrl);
        }

        XLWrapMaterializer mat = new XLWrapMaterializer();
        Model m;
        try {
            m = mat.generateModel(map);
        } catch (XLWrapException ex) {
            throw new FishLinkException ("Error generating model "+ mappingUrl , ex);
        }
        m.setNsPrefix("constant", FishLinkConstants.RDF_BASE_URL + "constant/");
        m.setNsPrefix("type", FishLinkConstants.RDF_BASE_URL + "type/");        
        m.setNsPrefix("vocab", FishLinkConstants.RDF_BASE_URL + "vocab/");
        m.setNsPrefix("resource", FishLinkConstants.RDF_BASE_URL + "resource/");
        writeRDF(m, rdfFile);
        FishLinkUtils.report("Done writing rdf file to "+ rdfFile.getAbsolutePath());
    }

    /**
     * Writes the jena.rdf.model to the file
     * @param model model based on the xlwrap mapping file.
     * @param rdfFile File to write to
     * @throws FishLinkException 
     */
    private static void writeRDF(Model model, File rdfFile) throws FishLinkException{
        File root = new File(FishLinkPaths.RDF_FILE_ROOT);
        if (!root.exists()){
            throw new FishLinkException("Unable to find RDF_FILE_ROOT. " + FishLinkPaths.RDF_FILE_ROOT + " Please create it.");
        }
        if (!root.isDirectory()){
            throw new FishLinkException("RDF_FILE_ROOT " + FishLinkPaths.RDF_FILE_ROOT + " Is not a directory.");
        }
        FileWriter writer;
        try {
            writer = new FileWriter(rdfFile);
        } catch (IOException ex) {
            throw new FishLinkException("Unable to open " + rdfFile.getAbsolutePath(), ex);
        }
        //Allowed formats are "RDF/XML", "RDF/XML-ABBREV", "N-TRIPLE", "TURTLE", (and "TTL") and "N3"
        model.write(writer, "RDF/XML");
    }
}
