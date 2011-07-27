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
 *
 * @author Christian
 */
public class MapRun {

    public static void runMap(String pid) throws FishLinkException{
        String mappingUrl = FishLinkPaths.MAP_FILE_ROOT + pid + ".trig";
        File rdfFile = new File (FishLinkPaths.RDF_FILE_ROOT + pid + ".rdf");
        runMap(mappingUrl, rdfFile);
    }

    public static void runMap(File mappingFile, File rdfFile) throws FishLinkException{
        String mappingUrl = mappingFile.getAbsolutePath();
        runMap(mappingUrl, rdfFile);        
    }
    
    public static void runMap(File mappingFile, String pid) throws FishLinkException{
        String mappingUrl = "File:" + mappingFile.getAbsolutePath();
        File rdfFile = new File (FishLinkPaths.RDF_FILE_ROOT + pid + ".rdf");
        runMap(mappingUrl, rdfFile);        
    }

    private static void writeRDF(Model m, File rdfFile) throws FishLinkException{
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
                //"RDF/XML", "RDF/XML-ABBREV", "N-TRIPLE", "TURTLE", (and "TTL") and "N3"
        m.write(writer, "RDF/XML");
    }

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

}
