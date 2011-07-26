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

/**
 *
 * @author Christian
 */
public class MapRun {

    public static void runMap(String pid) throws FishLinkException{
        FishLinkUtils.report("Running map");     
        XLWrapMapping map;
        try {
            map = MappingParser.parse(FishLinkPaths.MAP_FILE_ROOT + pid + ".trig");
        } catch (XLWrapException ex) {
            throw new FishLinkException ("Error parsing "+ pid );
        }

        XLWrapMaterializer mat = new XLWrapMaterializer();
        Model m;
        try {
            m = mat.generateModel(map);
        } catch (XLWrapException ex) {
            throw new FishLinkException ("Error generating model "+ pid , ex);
        }
        m.setNsPrefix("constant", FishLinkPaths.RDF_BASE_URL + "constant/");
        m.setNsPrefix("type", FishLinkPaths.RDF_BASE_URL + "type/");        
        m.setNsPrefix("vocab", FishLinkPaths.RDF_BASE_URL + "vocab/");
        m.setNsPrefix("resource", FishLinkPaths.RDF_BASE_URL + "resource/");

        File out = new File (FishLinkPaths.RDF_FILE_ROOT);
        if (!out.exists()){
            throw new FishLinkException("Unable to find RDF_FILE_ROOT. " + FishLinkPaths.RDF_FILE_ROOT + " Please create it.");
        }
        out = new File (FishLinkPaths.RDF_FILE_ROOT + pid + ".rdf");
        FileWriter writer;
        try {
            writer = new FileWriter(out);
        } catch (IOException ex) {
            throw new FishLinkException("Unable to open " + FishLinkPaths.RDF_FILE_ROOT + pid + ".rdf", ex);
        }
                //"RDF/XML", "RDF/XML-ABBREV", "N-TRIPLE", "TURTLE", (and "TTL") and "N3"
        m.write(writer, "RDF/XML");
        FishLinkUtils.report("Done writing rdf file to "+ out.getAbsolutePath());
    }

}
