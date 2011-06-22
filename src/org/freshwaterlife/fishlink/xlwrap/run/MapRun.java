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
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

/**
 *
 * @author Christian
 */
public class MapRun {

    public static void runMap(String doi) throws XLWrapException, IOException, XLWrapMapException{
        System.out.println("Running map");
        XLWrapMapping map = MappingParser.parse(FishLinkPaths.MAP_FILE_ROOT + doi + ".trig");

        XLWrapMaterializer mat = new XLWrapMaterializer();
        Model m = mat.generateModel(map);
        m.setNsPrefix("ex", FishLinkPaths.RDF_BASE_URL);

        File out = new File (FishLinkPaths.RDF_FILE_ROOT);
        if (!out.exists()){
            throw new XLWrapMapException("Unable to find RDF_FILE_ROOT. " + FishLinkPaths.RDF_FILE_ROOT + " Please create it.");
        }
        out = new File (FishLinkPaths.RDF_FILE_ROOT + doi + ".rdf");
        FileWriter writer = new FileWriter(out);
                //"RDF/XML", "RDF/XML-ABBREV", "N-TRIPLE", "TURTLE", (and "TTL") and "N3"
        m.write(writer, "RDF/XML");
        System.out.println("Done writing rdf file to "+ out.getAbsolutePath());
    }

}
