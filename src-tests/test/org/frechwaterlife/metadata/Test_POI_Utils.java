
package test.org.frechwaterlife.metadata;

import junit.framework.TestCase;
import org.freshwaterlife.fishlink.metadatacreator.POI_Utils;
import org.junit.Test;
import static org.junit.Assert.assertEquals;

/**
 *
 * @author Christian
 */
public class Test_POI_Utils extends TestCase{

    
    	@Test
	public void testAlphaToIndexAndBack() {
            for (int i = 0; i< 1000; i++){
                String alpha = POI_Utils.indexToAlpha(i);
                int col = POI_Utils.alphaToIndex(alpha);
		assertEquals("AlphaToIndex and back failure",i, col);
            }
	}

    	@Test
	public void testSpecialCases() {
           assertEquals("0 to A ","A", POI_Utils.indexToAlpha(0));
           assertEquals("25 to Z ","Z", POI_Utils.indexToAlpha(25));
           assertEquals("26 to AA ","AA", POI_Utils.indexToAlpha(26));
           assertEquals("51 to AZ ","AZ", POI_Utils.indexToAlpha(51));
           assertEquals("52 to BA ","BA", POI_Utils.indexToAlpha(52));
	}
}
