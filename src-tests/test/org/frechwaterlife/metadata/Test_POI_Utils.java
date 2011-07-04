
package test.org.frechwaterlife.metadata;

import junit.framework.TestCase;
import org.freshwaterlife.fishlink.FishLinkUtils;
import org.junit.Test;

/**
 *
 * @author Christian
 */
public class Test_POI_Utils extends TestCase{

    
    	@Test
	public void testAlphaToIndexAndBack() {
            for (int i = 0; i< 1000; i++){
                String alpha = FishLinkUtils.indexToAlpha(i);
                int col = FishLinkUtils.alphaToIndex(alpha);
		assertEquals("AlphaToIndex and back failure",i, col);
            }
	}

    	@Test
	public void testSpecialCases() {
           assertEquals("0 to A ","A", FishLinkUtils.indexToAlpha(0));
           assertEquals("25 to Z ","Z", FishLinkUtils.indexToAlpha(25));
           assertEquals("26 to AA ","AA", FishLinkUtils.indexToAlpha(26));
           assertEquals("51 to AZ ","AZ", FishLinkUtils.indexToAlpha(51));
           assertEquals("52 to BA ","BA", FishLinkUtils.indexToAlpha(52));
	}
}
