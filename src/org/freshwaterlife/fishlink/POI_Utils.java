package org.freshwaterlife.fishlink;

/**
 *
 * @author Christian
 */
public class POI_Utils {

    /**
     * @param alpha index
     * @return zero-based numerical index
     */
    public static int alphaToIndex(String alpha) {
	char[] letters = alpha.toUpperCase().toCharArray();
	int index = 0;
	for (int i = 0; i < letters.length; i++)
            index += ((letters[letters.length-i-1]) - 64) * (Math.pow(26, i)); // A is 64
	return --index;
    }

    /**
     *
     * @param index zero-based numerical index
     * @return alpha index
     */
    public static String indexToAlpha(int index){
        String reply = "";
        if (index < 26){
           char first = (char)( index + 65);
           reply = first + "";
        } else {
           char last = (char)(index - ((index / 26) * 26) + 65);
           String rest = indexToAlpha((index / 26)-1);
           reply = rest + last;
        }
        if (index != alphaToIndex(reply)){
            System.err.println("Error converting " + index);
            System.err.println("Answer was "+ reply);
            System.err.println("Which comes out as "+ alphaToIndex(reply));
            throw new AssertionError("Error in indexToAlpha");
        }
        return reply;
    }

    public static void main(String[] args) {
        String test = indexToAlpha(48);
    }
}
