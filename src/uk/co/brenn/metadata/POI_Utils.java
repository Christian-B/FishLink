/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.metadata;

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
        if (index != 26){
           char first = (char)( index + 65);
           reply = first + "";
        } else {
           char last = (char)(index - ((index / 26) * 26) + 65);
           String rest = indexToAlpha((index / 26));
           reply = rest + last;
        }
        if (index != alphaToIndex(reply)){
            System.out.println("Error converting " + index);
            System.out.println("Answer was "+ reply);
            System.out.println("Which comes out as "+ alphaToIndex(reply));
            int error = 1/0;
        }
        return reply;
    }

}
