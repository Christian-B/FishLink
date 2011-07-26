package org.freshwaterlife.fishlink;

import org.freshwaterlife.fishlink.xlwrap.Constants;

/**
 * Enum for the different types of actions that can be done on Zeros and Nulls.
 * 
 * @author Christian
 */
public enum ZeroNullType {
    /**
     * (Default) indicates that Nulls and Zeros should be kept as is.
     * All values including zero and nulls remain unchanged. 
     * <p>Uses the Text from {@link Constants#KEEP_LABEL}
     */
    KEEP(Constants.KEEP_LABEL), 
    /**
     * Indicates that any zero values found should be considered null. 
     * All none zero values (including nulls) remain as is.
     * <p>Uses the Text from {@link Constants#ZERO_AS_NULLS_LABEL}
     */
    ZEROS_AS_NULLS (Constants.ZERO_AS_NULLS_LABEL), 
    /**
     * Indicates that any null values found should be considered zero. 
     * All none null values (including zero) remain as is.
     * <p>Uses the Text from {@link Constants#NULLS_AS_ZERO_LABEL}
     */
    NULLS_AS_ZERO(Constants.NULLS_AS_ZERO_LABEL);
        
    String text;
    
    private ZeroNullType(String label){
        text = label;
    }
    
    /**
     * Returns the String representation of this Object based on the Values found in Constants
     * @return Matching Constants
     */
    @Override
    public String toString(){
        return text;
    }
    
    /**
     * Converts text into a ZeroNullType.
     * 
     * Null and empty Text will return the default KEEP.
     * Based on the matching Strings found in the Constant class.
     * 
     * @param label Text to be converted into a  ZeroNullType.
     * @return a ZeroNullType with KEEp being the default.
     * @throws FishLinkException It a none null, none empty text does not match even ignoring case.
     */
    public static ZeroNullType parse(String label) throws FishLinkException{
        if (label == null){
            return KEEP;
        }
        if (label.isEmpty()){
            return KEEP;
        }
        for (ZeroNullType zeroNullType : ZeroNullType.values()){
            if (label.equalsIgnoreCase(zeroNullType.text)){
                return zeroNullType;
            }
        }
        throw new FishLinkException("Unexpected ZeroVSNull type " + label);
    }
  
}
