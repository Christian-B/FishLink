package org.freshwaterlife.fishlink;

/**
 * Single collection point for all String Constants Other than Paths.
 * 
 * Also has methods based on Constants and collections of Constants.
 * @author Christian
 */
public class FishLinkConstants {

    /**
     * Base URL to be used in the RDF 
     */
    public static final String RDF_BASE_URL = "http://rdf.freshwaterlife.org/";

    /**
     * Name of the Sheet that holds the vocabulary or dropDown Lists
     */
    public static final String LIST_SHEET = "Lists";

    /**
     * Name of the sheet in the MetaMaster that holds the dropdown creation rules.
     */
    public static final String DROP_DOWN_SHEET = "DropDowns";

    //Column A texts
    /**
     * Value expected in Column "A" of the Row that hold the category information.
     */
    public static final String CATEGORY_LABEL = "Category";
    /**
     * Value expected in Column "A" of the Row that hold the field information.
     */
    public static final String FIELD_LABEL = "Field";
    /**
     * Value expected in Column "A" of the Row that hold the id or value link information.
     */
    public static final String ID_VALUE_LABEL = "Id/Value Column"; //Id/Value Column (or "All")  //Old ones may have row
    /**
     * Value expected in Column "A" of the Row that hold the External Sheet information.
     */
    public static final String EXTERNAL_LABEL = "external sheet";
    /**
     * Value expected in Column "A" of the Row that hold the zero vs nulls information.
     */
    public static final String ZEROS_VS_NULLS_LABEL = "Zeros vs Nulls";
    /**
     * Part of the value expected in Column "A" of the Row that divided the standard information from the optional constants.
     */
    public static final String CONSTANTS_DIVIDER = "Constants"; //-- Constants --
    /**
     * Value expected in Column "A" of the Row that hold the original headers.
     */
    public static final String HEADER_LABEL = "Header";
    /**
     * Value expected in Column "A" of the Row that hold the Sub|Type information.
     */
    private static final String SUB_TYPE_LABEL = "SubType";

    //categeries
    /**
     * Holds the vocabulary name for the category Location
     */
    private static final String LOCATION_LABEL = "Location";
    /**
     * Holds the vocabulary name for the category Observation
     */
    public static final String OBSERVATION_LABEL = "Observation";
    /**
     * Holds the vocabulary name for the category Entity
     */
    private static final String ENTITY_LABEL = "Entity";
    /**
     * Holds the vocabulary name for the category Site
     */
    private static final String SITE_LABEL = "Site";
    /**
     * Holds the vocabulary name for the category Survey
     */
    private static final String SURVEY_LABEL = "Survey";

    //fields
    /**
     * Holds the vocabulary name for the field
     */
    private static final String CONTIBUTOR_LABEL = "Contributor";
    /**
     * Holds the vocabulary name for the field Owner
     */
    private static final String OWNER_LABEL = "Owner";
    /**
     * Holds the vocabulary name for the field Id
     */
    public static final String ID_LABEL = "Id"; 
    /**
     * Holds the vocabulary name for the field Value
     */
    public static final String VALUE_LABEL = "Value"; 

    /**
     * Holds the vocabulary name for the label that says id column will be found automatically 
     */
    public static final String AUTOMATIC_LABEL = "automatic";
    
    /**
     * Holds the vocabulary name for the label that says Observation detail will be applied to all Observation values 
     */
    public static final String ALL_LABEL = "all";
    
    //Zeros vs Nulls values
    /**
     * Constant text value to go with {@link ZeroNullType#KEEP}
     */
    public static final String KEEP_LABEL = "Keep as is";
   /**
     * Constant text value to go with {@link ZeroNullType#ZEROS_AS_NULLS}
     */
    public static final String ZERO_AS_NULLS_LABEL = "Zeros as Nulls";
   /**
     * Constant text value to go with {@link ZeroNullType#NULLS_AS_ZERO}
     */
    public static final String NULLS_AS_ZERO_LABEL = "Nulls as Zeros";
     
    /**
     * Determines which Categories will Automatically be assumed to be Related if found on the same row.
     * 
     * If there is a Site and a Location on the same row the assumption is that site is based at that location.
     * Similar if there is a Site than every Observation in that row will be assumed to have been taken at that site.
     * @param category the Subject Category
     * @return the assumed Object Category or Null. Site returns Location, Observation returns Site. 
     *     Anything else returns null. 
     */
    public static String autoRelatedCategory(String category){
        if (category.equalsIgnoreCase(SITE_LABEL)){
            return LOCATION_LABEL;
        }
        if (category.equalsIgnoreCase(OBSERVATION_LABEL)){
            return SITE_LABEL;
        }
        return null;
    }

    /**
     * This matches Fields (predicates) to the Category of the Object it points to.
     *
     * Cases where the Field name matches the Category such as vocab:hasSite pointing to Site
     *      are assumed to have been caught by calling class.
     * Matches are:
     * <ul>
     *     <li> Contributor := Entity
     *     <li> Owner := Entity
     *     <li> SubSite := Entity
     * </ul>
     * @param field Fields as in annotation sheet (Without the vocab:has)
     * @return Match or Null.
     */
    public static String refersToCategory(String field){
        //Caller Used NameChecker.isCategory to find ones where the field name = a categeroy name
        if (field.equalsIgnoreCase(CONTIBUTOR_LABEL)){
            return ENTITY_LABEL;
        }
        if (field.equalsIgnoreCase(OWNER_LABEL)){
            return ENTITY_LABEL;
        }
        if (field.toLowerCase().contains("subsite")){
            return SITE_LABEL;
        }
        return null;
    }

    /**
     * Identifies a field has having to be represented with the predicate rdf:type
     *
     * @param field Fields as in annotation sheet (Without the vocab:has)
     * @return True if and only if field == subStite, otherwise false.
     */
    public static boolean isRdfTypeField(String field){
        return SUB_TYPE_LABEL.equalsIgnoreCase(field);
    }
}
