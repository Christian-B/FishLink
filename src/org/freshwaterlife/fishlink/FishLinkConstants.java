package org.freshwaterlife.fishlink;

/**
 *
 * @author Christian
 */
public class FishLinkConstants {

    /**
     * Base URL to be used in the RDF 
     */
    public static final String RDF_BASE_URL = "http://rdf.freshwaterlife.org/";

    public static final String LIST_SHEET = "Lists";

    public static final String DROP_DOWN_SHEET = "DropDowns";

    //Column A texts
    public static final String CATEGORY_LABEL = "Category";
    public static final String FIELD_LABEL = "Field";
    public static final String ID_VALUE_LABEL = "Id/Value Column"; //Id/Value Column (or "All")  //Old ones may have row
    public static final String EXTERNAL_LABEL = "external sheet";
    public static final String ZEROS_VS_NULLS_LABEL = "Zeros vs Nulls";
    public static final String CONSTANTS_DIVIDER = "Constants"; //-- FishLinkConstants --
    public static final String HEADER_LABEL = "Header";
    //categeries
    private static final String LOCATION_LABEL = "Location";
    public static final String OBSERVATION_LABEL = "Observation";
    private static final String ENTITY_LABEL = "Entity";
    private static final String SITE_LABEL = "Site";
    private static final String SURVEY_LABEL = "Survey";

    //fields
    private static final String CONTIBUTOR_LABEL = "Contributor";
    private static final String OWNER_LABEL = "Owner";
    public static final String ID_LABEL = "Id"; 
    public static final String VALUE_LABEL = "Value"; 

    public static final String AUTOMATIC_LABEL = "automatic";
    public static final String ALL_LABEL = "all";
    private static final String SUB_TYPE_LABEL = "SubType";

    //Zeros vs Nulls values
    public static final String KEEP_LABEL = "Keep as is";
    public static final String ZERO_AS_NULLS_LABEL = "Zeros as Nulls";
    public static final String NULLS_AS_ZERO_LABEL = "Nulls as Zeros";
     
    public static String autoRelatedCategory(String category){
        if (category.equalsIgnoreCase(SITE_LABEL)){
            return LOCATION_LABEL;
        }
        if (category.equalsIgnoreCase(OBSERVATION_LABEL)){
            return SITE_LABEL;
        }
        return null;
    }

    public static String refersToCategory(String field){
        //Use NameChecker.isCategory to find ones where the field name = a categeroy name
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
    
    public static boolean isRdfTypeField(String field){
        return SUB_TYPE_LABEL.equalsIgnoreCase(field);
    }
}
