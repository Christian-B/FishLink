package uk.co.brenn.xlwrap;

/**
 *
 * @author christian
 */
public class Constants {

    //categeries
    private static final String LOCATION_LABEL = "Location";
    public static final String OBSERVATION_LABEL = "Observation";
    private static final String PERSON_LABEL = "Person";
    private static final String SITE_LABEL = "Site";
    private static final String SURVEY_LABEL = "Survey";

    //fields
    private static final String CONTIBUTOR_LABEL = "Contributor";
    private static final String OWNER_LABEL = "Owner";

    public static final String ALL_LABEL = "all";

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
            return PERSON_LABEL;
        }
        if (field.equalsIgnoreCase(OWNER_LABEL)){
            return PERSON_LABEL;
        }
        if (field.toLowerCase().contains("subsite")){
            return SITE_LABEL;
        }
        return null;
    }
}
