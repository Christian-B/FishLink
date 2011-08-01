package org.freshwaterlife.fishlink.xlwrap;

import org.freshwaterlife.fishlink.FishLinkException;
import at.jku.xlwrap.spreadsheet.Sheet;
import java.io.BufferedWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import org.freshwaterlife.fishlink.FishLinkConstants;
import org.freshwaterlife.fishlink.FishLinkUtils;
import org.freshwaterlife.fishlink.ZeroNullType;

/**
 * Methods for Writing the part of the mapping file for s single column. 
 *
 * @author Christian
 */
public class SheetWrite extends AbstractSheet{

    /**
     * The number of this sheet in its workbook.
     * Sheet number is required in the mapping file
     */
    private int sheetNumber;
    /**
     * The URI to the workbook which holds this this sheet.
     * DataPath is required in the mapping file
     */
    private String dataPath;
    /**
     * The pid to the workbook which holds this this sheet.
     * The pid is used to create workbook specific URIs.
     */
    private String pid;

    /**
     * Stores Link to the NameChecker for later use.
     */
    private NameChecker masterNameChecker;

   /** 
     * Maps the Categories of which there can only be one instance per row to the column that hold the Id.
     * "row" is used if no Id column found.
     */
    private HashMap<String,String> idValueLinks;
    /**
     * Maps the Categories the URI for that category.
     */
    private HashMap<String,String> categoryUris;
    /**
     * Keeps track of the columns that have data which should be applied to all Observations.
     */
    private ArrayList<String> allColumns;

   /**
     * Constructor based on the Sheet provided.
     * 
     * Minimal checking is done by superClass.
     * 
     * @param nameChecker Reference to the NameChecker to use as required.
     * @param annotatedSheet Sheet which will (later) be added to the Mapping File
     * @param sheetNumber Number of this Sheet within the workbook.
     * @param uri URI to the workbook this sheet is part of. 
     * @param pid PID to be used in URIs
     * @throws FishLinkException Thrown if sheet is corrupt or badly formatted.
     */
    SheetWrite (NameChecker nameChecker, Sheet annotatedSheet, int sheetNumber, String url, String pid)
            throws FishLinkException{
        super(annotatedSheet);
        this.pid = pid;
        this.sheetNumber = sheetNumber;
        this.dataPath = url;
        this.masterNameChecker = nameChecker;
        idValueLinks = new HashMap<String,String>();
        categoryUris = new HashMap<String,String>();
        allColumns = new ArrayList<String>();
    }

    /**
     * Unified name for the template
     * @return Unified name for the template
     */
    private String template(){
        return sheet.getName() + "template";
    }

    /**
     * Writes the mapping for this Sheet (see XLWrap documentation)
     * @param writer
     * @throws FishLinkException 
     */
   void writeMapping(BufferedWriter writer) throws FishLinkException{
        try {
            writer.write("	xl:offline \"false\"^^xsd:boolean ;");
            writer.newLine();
            writer.write("	xl:template [");
            writer.newLine();
            writer.write("		xl:fileName \"" + dataPath + "\" ;");
            writer.newLine();
            writer.write("		xl:sheetNumber \""+ sheetNumber +"\" ;");
            writer.newLine();
            writer.write("		xl:templateGraph :");
            writer.write(template());
            writer.write(" ;");
            writer.newLine();
            writer.write("		xl:transform [");
            writer.newLine();
            writer.write("			a rdf:Seq ;");
            writer.newLine();
            writer.write("			rdf:_1 [");
            writer.newLine();
            writer.write("				a xl:RowShift ;");
            writer.newLine();
            writer.write("				xl:restriction \"B" + firstData + ":" + lastDataColumn + firstData + "\" ;");
            writer.newLine();
            writer.write("				xl:steps \"1\" ;");
            writer.newLine();
            writer.write("			] ;");
            writer.newLine();
            writer.write("		]");
            writer.newLine();
            writer.write("	] ;");
            writer.newLine();
        } catch (IOException ex) {
            throw new FishLinkException("Unable to write mapping", ex);
        }
    }

   /**
    * Determines if the String contains an link to another id or value column.
    * @param idValueLink Possible link
    * @return False If and only if it can be determined that the link is null, blank or a known ignore value.
    *       Otherwise True
    */
    private boolean hasIdValueLink(String idValueLink){
        if (idValueLink == null) return false;
        if (idValueLink.isEmpty()) return false;
        if (idValueLink.equalsIgnoreCase("n/a")) return false;
        if (idValueLink.equalsIgnoreCase(FishLinkConstants.AUTOMATIC_LABEL)) return false;
        return true;
    }
    
    /**
     * Write the URI for the Subject of a Tripe.
     * 
     * Main job of this method is to distinguish between different types of 
     *    data and then call the correct WriteURIForXXX method.
     * <p>
     * Result will be a XLwarp function for creating the applicable URI.
     * 
     * @param writer
     * @param category category for the column being written
     * @param field field for column being written 
     * @param idValueLink Possible link to the idValueLink
     * @param dataColumn Column currently being written
     * @param dataZeroNull ZeroNull Setting or column being written
     * @throws FishLinkException 
     */
    private void writeUriForSubject(BufferedWriter writer, String category, String field, String idValueLink,
            String dataColumn, ZeroNullType dataZeroNull) throws FishLinkException {
        if(category.equalsIgnoreCase(FishLinkConstants.OBSERVATION_LABEL)) { 
            if (field.equalsIgnoreCase(FishLinkConstants.VALUE_LABEL)) {
                writeUriForObservationValue(writer, idValueLink, dataColumn, dataZeroNull);
            } else {
                writeUriForObservationNonValue(writer, idValueLink, dataColumn, dataZeroNull);
            }
        } else { 
            if (hasIdValueLink(idValueLink)){
                throw new FishLinkException(sheet.getSheetInfo() + " Data Column " + dataColumn + " with category " +
                        category + " and field " + field + " is not allowed to have an id/Value Link");
            }
            if (field.toLowerCase().equals(FishLinkConstants.ID_LABEL)){
                writeUriForId(writer, category, dataColumn, dataZeroNull);
            } else {
                writeUriForOther(writer, category, dataColumn, dataZeroNull);
            }
        }
    }
    
    /**
     * Writes a URI based on the CELL in which the Observation value was saved.
     * 
     * @param writer
     * @param idValueLink Should be empty
     * @param dataColumn Column currently being written
     * @param dataZeroNull ZeroNull Setting or column being written
     * @throws FishLinkException 
     */
    private void writeUriForObservationValue(BufferedWriter writer, String idValueLink, String dataColumn, 
            ZeroNullType dataZeroNull) throws FishLinkException {
        String uri = categoryUris.get(FishLinkConstants.OBSERVATION_LABEL);
        if (hasIdValueLink(idValueLink)){
            throw new FishLinkException(sheet.getSheetInfo() + " Data Column " + dataColumn + " with category " +
                    FishLinkConstants.OBSERVATION_LABEL + " and field " + FishLinkConstants.VALUE_LABEL + 
                    " is not allowed to have an id/Value Link");
        }
        try {
            writer.write("[ xl:uri \"CELL_URI('" + uri + "', " + dataColumn + firstData + ",'"
                    + dataZeroNull + "')\"^^xl:Expr ] ");
        } catch (IOException ex) {
            throw new FishLinkException("Unable to write uri", ex);
        }                
        return;
    }
    
    /**
     * Writes a URI based on the CELL in which the Observation's value was saved.
     * 
     * The URI written will be exactly the same as the one for the Value Column.
     * This will allow the underlying model to combine the triples into a single subject.
     * <p>
     * The Difference between the function written here and the one write for the value column 
     *    is that it will also check that the data cell (this column) is not null too.
     * 
     * @param writer
     * @param idValueLink Link to the column that holds the Observation's Value.
     * @param dataColumn Column currently being written
     * @param dataZeroNull ZeroNull Setting or column being written
     * @throws FishLinkException 
     */
    private void writeUriForObservationNonValue(BufferedWriter writer, String idValueLink, 
            String dataColumn, ZeroNullType dataZeroNull) throws FishLinkException {
        String uri = categoryUris.get(FishLinkConstants.OBSERVATION_LABEL);
        if (!hasIdValueLink(idValueLink)){
            throw new FishLinkException(sheet.getSheetInfo() + " Data Column " + dataColumn + " with category " +
                    FishLinkConstants.OBSERVATION_LABEL + " and field " + FishLinkConstants.VALUE_LABEL + 
                    " needs an id/Value Link");
        }
        String idNullZeroString = getCellValue (idValueLink, idValueLinkRow);
        ZeroNullType idZeroNull = ZeroNullType.parse(idNullZeroString);
        try {
            writer.write("[ xl:uri \"CELL_URI('" + uri + "', " + idValueLink + firstData + ", '" + idZeroNull + "' ,"
                + dataColumn + firstData + ",'" + dataZeroNull + "')\"^^xl:Expr ] ");
        } catch (IOException ex) {
            throw new FishLinkException("Unable to write uri", ex);
        }
    }

    /**
     * Writes an URI based on the ID values found in this column.
     * 
     * @param writer
     * @param category category for the column being written
     * @param dataColumn Column currently being written
     * @param zeroNull ZeroNull Setting or column being written
     * @throws FishLinkException 
     */
     private void writeUriForId(BufferedWriter writer, String category, String dataColumn, ZeroNullType zeroNull) 
            throws FishLinkException {
        String uri = categoryUris.get(category);
        try {
            writer.write("[ xl:uri \"ID_URI('" + uri + "', " + dataColumn + firstData + ",'"
                    + zeroNull + "')\"^^xl:Expr ] ");
        } catch (IOException ex) {
            throw new FishLinkException("Unable to write uri", ex);
        }
    }

    /**
     * Writes an URI based on the ID values for the given category.
     * 
     * if this category has an Id column the URI written will be exactly the same as the one for the Id Column.
     * This will allow the underlying model to combine the triples into a single subject.
     * <p>
     * The Difference between the function written here and the one write for the Id column 
     *    is that it will also check that the data cell (this column) is not null too.
     * <p>
     * If there is no Id Column for s given category a row based URI will be generated.
      * 
     * @param writer
     * @param category category for the column being written
     * @param dataColumn Column currently being written
     * @param zeroNull ZeroNull Setting or column being written
     * @throws FishLinkException 
     */
    private void writeUriForOther(BufferedWriter writer, String category, String dataColumn, 
            ZeroNullType dataZeroVsNulls) throws FishLinkException {
        String uri = categoryUris.get(category);
        //"row" is added for automatic ones where no id column found.
        String idValueLink = idValueLinks.get(category);
        if (idValueLink.equalsIgnoreCase("row")){
            try {
                writer.write("[ xl:uri \"ROW_URI('" + uri + "', " + dataColumn + firstData + ",'" +
                        dataZeroVsNulls + "')\"^^xl:Expr ] ");
            } catch (IOException ex) {
                throw new FishLinkException("Unable to write uri", ex);
            }
        } else {
            String zeroNullString = getCellValue (idValueLink, ZeroNullRow);
            ZeroNullType idZeroNull = ZeroNullType.parse(zeroNullString);
            try {
                writer.write("[ xl:uri \"ID_URI('" + uri + "', " + idValueLink + firstData + ",'" + idZeroNull + "'," + 
                        dataColumn + firstData + ",'" + dataZeroVsNulls + "')\"^^xl:Expr ] ");
            } catch (IOException ex) {
                throw new FishLinkException("Unable to write uri", ex);
            }
        }
    }

    /**
     * Writes the predicate of a Triple.
     * 
     * All predicates that do not start with "is" or "has" will have "has" added.
     * 
     * @param writer
     * @param vocab Vocabulary to be used
     * @throws FishLinkException 
     */
    private void writeVocab (BufferedWriter writer, String vocab) throws FishLinkException {
        try {
            if (vocab.startsWith("is") || vocab.startsWith("has")){
                writer.write("	vocab:" + vocab + "\t");
            } else {
                writer.write("	vocab:has" + vocab + "\t");
            }
        } catch (IOException ex) {
            throw new FishLinkException("Unable to write vocab", ex);
        }            
    }

    /**
     * Writes both the Predicate and Object for triples with the predicate rdf:type.
     * @param writer
     * @param type The Object or type being asserted.
     * @throws FishLinkException 
     */
    private void writeRdfType (BufferedWriter writer, String type) throws FishLinkException {
        try {
           writer.write ("	rdf:type");
           writeType(writer, type);
        } catch (IOException ex) {
            throw new FishLinkException("Unable to write vocab", ex);
        }            
    }
 
    /**
     * Writes a type adding the prefix type:
     * @param writer
     * @param type The Object or type being asserted.
     * @throws FishLinkException 
     */
    private void writeType (BufferedWriter writer, String type) throws FishLinkException {
        try {
           writer.write (" type:" + type);
           writer.write (" ;");
           writer.newLine();
        } catch (IOException ex) {
            throw new FishLinkException("Unable to write vocab", ex);
        }            
    }

   /**
     * Write the data from one of the Constant rows in the Annotated sheet such that it will be repeated for each row.
     * @param writer
     * @param category category for the dataColumn being written
     * @param dataColumn Column currently being written
     * @param row Row in the Annotation Sheet the Constant is found in.
     * @throws FishLinkException 
     */
    private void writeConstant (BufferedWriter writer, String category, String dataColumn, int row) 
            throws FishLinkException{
        String field = getCellValue ("A", row);
        if (field == null){ //Should never happen
            return;
        }
        String value = getCellValue (dataColumn, row);
        if (value == null || value.isEmpty() || value.equalsIgnoreCase("n/a")){
            return;
        }
        if (FishLinkConstants.isRdfTypeField(field)){
            masterNameChecker.checkSubType(sheet.getSheetInfo(), category, value);
            writeRdfType(writer, value);
        } else {
            masterNameChecker.checkConstant(sheet.getSheetInfo(), field, value);
            writeVocab (writer, field);
            if (!value.startsWith("'")){
                value = "'" + value ;
            }
            if (!value.endsWith("'")){
                value = value + "'";
            }
            try {
                writer.write("[ xl:uri \"'" + FishLinkConstants.RDF_BASE_URL + "constant/' & URLENCODE(" + value + ")\"^^xl:Expr ] ;");
                writer.newLine();
            } catch (IOException ex) {
                throw new FishLinkException("Unable to write constant", ex);
            }            
        }
    }

    /**
     * Gets the ZeroNull type for this dataColumn or the default ZeroNullType
     * @param dataColumn
     * @return The Type specified or the default type if none specified
     * @throws FishLinkException If the String value is unknown.
     */
    private ZeroNullType getZeroNullType (String dataColumn) throws FishLinkException {
        if (ZeroNullRow >= 0){
            String ignoreZeroString = getCellValue (dataColumn, ZeroNullRow);
            return ZeroNullType.parse(ignoreZeroString);
        } else{
            //Get the default
            return ZeroNullType.parse(null);
        }
    }

    /**
     * Determines the type of data/Object that a particular field in a particular column points to.
     * @param subjectCategory Category of the Triple's subject 
     * @param field Vocabulary used to create the Triple's Predicate
     * @return Type of the Triple's Object
     * @throws FishLinkException 
     */
    private String refersToCategory(String subjectCategory, String field) throws FishLinkException{
       if ( masterNameChecker.isCategory(field)) {
           return field;
       }
       if (FishLinkConstants.refersToSameCategery(field)){
           return subjectCategory;
       }
       return FishLinkConstants.refersToCategory(field);
    }

    /**
     * Writes the Predicate and Object of a Triple
     * @param writer
     * @param subjectCategory Category of the Triple's subject 
     * @param dataColumn Column currently being written
     * @param zeroNull Value for the dataColumn
     * @throws FishLinkException 
     */
    private void writeData(BufferedWriter writer, String subjectCategory, String dataColumn, ZeroNullType zeroNull)
            throws FishLinkException {
        //Write the Predicate
        String field = getCellValue (dataColumn, fieldRow);
        writeVocab(writer, field);
        //Write the Object
        //Check if it is expected to be a link to another object
        String category = refersToCategory(subjectCategory, field);
        try {
            if (category  == null) {
                //Write raw data with Zero Null conversion as applicable
                switch (zeroNull){
                    case KEEP:
                        writer.write("\"" + dataColumn + firstData + "\"^^xl:Expr ;");
                        break;
                    case ZEROS_AS_NULLS:   
                        writer.write("\"ZERO_AS_NULL(" + dataColumn + firstData + ")\"^^xl:Expr ;");
                        break;
                    case NULLS_AS_ZERO:   
                        writer.write("\"NULL_AS_ZERO(" + dataColumn + firstData + ")\"^^xl:Expr ;");
                        break;
                    default:
                        throw new FishLinkException("Unexpected ZeroNullType " + zeroNull);
                }
            } else {
                //Write a link to another column
                String uri = getUri(category, dataColumn);
                writer.write("[ xl:uri \"ID_URI('" + uri + "'," + dataColumn + firstData + ", '" + zeroNull + 
                        "')\"^^xl:Expr ];");
            }
            writer.newLine();
            }  catch (IOException ex) {
                throw new FishLinkException("Unable to write data", ex);
            }            
    }

    /**
     * Writes the auto related columns to this dataColumn.
     * 
     * This links a category to another category that are assume to be automatically related.
     * For example if these is a Site and a Location the Site is assumed to be at that location.
     * 
     * @param writer
     * @param category category for the dataColumn being written
     * @param dataColumn Column currently being written
     * @param dataZeroNull ZeroNull Setting for column being written
     * @throws FishLinkException 
     */
    private void writeAutoRelated(BufferedWriter writer, String category, String dataColumn, ZeroNullType dataZeroNull)
            throws FishLinkException {
        String related = FishLinkConstants.autoRelatedCategory(category);
        if (related == null){
            return; //No autorelated category
        }
        String uri = categoryUris.get(related);
        if (uri == null){
            return; //No data for that category
        }
        writeVocab(writer, related);
        writeUriForOther(writer, related, dataColumn, dataZeroNull);
        try {
            writer.write(";");
            writer.newLine();
        }  catch (IOException ex) {
            throw new FishLinkException("Unable to write ", ex);
        }            
    }

    /**
     * Applies the data assigned to all Observations to this column.
     * 
     * Only applies to Observation.Value so does nothing in any other case.
     * 
     * @param writer
     * @param category category for the dataColumn being written
     * @param field field for the dataColumn being written
     * @throws FishLinkException 
     */
    private void writeAllRelated(BufferedWriter writer, String category, String field)
            throws FishLinkException {
        if (!category.equalsIgnoreCase(FishLinkConstants.OBSERVATION_LABEL)){
            return;
        }
       if (!field.equalsIgnoreCase(FishLinkConstants.VALUE_LABEL)){
            return;
        }
        for (String allColumn : allColumns){
            String allZeroNullString = getCellValue (allColumn, ZeroNullRow);
            ZeroNullType allZeroNull = ZeroNullType.parse(allZeroNullString);
            writeData(writer, category, allColumn, allZeroNull);
        }
    }

    /**
     * This function checks if a dataColumn should be written and then calls function to write it. 
     * @param writer
     * @param dataColumn Column currently being written
     * @return True if and only if something was actually written.
     * @throws FishLinkException 
     */
    private boolean writeTemplateColumn(BufferedWriter writer, String dataColumn) throws FishLinkException{
        String category = getCellValue (dataColumn, categoryRow);
        //ystem.out.println(category + " " + dataColumn + " " + categoryRow);
        String field = getCellValue (dataColumn, fieldRow);
        String idValueLink = getCellValue (dataColumn, idValueLinkRow);
        String external = getExternal(dataColumn);
        if (category == null || category.toLowerCase().equals("undefined")) {
            FishLinkUtils.report("Skippig column " + dataColumn + " as no Category provided");
            return false;
        }
        if (field == null){
            FishLinkUtils.report("Skippig column " + dataColumn + " as no Feild provided");
            return false;
        }
        field = masterNameChecker.checkField(sheet.getSheetInfo(), category, field);
        if (field.equalsIgnoreCase("id") && !external.isEmpty()){
            FishLinkUtils.report("Skipping column " + dataColumn + " as it is an external id");
            return false;
        }
        if (idValueLink != null && idValueLink.equals(FishLinkConstants.ALL_LABEL)){
            FishLinkUtils.report("Skipping column " + dataColumn + " as it is an all column.");
            return false;
        }       
        writeTemplateColumn(writer, category, field, idValueLink, external,dataColumn);
        return true;
    }

    /**
     * This function writes a dataColumn, including its constants, automatic and all linked in columns, 
     * and the alternative rdfTypes.
     * 
     * @param writer
     * @param category category for the column being written
     * @param field field for column being written 
     * @param idValueLink Possible link to the idValueLink
     * @param dataColumn Column currently being written
     * @param external An external link to another sheet (could even be in another workbook)
     * @throws FishLinkException 
     */
    private void writeTemplateColumn(BufferedWriter writer, String category, String field, 
            String idValueLink, String external, String dataColumn) throws FishLinkException{
        ZeroNullType zeroNull = getZeroNullType(dataColumn);
        writeUriForSubject(writer, category, field, idValueLink, dataColumn, zeroNull);
        try {
            writer.write (" a ");
            writeType (writer, category);
        }  catch (IOException ex) {
            throw new FishLinkException("Unable to write a type", ex);
        }            
        writeData(writer, category, dataColumn, zeroNull);
        writeAutoRelated(writer, category, dataColumn, zeroNull);
        writeAllRelated(writer, category, field);
        for (int row = firstConstant; row <= lastConstant; row++){
            writeConstant(writer,  category, dataColumn, row);
        }

        try {
            if (external.isEmpty()){
               writeRdfType (writer, pid + "_" + sheet.getName() + "_" + category);
             }

            writer.write(".");
            writer.newLine();
            writer.newLine();
        }  catch (IOException ex) {
            throw new FishLinkException("Unable to other type ", ex);
        }            
    }

    /**
     * Gets the value of the ExtarnalSheet row if any otherwise a blank String
     * @param dataColumn Column currently being written
     * @return
     * @throws FishLinkException 
     */
    private String getExternal(String dataColumn) throws FishLinkException {
        if (externalSheetRow < 1){
            return "";
        }
        String externalFeild = getCellValue (dataColumn, externalSheetRow);
        if (externalFeild == null){
            return "";
        }
        return externalFeild;
    }

    /**
     * Gets the URI for a category with no external link.
     * 
     * @param category
     * @return 
     */
    private String getCategoryUri(String category){
        return  FishLinkConstants.RDF_BASE_URL + "resource/" + category + "_" + pid + "_" + sheet.getName() + "/";
    }

    /**
     * gets the URI for a category.
     * 
     * This could be an external Link to a URI on a different Sheet (even different workbook)
     * 
     * @param category category for the column being written
     * @param dataColumn Column currently being written
     * @return
     * @throws FishLinkException 
     */
    private String getUri(String category, String dataColumn) throws FishLinkException {
        String externalField = getExternal(dataColumn);
        if (externalField.isEmpty()){
            return getCategoryUri(category);
        }
        //External link
        String externalPid;
        String externalSheet;
        if (externalField.startsWith("[")){
            externalPid = externalField.substring(1, externalField.indexOf(']'));
            externalSheet = externalField.substring( externalField.indexOf(']')+1);
        } else {
            externalPid = pid;
            externalSheet = externalField;
        }
        return  FishLinkConstants.RDF_BASE_URL + "resource/" + category + "_" + externalPid + "_" + externalSheet + "/";
    }

    /**
     * Completes the category level information.
     * @param category category for the dataColumn being written
     * @param dataColumn Column currently being written
     * @throws FishLinkException 
     */
    private void findCategoryLevelDataForColumn(String category, String dataColumn) throws FishLinkException{
        String field = getCellValue (dataColumn, fieldRow);
        String id = idValueLinks.get(category);
        if (field.equalsIgnoreCase("id")){
            if (id == null || id.equalsIgnoreCase("row")){
                idValueLinks.put(category, dataColumn);
                categoryUris.put(category, getUri(category,  dataColumn));
            } else {
                throw new FishLinkException("Found two different id columns of type " + category);
            }
        } else {
            if (id == null){
                idValueLinks.put(category, "row");
            }//else leave the dataColumn or "row" already there.
            if (categoryUris.get(category) == null){
                categoryUris.put(category, getCategoryUri(category));
            }
        }
    }

    /**
     * Looks for any All columns and records them for later use.
     * 
     * @param category category for the dataColumn being written
     * @param dataColumn Column currently being written
     * @throws FishLinkException 
     */
    private void findAllColumn(String category, String dataColumn) throws FishLinkException{
        String idColumn = getCellValue (dataColumn, idValueLinkRow);
        if (idColumn == null || !idColumn.equalsIgnoreCase("all")){
            return;
        }
        if (category.equalsIgnoreCase(FishLinkConstants.OBSERVATION_LABEL)){
            String field = getCellValue (dataColumn, fieldRow);
            if (field == null || field.isEmpty()){
                throw new FishLinkException ("All id.Value Column " + dataColumn + " missing a field value");
            }
            allColumns.add(dataColumn);
        } else {
            throw new FishLinkException ("All id.Value Column only supported for Categeroy " +
                    FishLinkConstants.OBSERVATION_LABEL);
        }
    }

    /**
     * Records the information for each used Category and finds the All columns
     * @throws FishLinkException 
     */
    private void recordCategoryLevelDataAndFindAllColumns() throws FishLinkException{
        int maxColumn = FishLinkUtils.alphaToIndex(lastDataColumn);
        for (int i = 1; i < maxColumn; i++){
            String dataColumn = FishLinkUtils.indexToAlpha(i);
            String category = getCellValue (dataColumn, categoryRow);
            if (category == null || category.isEmpty()){
                //do nothing
            } else {
                findCategoryLevelDataForColumn(category, dataColumn);
                findAllColumn(category, dataColumn);
            }
        }
     }

    /**
     * Writes the Template part of the mapping file
     * @param writer
     * @throws FishLinkException Any found exception possibly wrapped
     */
    void writeTemplate(BufferedWriter writer) throws FishLinkException{
        FishLinkUtils.report("Writing template for "+sheet.getSheetInfo());
        recordCategoryLevelDataAndFindAllColumns();
        try {
            writer.write(":");
            writer.write(template());
            writer.write(" {");
            writer.newLine();
        }  catch (IOException ex) {
            throw new FishLinkException("Unable to write template ", ex);
        }            
        int maxColumn = FishLinkUtils.alphaToIndex(lastDataColumn);
        boolean foundColumn = false;
        //Start at 1 to ignore label column;
        for (int i = 1; i < maxColumn; i++){
            String column = FishLinkUtils.indexToAlpha(i);
            if (writeTemplateColumn(writer, column)){
                foundColumn = true;
            }
        }
        if (!foundColumn){
            throw new FishLinkException("No mappable columns found in sheet " +  sheet.getSheetInfo());
        }
        try {
            writer.write("}");
            writer.newLine();
        }  catch (IOException ex) {
            throw new FishLinkException("Unable to write template end ", ex);
        }            
    }

}


