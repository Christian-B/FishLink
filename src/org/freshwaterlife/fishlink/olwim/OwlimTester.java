/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package org.freshwaterlife.fishlink.olwim;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.Reader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CountDownLatch;
import org.openrdf.model.Graph;
import org.openrdf.model.Namespace;
import org.openrdf.model.Resource;
import org.openrdf.model.Statement;
import org.openrdf.model.URI;
import org.openrdf.model.Value;
import org.openrdf.model.impl.GraphImpl;
import org.openrdf.model.impl.URIImpl;
import org.openrdf.model.vocabulary.RDF;
import org.openrdf.query.Binding;
import org.openrdf.query.BindingSet;
import org.openrdf.query.BooleanQuery;
import org.openrdf.query.GraphQuery;
import org.openrdf.query.GraphQueryResult;
import org.openrdf.query.Query;
import org.openrdf.query.QueryLanguage;
import org.openrdf.query.TupleQuery;
import org.openrdf.query.TupleQueryResult;
import org.openrdf.repository.Repository;
import org.openrdf.repository.RepositoryConnection;
import org.openrdf.repository.RepositoryException;
import org.openrdf.repository.RepositoryResult;
import org.openrdf.repository.config.RepositoryConfig;
import org.openrdf.repository.config.RepositoryConfigException;
import org.openrdf.repository.manager.LocalRepositoryManager;
import org.openrdf.repository.manager.RepositoryManager;
import org.openrdf.repository.sail.SailRepository;
import org.openrdf.rio.RDFFormat;
import org.openrdf.rio.RDFHandler;
import org.openrdf.rio.RDFHandlerException;
import org.openrdf.rio.RDFParseException;
import org.openrdf.rio.RDFParser;
import org.openrdf.rio.Rio;
import org.openrdf.rio.UnsupportedRDFormatException;
import org.openrdf.sail.memory.MemoryStore;

/**
 *
 * @author christian
 */
public class OwlimTester {

    private static final String QUERYFILE = "./queries/cyabTest.sparql";
    private static final String TTL_FILE = "./data/fishlink.ttl";
    private static final String REPOSITORY_ID = "fishlink";
    private static final String DATA_DIR = "./output/rdf/";


	// From repository.getConnection() - the connection through which we will
	// use the repository
    private RepositoryConnection repositoryConnection;

	// From repositoryManager.getRepository(...) - the actual repository we will
	// work with
	private Repository repository;

    // The repository manager
	private RepositoryManager repositoryManager;

    // A map of namespace-to-prefix
	private Map<String, String> namespacePrefixes = new HashMap<String, String>();

    OwlimTester() throws RepositoryConfigException, RepositoryException, Exception{

        repositoryManager = new LocalRepositoryManager(new File("."));

        // Initialise the repository manager
        repositoryManager.initialize();

    }

   private void newRepository() throws RepositoryException, RepositoryConfigException, Exception{
        repositoryManager.removeRepositoryConfig(REPOSITORY_ID);
		// The configuration file
		File configFile = new File(TTL_FILE);
		System.out.println("Using configuration file: " + configFile.getAbsolutePath());

		// Parse the configuration file, assuming it is in Turtle format
		final Graph graph = parseFile(configFile, RDFFormat.TURTLE, "http://rdf.freshwaterlife.org#");

		// Look for the subject of the first matching statement for
		// "?s type Repository"
		Iterator<Statement> iter = graph.match(null, RDF.TYPE, new URIImpl(
				"http://www.openrdf.org/config/repository#Repository"));
		Resource repositoryNode = null;
		if (iter.hasNext()) {
			Statement st = iter.next();
			repositoryNode = st.getSubject();
		}
        System.out.println(repositoryNode);
        System.out.println(repositoryNode.stringValue());
        // Create a configuration object from the configuration file and add
        // it
        // to the repositoryManager
        RepositoryConfig repositoryConfig = RepositoryConfig.create(graph, repositoryNode);
        repositoryManager.addRepositoryConfig(repositoryConfig);

        openRepository();
    }

   private void openRepository() throws RepositoryException, RepositoryConfigException{
        //repositoryManager.removeRepositoryConfig(REPOSITORY_ID);
        repository = repositoryManager.getRepository(REPOSITORY_ID);

        if (repository == null){
            int error = 1/0;
        }
        // Open a connection to this repository
		repositoryConnection = repository.getConnection();
		repositoryConnection.setAutoCommit(false);
    }

   /**
	 * Parse the given RDF file and return the contents as a Graph
	 *
	 * @param configurationFile
	 *            The file containing the RDF data
	 * @return The contents of the file as an RDF graph
	 */
	private Graph parseFile(File configurationFile, RDFFormat format, String defaultNamespace) throws Exception {
		final Graph graph = new GraphImpl();
		RDFParser parser = Rio.createParser(format);
		RDFHandler handler = new RDFHandler() {
			public void endRDF() throws RDFHandlerException {
			}

			public void handleComment(String arg0) throws RDFHandlerException {
			}

			public void handleNamespace(String arg0, String arg1) throws RDFHandlerException {
			}

			public void handleStatement(Statement statement) throws RDFHandlerException {
				graph.add(statement);
			}

			public void startRDF() throws RDFHandlerException {
			}
		};
		parser.setRDFHandler(handler);
		parser.parse(new FileReader(configurationFile), defaultNamespace);
		return graph;
	}

    // A list of RDF file formats used in loadFile().
	private static final RDFFormat allFormats[] = new RDFFormat[] { RDFFormat.NTRIPLES, RDFFormat.N3, RDFFormat.RDFXML,
			RDFFormat.TURTLE, RDFFormat.TRIG, RDFFormat.TRIX };

    private void loadFile(File file) throws IOException, RepositoryException {
        System.out.println ("Loading: " + file.getAbsolutePath());
		URI context =  new URIImpl(file.toURI().toString());
		boolean loaded = false;
      // Try all formats
		for (RDFFormat rdfFormat : allFormats) {
			Reader reader = null;
			try {
				reader = new BufferedReader(new FileReader(file), 1024 * 1024);
				repositoryConnection.add(reader, "http://example.org/owlim#", rdfFormat, context);
				repositoryConnection.commit();
				System.out.println("Loaded file '" + file.getName() + "' (" + rdfFormat.getName() + ").");
				loaded = true;
				break;
			} catch (UnsupportedRDFormatException e) {
				// Format not supported, so try the next format in the list.
			} catch (RDFParseException e) {
				// Can't parse the file, so it is probably in another format.
				// Try the next format.
			} finally {
				if (reader != null)
					reader.close();
			}
			if (!loaded)
				repositoryConnection.rollback();
		}
		if (!loaded)
			System.out.println("Failed to load '" + file.getName() + "'.");
	}

    private void loadFiles (File node) throws IOException, RepositoryException{
        if (node.isDirectory()) {
            System.out.println ("Loading children of "+ node.getAbsolutePath());
            File[] children = node.listFiles();
            for (File child : children) {
                loadFiles(child);
            }
        } else {
            loadFile(node);
        }
    }

    /**
	 * Two approaches for finding the total number of explicit statements in a
	 * repository.
	 *
	 * @return The number of explicit statements
	 */
	private long numberOfExplicitStatements() throws Exception {

		// This call should return the number of explicit statements.
		long explicitStatements = repositoryConnection.size();

		// Another approach is to get an iterator to the explicit statements
		// (by setting the includeInferred parameter to false) and then counting
		// them.
		RepositoryResult<Statement> statements = repositoryConnection.getStatements(null, null, null, false);
		explicitStatements = 0;

		while (statements.hasNext()) {
			statements.next();
			explicitStatements++;
		}
		statements.close();
		return explicitStatements;
	}

	/**
	 * A method to count only the inferred statements in the repository. No
	 * method for this is available through the Sesame API, so OWLIM uses a
	 * special context that is interpreted as instruction to retrieve only the
	 * implicit statements, i.e. not explicitly asserted in the repository.
	 *
	 * @return The number of implicit statements.
	 */
	private long numberOfImplicitStatements() throws Exception {
		// Retrieve all inferred statements
		RepositoryResult<Statement> statements = repositoryConnection.getStatements(null, null, null, true,
				new URIImpl("http://www.ontotext.com/implicit"));
		long implicitStatements = 0;

		while (statements.hasNext()) {
			statements.next();
			implicitStatements++;
		}
		statements.close();
		return implicitStatements;
	}

    /**
	 * Show some initialisation statistics
	 */
	public void showInitializationStatistics(long startupTime) throws Exception {

        long explicitStatements = numberOfExplicitStatements();
        long implicitStatements = numberOfImplicitStatements();

        System.out.println("Loaded: " + explicitStatements + " explicit statements.");
        System.out.println("Inferred: " + implicitStatements + " implicit statements.");

        if (startupTime > 0) {
            double loadSpeed = explicitStatements / (startupTime / 1000.0);
            System.out.println(" in " + startupTime + "ms.");
            System.out.println("Loading speed: " + loadSpeed + " explicit statements per second.");
        } else {
            System.out.println(" in less than 1 second.");
        }
        System.out.println("Total number of statements: " + (explicitStatements + implicitStatements));
	}

	/**
	 * Iterates and collects the list of the namespaces, used in URIs in the
	 * repository
	 */
	public void iterateNamespaces() throws Exception {
		System.out.println("===== Namespace List ==================================");

		System.out.println("Namespaces collected in the repository:");
		RepositoryResult<Namespace> iter = repositoryConnection.getNamespaces();

		while (iter.hasNext()) {
			Namespace namespace = iter.next();
			String prefix = namespace.getPrefix();
			String name = namespace.getName();
			namespacePrefixes.put(name, prefix);
			System.out.println(prefix + ":\t" + name);
		}
		iter.close();
	}

	/**
	 * Parse the query file and return the queries defined there for further
	 * evaluation. The file can contain several queries; each query starts with
	 * an id enclosed in square brackets '[' and ']' on a single line; the text
	 * in between two query ids is treated as a SPARQL query. Each line starting
	 * with a '#' symbol will be considered as a single-line comment and
	 * ignored. Query file syntax example:
	 * 
	 * #some comment [queryid1] <query line1> <query line2> ... <query linen>
	 * #some other comment [nextqueryid] <query line1> ... <EOF>
	 * 
	 * @param queryFile
	 * @return an array of strings containing the queries. Each string starts
	 *         with the query id followed by ':', then the actual query string
	 */
	private static String[] collectQueries(String queryFile) throws Exception {
		try {
			List<String> queries = new ArrayList<String>();
			BufferedReader input = new BufferedReader(new FileReader(queryFile));
			String nextLine = null;

			for (;;) {
				String line = nextLine;
				nextLine = null;
				if (line == null) {
					line = input.readLine();
				}
				if (line == null) {
					break;
				}
				line = line.trim();
				if (line.length() == 0) {
					continue;
				}
				if (line.startsWith("#")) {
					continue;
				}
				if (line.startsWith("^[") && line.endsWith("]")) {
					StringBuilder buff = new StringBuilder(line.substring(2, line.length() - 1));
					buff.append(": ");

					for (;;) {
						line = input.readLine();
						if (line == null) {
							break;
						}
						line = line.trim();
						if (line.length() == 0) {
							continue;
						}
						if (line.startsWith("#")) {
							continue;
						}
						if (line.startsWith("^[")) {
							nextLine = line;
							break;
						}
						buff.append(line);
						buff.append(System.getProperty("line.separator"));
					}

					queries.add(buff.toString());
				}
			}

			String[] result = new String[queries.size()];
			for (int i = 0; i < queries.size(); i++) {
				result[i] = queries.get(i);
			}
			input.close();
			return result;
		} catch (Exception e) {
			System.out.println("Unable to load query file '" + queryFile + "':" + e);
			return new String[0];
		}
	}

	private static final QueryLanguage[] queryLanguages = new QueryLanguage[] {
        QueryLanguage.SPARQL,
        //QueryLanguage.SERQL,
        //QueryLanguage.SERQO
    };

	/**
	 * The purpose of this method is to try to parse a query locally in order to
	 * determine if the query is a tuple (SELECT), boolean (ASK) or graph
	 * (CONSTRUCT/DESCRIBE) query. This happens automatically if the repository
	 * is local, but for a remote repository the local HTTPClient side can not
	 * work it out. Therefore a temporary in memory SAIL is created and used to
	 * determine the query type.
	 *
	 * @param query Query string to be parsed
	 * @param language The query language to assume
	 * @return A parsed query object or null if not possible
	 * @throws RepositoryException If the local repository used to test the query type failed for some reason
	 */
	private Query prepareQuery(String query, QueryLanguage language) throws RepositoryException {
		Repository tempRepository = new SailRepository(new MemoryStore());
		tempRepository.initialize();

		RepositoryConnection tempConnection = tempRepository.getConnection();

		try {
			try {
				tempConnection.prepareTupleQuery(language, query);
				return repositoryConnection.prepareTupleQuery(language, query);
			} catch (Exception e) {
			}

			try {
				tempConnection.prepareBooleanQuery(language, query);
				return repositoryConnection.prepareBooleanQuery(language, query);
			} catch (Exception e) {
			}

			try {
				tempConnection.prepareGraphQuery(language, query);
				return repositoryConnection.prepareGraphQuery(language, query);
			} catch (Exception e) {
			}

			return null;
		} finally {
			try {
				tempConnection.close();
				tempRepository.shutDown();
			} catch (Exception e) {
			}
		}
	}

    private Query prepareQuery(String query) throws Exception {

		for (QueryLanguage language : queryLanguages) {
			Query result = prepareQuery(query, language);
			if (result != null)
				return result;
		}
		// Can't prepare this query in any language
		return null;
	}

	/**
	 * Auxiliary method, printing an RDF value in a "fancy" manner. In case of
	 * URI, qnames are printed for better readability
	 *
	 * @param value
	 *            The value to beautify
	 */
	public String beautifyRDFValue(Value value) throws Exception {
		if (value instanceof URI) {
			URI u = (URI) value;
			String namespace = u.getNamespace();
			String prefix = namespacePrefixes.get(namespace);
			if (prefix == null) {
				prefix = u.getNamespace();
			} else {
				prefix += ":";
			}
			return prefix + u.getLocalName();
		} else {
            String full = value.toString();
            //if (full.indexOf("^^") > 0){
            //    full = full.substring(0, full.indexOf("^^"));
            //}
			return full;
		}
	}

    private void executeSingleQuery(String query) {
		try {
			Query preparedQuery = prepareQuery(query);
			if (preparedQuery == null) {
				System.out.println("Unable to parse query: " + query);
				return;
			}

			if (preparedQuery instanceof BooleanQuery) {
				System.out.println("Result: " + ((BooleanQuery) preparedQuery).evaluate());
				return;
			}

			if (preparedQuery instanceof GraphQuery) {
                System.out.println("GraphQuery");

				GraphQuery q = (GraphQuery) preparedQuery;
				long queryBegin = System.nanoTime();

				GraphQueryResult result = q.evaluate();
				int rows = 0;
				while (result.hasNext()) {
					Statement statement = result.next();
					rows++;
					System.out.print(beautifyRDFValue(statement.getSubject()));
					System.out.print(" " + beautifyRDFValue(statement.getPredicate()) + " ");
					System.out.print(" " + beautifyRDFValue(statement.getObject()) + " ");
					Resource context = statement.getContext();
					if (context != null)
						System.out.print(" " + beautifyRDFValue(context) + " ");
					System.out.println();
				}
				System.out.println();

				result.close();

				long queryEnd = System.nanoTime();
				System.out.println(rows + " result(s) in " + (queryEnd - queryBegin) / 1000000 + "ms.");
				System.out.println();
			}

			if (preparedQuery instanceof TupleQuery) {
                System.out.println("TupleQuery");
				TupleQuery q = (TupleQuery) preparedQuery;
				long queryBegin = System.nanoTime();

				TupleQueryResult result = q.evaluate();

				int rows = 0;
				while (result.hasNext()) {
					BindingSet tuple = result.next();
					if (rows == 0) {
						for (Iterator<Binding> iter = tuple.iterator(); iter.hasNext();) {
							System.out.print(iter.next().getName());
							System.out.print("\t");
						}
						System.out.println();
						System.out.println("---------------------------------------------");
					}
					rows++;
                    for (Iterator<Binding> iter = tuple.iterator(); iter.hasNext();) {
                        try {
                            System.out.print(beautifyRDFValue(iter.next().getValue()) + "\t");
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
					}
                    System.out.println();
				}
                System.out.println();

				result.close();

				long queryEnd = System.nanoTime();
				System.out.println(rows + " result(s) in " + (queryEnd - queryBegin) / 1000000 + "ms.");
				System.out.println();
			}
		} catch (Throwable e) {
			System.out.println("An error occurred during query execution: " + e.getMessage());
		}
	}
    /**
	 * Demonstrates query evaluation. First parse the query file. Each of the
	 * queries is executed against the prepared repository. If the printResults
	 * is set to true the actual values of the bindings are output to the
	 * console. We also count the time for evaluation and the number of results
	 * per query and output this information.
	 */
	public void evaluateQueries() throws Exception {
		System.out.println("===== Query Evaluation ======================");

		if (QUERYFILE == null) {
			System.out.println("No query file given in parameter '" + QUERYFILE + "'.");
			return;
		}

		long startQueries = System.currentTimeMillis();

		// process the query file to get the queries
		String[] queries = collectQueries(QUERYFILE);

		final CountDownLatch numberOfQueriesToProcess = new CountDownLatch(queries.length);
		// evaluate each query and, optionally, print the bindings
		for (int i = 0; i < queries.length; i++) {
			final String name = queries[i].substring(0, queries[i].indexOf(":"));
			final String query = queries[i].substring(name.length() + 2).trim();
			System.out.println("Executing query '" + name + "'");

        	executeSingleQuery(query);
		} // for

		long endQueries = System.currentTimeMillis();
		System.out.println("Queries run in " + (endQueries - startQueries) + " ms.");
	}

    /**
	 * Shutdown the repository and flush unwritten data.
	 */
	public void shutdown() {
		System.out.println("===== Shutting down ==========");
		if (repository != null) {
			try {
				repositoryConnection.close();
				repository.shutDown();
				repositoryManager.shutDown();
			} catch (Exception e) {
				System.out.println("An exception occurred during shutdown: " + e.getMessage());
			}
		}
	}
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws RepositoryConfigException, RepositoryException, IOException, Exception {
        OwlimTester tester = new OwlimTester();
	    long initializationStart = System.currentTimeMillis();
        //tester.openRepository();
        tester.newRepository();
        File file = new File(DATA_DIR);
        tester.loadFiles(file);
        tester.showInitializationStatistics(System.currentTimeMillis() - initializationStart);
        tester.iterateNamespaces();
        tester.evaluateQueries();
        tester.shutdown();
    }

}
