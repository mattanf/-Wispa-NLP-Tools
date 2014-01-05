package israelDataGoogleQuerier;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.TreeMap;

import javax.swing.JOptionPane;

import com.google.code.geocoder.Geocoder;
import com.google.code.geocoder.GeocoderRequestBuilder;
import com.google.code.geocoder.model.GeocodeResponse;
import com.google.code.geocoder.model.GeocoderRequest;
import com.google.code.geocoder.model.GeocoderResult;
import com.google.code.geocoder.model.GeocoderStatus;
import com.pairapp.utilities.StringUtilities;

public class GeocodeDataManager {

	private static final String GEOCODE_FILE_PREFIX = "Geo-";
	private static final String GEOCODE_FILE_POSTFIX = ".dat";

	public class AppendingObjectOutputStream extends ObjectOutputStream {

		public AppendingObjectOutputStream(OutputStream out) throws IOException {
			super(out);
		}

		@Override
		protected void writeStreamHeader() throws IOException {
			// do not write a header, but reset:
			// this line added after another question
			// showed a problem with the original
			reset();
		}

	}

	static class QueryData {
		private GeocodeQueryDescriptor desc;
		private GeocoderResult[] ress;

		public QueryData(GeocodeQueryDescriptor desc, GeocoderResult[] ress) {
			this.desc = desc;
			this.ress = ress;
		}

		public GeocodeQueryDescriptor getDescription() {
			return desc;
		}

		public GeocoderResult[] getResults() {
			return ress;
		}
	}

	final static private Geocoder geocoder = new Geocoder();

	private File file;
	private TreeMap<String, QueryData> existingQueries;
	private HashMap<GeocoderResult, GeocoderResult> uniqueGeoResults;
	private ObjectOutputStream objOutStream;
	private String baseLocationName;
	private String baseLocationNameShort;
	private static int queryErrorCount = 0;
	private static int queriesPerformed = 0;
	private static double sleepTimeBetweenQueries = 0;

	public GeocodeDataManager(String workingDir, String baseLocationName) {

		File cityFile = new File(workingDir, GEOCODE_FILE_PREFIX +
				StringUtilities.removeUnsafeFileChars(baseLocationName) + GEOCODE_FILE_POSTFIX);

		if (baseLocationName.contains(","))
			baseLocationNameShort = baseLocationName.substring(0, baseLocationName.indexOf(","));
		this.baseLocationName = baseLocationName;
		this.file = cityFile;
		existingQueries = new TreeMap<String, QueryData>(String.CASE_INSENSITIVE_ORDER);
		uniqueGeoResults = new HashMap<GeocoderResult, GeocoderResult>();
		try {
			if (file.exists())
				loadExistingData();

		} catch (Exception e) {
			System.out.println("Unexpected exception at GeocodeDataManager initialization.\n" + e.getMessage());
			e.printStackTrace();
			System.exit(0);
		}
	}

	public boolean isEmpty() {
		return uniqueGeoResults.isEmpty();
	}

	public int getUniqueResultCount() {
		return uniqueGeoResults.size();
	}

	public void close() {
		if (objOutStream != null) {
			try {
				objOutStream.close();
			} catch (Exception ex) {
			}
			objOutStream = null;
		}
	}

	public Iterator<GeocoderResult> getUniqueGeoResults() {
		return uniqueGeoResults.values().iterator();
	}

	public String getBaseLocationName() {
		return baseLocationName;
	}

	private void loadExistingData() throws IOException, ClassNotFoundException {

		System.out.println("Begin reading file info: " + file.getAbsolutePath() + ".");
		FileInputStream fis = new FileInputStream(file);
		ObjectInputStream inpObj = new ObjectInputStream(fis);
		long fileSize = file.length();
		//long lastCaughtSize = 0;
		try {
			int MAGIC_NUMBER = 6;
			ArrayList<GeocoderResult> arrRes = new ArrayList<GeocoderResult>();
			while (fis.getChannel().position() < fileSize - MAGIC_NUMBER) {
				//lastCaughtSize = fis.getChannel().position();
				//GeocoderResult res = null;
				GeocodeQueryDescriptor qd = (GeocodeQueryDescriptor) inpObj.readObject();
				arrRes.clear();
				int resultsToExtract = Math.max(qd.getNumExtraRecords(), (qd.isRecordExists() ? 1 : 0));

				for (; resultsToExtract > 0; --resultsToExtract)
					arrRes.add((GeocoderResult) inpObj.readObject());

				// System.out.println("Read: " + qd.getQueriedLocationName() + ".");

				addEntryToDictionary(qd, arrRes);
			}
			System.out.println("End reading file info: " + file.getAbsolutePath() + ".");

			inpObj.close();
		} catch (IOException e) {
			System.out.println("Error reading file: " + file.getAbsolutePath() + ".");
			int res = JOptionPane.showConfirmDialog(null, "ForumSource file is damaged. would you like to recreate?",
					"WARNING", JOptionPane.YES_NO_OPTION);
			boolean resaveSuccess = false;
			if (res == JOptionPane.YES_OPTION) {
				inpObj.close();
				resaveSuccess = resaveData(file);
			}
			if (resaveSuccess == false)
				System.exit(0);
		}
	}

	private boolean resaveData(File file) {
		boolean isSuccefull = true;
		File tmpFile;
		try {
			tmpFile = File.createTempFile("geocode-recover", "dat");

			FileOutputStream fos = new FileOutputStream(tmpFile, false);
			ObjectOutputStream objStrm = new ObjectOutputStream(fos);

			Iterator<QueryData> it = existingQueries.values().iterator();
			isSuccefull = true;
			while (it.hasNext() && isSuccefull) {
				QueryData qd = it.next();
				isSuccefull = writeEntryToFile(objStrm, qd);
			}
			if (isSuccefull) {
				objStrm.close();
				if (file.exists())
					file.delete();
				isSuccefull = tmpFile.renameTo(file);
			}
		} catch (Exception e) {

			e.printStackTrace();
		}

		return isSuccefull;

	}

	public QueryData getOrCreateQueryData(String string) {
		return getOrCreateQueryData(string, true);
	}

	public QueryData getOrCreateQueryData(String locationName, boolean allowCreate) {
		QueryData qd = existingQueries.get(locationName);
		if ((qd != null) || (!allowCreate)) {
			return qd;
		} else
			return createQueryData(locationName);

	}

	static public boolean isQueryingAllowed() {
		return (queryErrorCount < 5);
	}

	private QueryData createQueryData(String locationName) {

		QueryData resQueryData = null;
		if (isQueryingAllowed()) {
			boolean doReattempt = true;
			boolean doStore = false;
			GeocodeResponse geocoderResponse = null;
			int overLimitHitCount = 0;
			while (doReattempt) {
				geocoderResponse = doGeocodeQuery(locationName);
				doReattempt = false;

				GeocoderStatus lastGeocoderResponse = geocoderResponse.getStatus();
				switch (lastGeocoderResponse) {
				// "OVER_QUERY_LIMIT" indicates that you are over your quota.
				case OVER_QUERY_LIMIT:
					sleepTimeBetweenQueries += 20.0;
					++overLimitHitCount;
					if (overLimitHitCount > 5)
						queryErrorCount = 1000;
					doReattempt = isQueryingAllowed();
					break;
				// "UNKNOWN_ERROR" indicates that the request could not be processed due to a server error. The request
				// may succeed if you try again.
				// "REQUEST_DENIED" indicates that your request was denied, generally because of lack of a sensor
				// parameter.
				// "INVALID_REQUEST" generally indicates that the query (address or latlng) is missing.
				case ERROR:
				case UNKNOWN_ERROR:
				case REQUEST_DENIED:
				case INVALID_REQUEST:
					++queryErrorCount;
					break;
				// "OK" indicates that no errors occurred; the address was successfully parsed and at least one geocode
				// was returned.
				// "ZERO_RESULTS" indicates that the geocode was successful but returned no results. This may occur if
				// the geocode was passed a non-existent address or a latlng in a remote location.
				case OK:
				case ZERO_RESULTS:
					doStore = true;
					break;
				}
				System.out.println((doStore ? "Success" : "Fail") + " (" + lastGeocoderResponse.name() + ").");
			}

			// store the data if possible
			if (doStore) {
				// Save the data internally
				List<GeocoderResult> tmpResults = geocoderResponse.getResults();
				ArrayList<GeocoderResult> results = new ArrayList<GeocoderResult>();

				if (tmpResults != null) {
					for (int i = 0; i < tmpResults.size(); ++i) {
						String resAddress = tmpResults.get(i).getFormattedAddress();
						boolean isShortNameMatch = (baseLocationNameShort != null) &&
								(resAddress.endsWith(baseLocationNameShort));
						if ((!resAddress.equals(baseLocationName)) &&
								((resAddress.endsWith(baseLocationName)) || (isShortNameMatch))) {
							results.add(tmpResults.get(i));
						}
					}
				}

				orderGeocodeResults(results);

				GeocoderResult res = null;
				if (results.size() > 0)
					res = results.get(0);
				GeocodeQueryDescriptor desc = new GeocodeQueryDescriptor(locationName, res, results.size());

				resQueryData = addEntryToDictionary(desc, results);
				ensureOutputStreamOpen();
				writeEntryToFile(objOutStream, resQueryData);
				close();

			}
		} else {
			System.out.println("Unable to perform query for " + locationName + ". System failed to many times. ");
		}
		return resQueryData;
	}

	/**
	 * @param results
	 */
	private void orderGeocodeResults(ArrayList<GeocoderResult> results) {
		Collections.sort(results, new Comparator<GeocoderResult>() {
			public int compare(GeocoderResult loc1, GeocoderResult loc2) {
				return getTypePriority(loc1) - getTypePriority(loc2);
			}

			public int getTypePriority(GeocoderResult res) {
				int matchOrder = res.isPartialMatch() ? 10 : 0;
				String curType = res.getTypes().size() > 0 ? res.getTypes().get(0) : null;
				if (curType != null) {
					if (curType.equals("route"))
						return matchOrder + 1;
					else if (curType.equals("neighborhood"))
						return matchOrder + 2;
					else if (curType.equals("park"))
						return matchOrder + 3;
					else if (curType.equals("point_of_interest"))
						return matchOrder + 4;
					else if (curType.equals("establishment"))
						return matchOrder + 4;
					else if (curType.equals("locality"))
						return matchOrder + 5;
				}
				return matchOrder + 10;
			}
		});
	}

	/**
	*/
	private void ensureOutputStreamOpen() {
		if (objOutStream == null) {
			try {
				boolean appendToFile = file.exists() && file.length() != 0;
				FileOutputStream fos = new FileOutputStream(file, appendToFile);
				if (appendToFile)
					objOutStream = new AppendingObjectOutputStream(fos);
				else
					objOutStream = new ObjectOutputStream(fos);
			} catch (Exception e) {
				System.out.println("Unexpected exception at GeocodeDataManager initialization.\n" + e.getMessage());
				e.printStackTrace();
				System.exit(0);
			}
		}

	}

	/**
	 * @param results
	 * @param resQueryData
	 */
	private QueryData addEntryToDictionary(GeocodeQueryDescriptor desc, ArrayList<GeocoderResult> results) {
		GeocoderResult firstRes = null;
		ArrayList<GeocoderResult> uniqueResults = new ArrayList<GeocoderResult>();
		if ((results != null) && (results.size() > 0)) {
			firstRes = addUniqueResultToDictionary(results.get(0));
			uniqueResults.add(firstRes);
			for (int i = 1; i < results.size(); ++i) {
				uniqueResults.add(addUniqueResultToDictionary(results.get(i)));
			}
		}
		// Add extra result
		GeocoderResult[] resArray = null;
		if (!uniqueResults.isEmpty())
			resArray = uniqueResults.toArray(new GeocoderResult[uniqueResults.size()]);
		QueryData resQueryData = new QueryData(desc, resArray);
		existingQueries.put(resQueryData.getDescription().getQueriedLocationName(), resQueryData);
		if (resQueryData.getDescription().getResultingFormattedName() != null)
			existingQueries.put(resQueryData.getDescription().getResultingFormattedName(), resQueryData);
		return resQueryData;
	}

	/**
	 * @param firstRes
	 * @return
	 */
	private GeocoderResult addUniqueResultToDictionary(GeocoderResult res) {
		GeocoderResult uniqueRes = uniqueGeoResults.get(res);
		if (uniqueRes == null)
			uniqueGeoResults.put(res, res);
		else
			res = uniqueRes;
		return res;
	}

	/**
	 * @param resQueryData
	 * @param outStream
	 */
	private boolean writeEntryToFile(ObjectOutputStream outStream, QueryData resQueryData) {
		if (resQueryData != null) {
			// Save the data to disk
			try {
				outStream.writeObject(resQueryData.getDescription());
				if (resQueryData.getDescription().isRecordExists()) {
					for (GeocoderResult res : resQueryData.getResults())
						outStream.writeObject(res);
				}
				outStream.flush();
			} catch (IOException e) {
				System.out.println("Error on outputing results to stream. " + e.getMessage());
				e.printStackTrace();
				try {
					outStream.close();
				} catch (Exception ex) {
					ex.printStackTrace();
					return false;
				}
			}
		}
		return true;
	}

	/**
	 * Performs a Geocode query in front of google's servers
	 * 
	 * @param locationName location to query
	 * @return results
	 */
	private GeocodeResponse doGeocodeQuery(String locationName) {
		sleepTimeBetweenQueries = Math.max(sleepTimeBetweenQueries - 0.01, 20);
		if (sleepTimeBetweenQueries > 0) {
			try {
				Thread.sleep((long) sleepTimeBetweenQueries);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		System.out.print("Querying location: " + locationName + " ... ");
		GeocoderRequest geocoderRequest = new GeocoderRequestBuilder().setAddress(locationName).getGeocoderRequest();
		geocoderRequest.setLanguage("he");
		GeocodeResponse geocoderResponse = geocoder.geocode(geocoderRequest);
		++queriesPerformed;

		return geocoderResponse;
	}

	public static int getQueriesPerformed() {
		return queriesPerformed;
	}

}
