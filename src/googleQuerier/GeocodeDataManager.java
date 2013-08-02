package googleQuerier;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.OutputStream;
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

public class GeocodeDataManager {

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
		private GeocoderResult res;

		public QueryData(GeocodeQueryDescriptor desc, GeocoderResult res) {
			this.desc = desc;
			this.res = res;
		}

		public GeocodeQueryDescriptor getDescription() {
			return desc;
		}

		public GeocoderResult getResult() {
			return res;
		}
	}

	final static private Geocoder geocoder = new Geocoder();

	private File file;
	private TreeMap<String, QueryData> existingQueries;
	private ObjectOutputStream objOutStream;
	private static int queryLimitHitCount = 0;
	private static int queryErrorCount = 0;
	private static int queriesPerformed = 0;

	public GeocodeDataManager(File cityFile) throws ClassNotFoundException, IOException {
		existingQueries = new TreeMap<String, QueryData>(String.CASE_INSENSITIVE_ORDER);
		this.file = cityFile;
		boolean isNeedAppend = false;
		if (file.exists()) {
			isNeedAppend = true;
			loadExistingData();
		}

		FileOutputStream fos = new FileOutputStream(file, isNeedAppend);
		if (isNeedAppend)
			objOutStream = new AppendingObjectOutputStream(fos);
		else
			objOutStream = new ObjectOutputStream(fos);
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

	private void loadExistingData() throws IOException, ClassNotFoundException {

		System.out.println("Begin reading file info: " + file.getAbsolutePath() + ".");
		FileInputStream fis = new FileInputStream(file);
		ObjectInputStream inpObj = new ObjectInputStream(fis);
		try {
			while (fis.available() > 0) {
				GeocoderResult res = null;
				GeocodeQueryDescriptor qd = (GeocodeQueryDescriptor) inpObj.readObject();
				if (qd.isRecordExists() == true) {
					res = (GeocoderResult) inpObj.readObject();
				}
				System.out.println("Read: " + qd.getQueriedLocationName() + ".");

				QueryData entry = new QueryData(qd, res);
				existingQueries.put(qd.getQueriedLocationName(), entry);
				if (qd.getResultingFormattedName() != null)
					existingQueries.put(qd.getResultingFormattedName(), entry);
			}
			System.out.println("End reading file info: " + file.getAbsolutePath() + ".");

			inpObj.close();
		} catch (IOException e) {
			System.out.println("Error reading file: " + file.getAbsolutePath() + ".");
			int res = JOptionPane.showConfirmDialog(null, "Source file is damaged. would you like to recreate?",
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
		boolean isSuccefull = false;
		File tmpFile;
		try {
			tmpFile = File.createTempFile("geocode-recover", "dat");

			FileOutputStream fos = new FileOutputStream(tmpFile, false);
			ObjectOutputStream objStrm = new ObjectOutputStream(fos);

			Iterator<QueryData> it = existingQueries.values().iterator();
			isSuccefull = true;
			while (it.hasNext() && isSuccefull) {
				isSuccefull = writeEntryToFile(objStrm, it.next());
			}
			if (isSuccefull) {
				objStrm.close();
				if (file.exists())
					file.delete();
				isSuccefull = tmpFile.renameTo(file);
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return isSuccefull;

	}

	public QueryData getOrCreateQueryData(String locationName) {
		QueryData qd = existingQueries.get(locationName);
		if (qd != null) {
			return qd;
		} else
			return createQueryData(locationName);
	}

	public boolean isQueryingAllowed() {
		return (queryLimitHitCount < 5 && queryErrorCount < 5 && objOutStream != null);
	}

	private QueryData createQueryData(String locationName) {

		QueryData resQueryData = null;
		if (isQueryingAllowed()) {
			if (queryLimitHitCount > 0)
				try {
					Thread.sleep(queryLimitHitCount * 30);
				} catch (InterruptedException e) {
					queryErrorCount++;
				}

			boolean doReattempt = true;
			boolean doStore = false;
			GeocodeResponse geocoderResponse = null;
			while (doReattempt) {
				geocoderResponse = doGeocodeQuery(locationName);
				doReattempt = false;

				GeocoderStatus lastGeocoderResponse = geocoderResponse.getStatus();
				switch (lastGeocoderResponse) {
				// "OVER_QUERY_LIMIT" indicates that you are over your quota.
				case OVER_QUERY_LIMIT:
					++queryLimitHitCount;
					doReattempt = isQueryingAllowed();
					break;
				// "UNKNOWN_ERROR" indicates that the request could not be processed due to a server error. The request may succeed if you try again.
				// "REQUEST_DENIED" indicates that your request was denied, generally because of lack of a sensor parameter.
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
				List<GeocoderResult> results = geocoderResponse.getResults();
				GeocoderResult res = null;
				if (results != null && results.size() > 0) {
					res = results.get(0);
				}
				GeocodeQueryDescriptor desc = new GeocodeQueryDescriptor(locationName, res);

				resQueryData = new QueryData(desc, res);
				addEntryToDictionary(resQueryData);
				writeEntryToFile(objOutStream, resQueryData);

			}
		} else {
			System.out.println("Unable to perform query for " + locationName + ". System failed to many times. ");
		}
		return resQueryData;
	}

	/**
	 * @param resQueryData
	 */
	private void addEntryToDictionary(QueryData resQueryData) {
		existingQueries.put(resQueryData.getDescription().getQueriedLocationName(), resQueryData);
		if (resQueryData.getDescription().getResultingFormattedName() != null)
			existingQueries.put(resQueryData.getDescription().getResultingFormattedName(), resQueryData);
	}

	/**
	 * @param resQueryData
	 * @param outStream
	 */
	private boolean writeEntryToFile(ObjectOutputStream outStream, QueryData resQueryData) {
		// Save the data to disk
		try {
			outStream.writeObject(resQueryData.getDescription());
			if (resQueryData.getDescription().isRecordExists()) {
				outStream.writeObject(resQueryData.getResult());
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
		return true;
	}

	/**
	 * Performs a Geocode query in front of google's servers
	 * 
	 * @param locationName location to query
	 * @return results
	 */
	private GeocodeResponse doGeocodeQuery(String locationName) {
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
