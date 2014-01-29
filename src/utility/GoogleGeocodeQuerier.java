package utility;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;

import com.google.code.geocoder.Geocoder;
import com.google.code.geocoder.GeocoderRequestBuilder;
import com.google.code.geocoder.model.GeocodeResponse;
import com.google.code.geocoder.model.GeocoderRequest;
import com.google.code.geocoder.model.GeocoderResult;
import com.google.code.geocoder.model.GeocoderStatus;

public class GoogleGeocodeQuerier {
	
	final static private Geocoder geocoder = new Geocoder();
	
	double sleepTimeBetweenQueries = 0;
	int queryErrorCount = 0;
	String language;
	private List<AddressComponent> orderByComponent;
	private List<AddressComponent> removeComponent;
	
	public GoogleGeocodeQuerier(String language)
	{
		this.language = language;
	}
	public boolean isQueryingAllowed() {
		return (queryErrorCount < 5);
	}
	
	public List<AddressComponent> getOrderByComponent() {
		return orderByComponent;
	}
	public void setOrderByComponent(List<AddressComponent> orderByComponent) {
		this.orderByComponent = orderByComponent;
	}
	public List<AddressComponent> getRemoveComponent() {
		return removeComponent;
	}
	public void setRemoveComponent(List<AddressComponent> removeComponent) {
		this.removeComponent = removeComponent;
	}
	
	public List<GeocoderResult> makeQuery(String locationName) {

		List<GeocoderResult> queryResults = null;
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
				ArrayList<GeocoderResult> results = new ArrayList<GeocoderResult>(tmpResults);
				removeGeocodeResults(results);
				orderGeocodeResults(results);
				queryResults = results;
			}			
		} else {
			System.out.println("Unable to perform query for " + locationName + ". System failed to many times. ");
		}
		return queryResults;
	}
	
	/**
	 * Removes results that are predefined as unneeded by the user
	 * @param results
	 */
	private void removeGeocodeResults(ArrayList<GeocoderResult> results) {
		if (getRemoveComponent() != null)
		{
			Iterator<GeocoderResult> iterator = results.iterator();
			while (iterator.hasNext())
			{
				GeocoderResult next = iterator.next();
				int matchCount = 0;
				for(AddressComponent ac : getRemoveComponent())
				{
					if (next.getTypes().contains(ac.getName()))
					{
						++matchCount;
					}
				}
				if (next.getTypes().size() == matchCount)
					iterator.remove();
			}
		}
	}
	
	/**
	 * Sorts the results by order defined by the user
	 * @param results
	 */
	private void orderGeocodeResults(ArrayList<GeocoderResult> results) {
		if (getOrderByComponent() != null)
		{
			final List<AddressComponent> orderByComponent = this.getOrderByComponent();
			Collections.sort(results, new Comparator<GeocoderResult>() {
				public int compare(GeocoderResult loc1, GeocoderResult loc2) {
					return getTypePriority(loc1) - getTypePriority(loc2);
				}

				public int getTypePriority(GeocoderResult res) {
					int matchedIndex = 0;
					for(; matchedIndex < orderByComponent.size() ; ++matchedIndex)
						if (res.getTypes().contains(orderByComponent.get(matchedIndex).getName()))
							break;
					return matchedIndex + (res.isPartialMatch() ? 100000 : 0);
				}
			});
		}
		
	}

	
	/**
	 * Performs a Geocode query in front of google's servers
	 * 
	 * @param locationName location to query
	 * @return results
	 */
	private GeocodeResponse doGeocodeQuery(String locationName) {
		sleepTimeBetweenQueries = Math.max(sleepTimeBetweenQueries - 0.01, 10);
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
		geocoderRequest.setLanguage("en");
		GeocodeResponse geocoderResponse = geocoder.geocode(geocoderRequest);
		
		return geocoderResponse;
	}
	
}
