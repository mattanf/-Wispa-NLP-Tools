package googleQuerier;

import java.io.Serializable;

import com.google.code.geocoder.model.GeocoderResult;

public class GeocodeQueryDescriptor implements Serializable {

	private static final long serialVersionUID = -4889257111978756549L;
	
	private String queriedLocationName;
	private String resultedFormatedName;
	private boolean isPartialMatch;
	private boolean recordExists;
	private int recordSizeInFile;
	protected boolean isValid;
	protected int FLU;
	
	public GeocodeQueryDescriptor(String queriedLocationName, String resultedFormatedName,
			boolean isPartialMatch, boolean recordExists, int recordSizeInFile)
	{
		this.queriedLocationName = queriedLocationName; 
		this.resultedFormatedName = resultedFormatedName;
		this.isPartialMatch = isPartialMatch;
		this.recordExists = recordExists;
		this.recordSizeInFile = recordSizeInFile;
		isValid = true;
		FLU = 0;						
	}
	
	public GeocodeQueryDescriptor(String locationName, GeocoderResult res) {
		this(locationName,
				res != null ? res.getFormattedAddress() : null,
				res != null ? res.isPartialMatch() : true,
				res != null, 0);
	}

	public String getQueriedLocationName() {
		return queriedLocationName;
	}
	public String getResultingFormattedName() {
		return resultedFormatedName;
	}
	public boolean isPartialMatch() {
		return isPartialMatch;
	}
	public boolean isRecordExists() {
		return recordExists;
	}
	public int getRecordSizeInFile() {
		return recordSizeInFile;
	}
	public void setRecordSizeInFile(int recordSizeInFile) {
		this.recordSizeInFile = recordSizeInFile;
	}
}
