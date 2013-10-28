package israelDataGoogleQuerier;

import java.io.Serializable;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import utility.XLSUtil;

/**
 * Contains the information for a single location for the
 * 
 * @author PC
 * 
 */
public class LocationEntry implements Comparable<LocationEntry>, Serializable {
	private static final long serialVersionUID = -7120985814348365694L;

	static final int ROW_CITY_ID = 2;
	static final int ROW_CITY_NAME = 3;
	static final int ROW_AREA_ID = 4;
	static final int ROW_AREA_NAME = 5;
	static final int SUGGESTED_CITY_NAME = 6;
	static final int SUGGESTED_AREA_NAME = 7;

	private String cityId;
	private String cityName;
	private String areaId;
	private String areaName;
	private String queryCityName;
	private String queryAreaName;

	public LocationEntry(String cityId, String cityName, String areaId, String areaName, String queryCityName,
			String queryAreaName) {
		this.cityId = cityId;
		this.cityName = cityName;
		this.areaId = areaId;
		this.areaName = areaName;
		this.queryCityName = queryCityName;
		this.queryAreaName = queryAreaName;

		if (queryCityName == null || queryCityName.isEmpty())
			this.queryCityName = cityName;
		if (queryAreaName == null || queryAreaName.isEmpty())
			this.queryAreaName = areaName;
		if (isCity()) {
			this.queryAreaName = null;
		}
	}

	public LocationEntry(Workbook streetWB, int readRow) {
		this((XLSUtil.getCellString(streetWB, readRow, ROW_CITY_ID)), (XLSUtil.getCellString(streetWB, readRow, ROW_CITY_NAME)),
				(XLSUtil.getCellString(streetWB, readRow, ROW_AREA_ID)), (XLSUtil.getCellString(streetWB, readRow, ROW_AREA_NAME)),
				(XLSUtil.getCellString(streetWB, readRow, SUGGESTED_CITY_NAME)), (XLSUtil.getCellString(streetWB, readRow,
						SUGGESTED_AREA_NAME)));

	}

	public String getCityId() {
		return cityId;
	}

	public String getCityName() {
		return cityName;
	}

	public String getAreaId() {
		return areaId;
	}

	public String getAreaName() {
		return areaName;
	}

	public String getQueryCityName() {
		return queryCityName;
	}

	public String getQueryAreaName() {
		return queryAreaName;
	}

	@Override
	public boolean equals(Object obj) {
		if (obj instanceof LocationEntry) {
			LocationEntry le = (LocationEntry) obj;
			return le.cityId.equals(cityId) && le.cityName.equals(cityName) && le.areaId.equals(areaId) &&
					le.areaName.equals(areaName);
		}
		return false;
	}

	@Override
	public int hashCode() {
		return cityId.hashCode() ^ cityName.hashCode() ^ areaId.hashCode() ^ areaName.hashCode();
	}

	@Override
	public int compareTo(LocationEntry le) {
		int c1 = cityId.compareTo(le.cityId);
		int c2 = cityName.compareTo(le.cityName);
		int c3 = areaId.compareTo(le.areaId);
		int c4 = areaName.compareTo(le.areaName);

		return signMul(c1, 1000) + signMul(c2, 100) + signMul(c3, 10) + signMul(c4, 1);
	}

	private int signMul(int num, int mul) {
		if (num < 0)
			num = -1;
		else if (num > 0)
			num = 1;
		return num * mul;
	}

	public LocationEntry generateCityLocation() {
		return new LocationEntry(cityId, cityName, cityId, cityName, null, null);
	}

	public boolean isCity() {
		return cityName.equals(areaName) || areaName == null || areaName.isEmpty();
	}

	public boolean isSameCity(LocationEntry loc) {
		return cityName.equals(loc.cityName);

	}

}
