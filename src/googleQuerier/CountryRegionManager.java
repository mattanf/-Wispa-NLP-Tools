package googleQuerier;


public class CountryRegionManager {

	static CountryRegionManager s_Manager = new CountryRegionManager();
	public static CountryRegionManager instance() {
		return s_Manager;
	}
	

}
