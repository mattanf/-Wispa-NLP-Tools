package utility;

public class AddressComponent {

	public static final AddressComponent acStreetAddress = new AddressComponent("street_address", "indicates a precise street address.");
	public static final AddressComponent acRoute = new AddressComponent("route", "indicates a named route (such as \"US 101\").");
	public static final AddressComponent acIntersection = new AddressComponent("intersection", "indicates a major intersection, usually of two major roads.");
	public static final AddressComponent acPolitical = new AddressComponent("political", "indicates a political entity. Usually, this type indicates a polygon of some civil administration.");
	public static final AddressComponent acCountry = new AddressComponent("country", "indicates the national political entity, and is typically the highest order type returned by the Geocoder.");
	public static final AddressComponent acAdministrativeAreaLevel1 = new AddressComponent("administrative_area_level_1", "indicates a first-order civil entity below the country level. Within the United States, these administrative levels are states. Not all nations exhibit these administrative levels.");
	public static final AddressComponent acAdministrativeAreaLevel2 = new AddressComponent("administrative_area_level_2", "indicates a second-order civil entity below the country level. Within the United States, these administrative levels are counties. Not all nations exhibit these administrative levels.");
	public static final AddressComponent acAdministrativeAreaLevel3 = new AddressComponent("administrative_area_level_3", "indicates a third-order civil entity below the country level. This type indicates a minor civil division. Not all nations exhibit these administrative levels.");
	public static final AddressComponent acColloquial_Area = new AddressComponent("colloquial_area", "indicates a commonly-used alternative name for the entity.");
	public static final AddressComponent acLocality = new AddressComponent("locality", "indicates an incorporated city or town political entity.");
	public static final AddressComponent acSublocality = new AddressComponent("sublocality", "indicates a first-order civil entity below a locality. For some locations may receive one of the additional types: sublocality_level_1 through to sublocality_level_5. Each sublocality level is a civil entity. Larger numbers indicate a smaller geographic area.");
	public static final AddressComponent acSublocalityLevel1 = new AddressComponent("sublocality_level_1", "explained in sublocality.");
	public static final AddressComponent acSublocalityLevel2 = new AddressComponent("sublocality_level_2", "explained in sublocality.");
	public static final AddressComponent acSublocalityLevel3 = new AddressComponent("sublocality_level_3", "explained in sublocality.");
	public static final AddressComponent acSublocalityLevel4 = new AddressComponent("sublocality_level_4", "explained in sublocality.");
	public static final AddressComponent acSublocalityLevel5 = new AddressComponent("sublocality_level_5", "explained in sublocality.");
	public static final AddressComponent acSublocalityLevel6 = new AddressComponent("sublocality_level_6", "explained in sublocality.");
	public static final AddressComponent acSublocalityLevel7 = new AddressComponent("sublocality_level_7", "explained in sublocality.");
	public static final AddressComponent acSublocalityLevel8 = new AddressComponent("sublocality_level_8", "explained in sublocality.");
	public static final AddressComponent acPremise = new AddressComponent("premise", "indicates a named location, usually a building or collection of buildings with a common name");
	public static final AddressComponent acSubpremise = new AddressComponent("subpremise", "indicates a first-order entity below a named location, usually a singular building within a collection of buildings with a common name");
	public static final AddressComponent acPostalCode = new AddressComponent("postal_code", "indicates a postal code as used to address postal mail within the country.");
	public static final AddressComponent acNaturalFeature = new AddressComponent("natural_feature", "indicates a prominent natural feature.");
	public static final AddressComponent acAirport = new AddressComponent("airport", "indicatesan airport.");
	public static final AddressComponent acPark = new AddressComponent("park", "indicates a named park.");
	public static final AddressComponent acInterentPoint = new AddressComponent("point_of_interest", "indicates a named point of interest. Typically, these \"POI\"s are prominent local entities that don't easily fit in another category such as \"Empire State Building\" or \"Statue of Liberty.\"");
	
	public static final AddressComponent acSynagogue = new AddressComponent("synagogue", "");
	public static final AddressComponent acChurch = new AddressComponent("church","");
	public static final AddressComponent acEstablishment = new AddressComponent("establishment", "");
	public static final AddressComponent acSubwayStation = new AddressComponent("subway_station", "");
	public static final AddressComponent acTrainStation = new AddressComponent("train_station", "");
	public static final AddressComponent acTransitStation = new AddressComponent("transit_station", "");
	
	private String name;
	private String description;
	
	public AddressComponent(String name, String description)
	{
		this.name = name;
		this.description = description;
	}
	
	@Override
	public boolean equals(Object o)
	{
		return (o == this) || ((o instanceof AddressComponent) && (((AddressComponent)o).getName().equals(getName())));
	}
	
	@Override
	public int hashCode()
	{
		return getName().hashCode();
	}
	
	public String getName() { return name; }
	public String getDescription() { return description; }
	

}
