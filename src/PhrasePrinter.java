
import java.io.InputStream;
import java.util.TreeSet;

import javax.swing.JOptionPane;

import com.pairapp.datalayer.ConfigurationDatalayer;
import com.pairapp.datalayer.ConfigurationDatalayer.ConfigSourceType;
import com.pairapp.engine.parser.ParseGlobalData;
import com.pairapp.engine.parser.PhraseRepository;
import com.pairapp.engine.parser.PhraseRepositoryDataProviderXml;

public class PhrasePrinter {

	public static void main(String[] args) {
		//Object[] possibilities = {"ham", "spam", "yam"};
		String sourceName;
		sourceName = (String)JOptionPane.showInputDialog(
                    null,
                    "Resource to print:\n",
                    "Question",
                    JOptionPane.PLAIN_MESSAGE,
                    null,
                    null,
                    "RealEstateHebrew-Mattan");
		// Generate the repository
		InputStream inpPhraseStream = ConfigurationDatalayer.getStream(sourceName, ConfigSourceType.Pharse);
		// Load the file
		PhraseRepositoryDataProviderXml dataProvider = new PhraseRepositoryDataProviderXml();
		boolean success = (dataProvider != null);
		// Try to initialize
		if (success == true)
			success &= dataProvider.init(inpPhraseStream);
		if (success == true) {
			PhraseRepository retRepo = new PhraseRepository();
			success &= retRepo.init(dataProvider, true);
			ParseGlobalData rep = new ParseGlobalData(retRepo);
			
			if (success == true)
			{
				java.util.Iterator<String> it = retRepo.getKeyIterator();
				TreeSet<String> tsKey = new TreeSet<>();
				while (it.hasNext())
				{
					tsKey.add(it.next());
				}
				java.util.Iterator<String> tsIt = tsKey.iterator();
				while (tsIt.hasNext())
				{
					String key = tsIt.next();
					System.out.println(key + "=" + rep.unrollString(retRepo.getPhrase(key, null), null));
				}
			}
		}
		

	}

}