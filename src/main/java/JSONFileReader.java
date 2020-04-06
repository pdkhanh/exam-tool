import org.json.simple.parser.JSONParser;
import org.json.simple.*;
import org.json.simple.parser.ParseException;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;



import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class JSONFileReader {
    public static Configuration readConfiguration() throws Exception{
        JSONParser parser = new JSONParser();
        Configuration config = new Configuration();

        try {
            Object obj = parser.parse(new FileReader("./config.json"));

            JSONObject jsonObject =  (JSONObject) obj;

            config.setInputFile((String) jsonObject.get("inputFileName"));
            config.setOutputFile((String) jsonObject.get("outputFileName"));
            config.setOutputSheet((String) jsonObject.get("outputSheetName"));
            config.setNumberOfFileToGenerate((int) (long) jsonObject.get("numberOfFileToGenerate"));

            JSONArray cars = (JSONArray) jsonObject.get("configuration");

            Iterator<JSONObject> iterator = cars.iterator();
            List<KeepQuesion> keepQuesionList = new ArrayList();
            while (iterator.hasNext()) {
                JSONObject o = iterator.next();
                KeepQuesion keepQuesion = new KeepQuesion((String) o.get("sheetName"), (long) o.get("numberQuestion"));
                keepQuesionList.add(keepQuesion);
            }
            config.setKeepQuesionList(keepQuesionList);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return config;
    }
}
