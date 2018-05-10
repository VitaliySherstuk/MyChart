package chart;


import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import java.util.*;

public class ParseJson {


    public List<DataChart> parseJsonRadar(String stringToJson, String[] parameters)
    {
        List<DataChart> list = new ArrayList<>();
        String correctJson = stringToJson.replace("'", "\"").replace("@", " ").replace("_", " ").replace("Agreed upon Timeline", "Agreed-upon Timeline");

       // System.out.println("INPUT JSON: " + stringToJson);
       // System.out.println("Correct JSON: " + correctJson);
        JSONParser parser = new JSONParser();
        try {

            JSONArray json = (JSONArray) parser.parse(correctJson);

            for(int i=0; i<json.size(); i++)
            {
                Map<String, String> mapData = new HashMap<>();
                JSONObject objectJson = (JSONObject)json.get(i);

                for(int y=0; y<parameters.length; y++) {

                    //System.out.print(parameters[y]);
                    double value = Double.valueOf(objectJson.get(parameters[y]).toString());
                    //System.out.println(" --> " + value);
                    if(Math.abs(value)>=10)
                    {
                        double correctValue = Math.round(value);
                        mapData.put(parameters[y].substring(1).toLowerCase(), String.valueOf(correctValue));
                    }
                    else
                    {
                        double correctValue = Math.round(value*100);
                        mapData.put(parameters[y].substring(1).toLowerCase(), String.valueOf(correctValue/100));
                    }

                }
                list.add(new DataChart(objectJson.get(" CSAT round").toString(), mapData));
            }

        } catch (ParseException e) {
            e.printStackTrace();
        }

        return list;
    }

    public Map<String, List<DataChart>> parseJsonBar(String stringToJson, Map<String, String[]>parametersBarColumn)
    {
        Set<String> diagramKind = parametersBarColumn.keySet();
        Map<String, List<DataChart>> listDiagramParameters = new HashMap<>();
        String correctJson = stringToJson.replace("'", "\"").replace("@", " ").replace("_", " ").replace("Agreed upon Timeline", "Agreed-upon Timeline");

        for(String key : diagramKind)
        {
            List<DataChart> listDataChart = new ArrayList<>();
            String[] nameParameters = parametersBarColumn.get(key);
            JSONParser parser = new JSONParser();
            try {

                JSONArray json = (JSONArray) parser.parse(correctJson);

                for(int i=0; i<json.size(); i++)
                {
                    Map<String, String> mapData = new HashMap<>();
                    JSONObject objectJson = (JSONObject)json.get(i);
                    for(int y=0; y<nameParameters.length; y++) {

                        String parameter = objectJson.get(nameParameters[y]).toString();
                        double value = Double.valueOf(parameter);
                        if(Math.abs(value)>=10)
                        {
                            double correctValue = Math.round(value);
                            mapData.put(nameParameters[y].substring(1).toLowerCase(), String.valueOf(correctValue));
                        }
                        else
                        {
                            double correctValue = Math.round(value*100);
                            mapData.put(nameParameters[y].substring(1).toLowerCase(), String.valueOf(correctValue/100));
                        }

                    }
                    listDataChart.add(new DataChart(objectJson.get(" CSAT round").toString(), mapData));

                }

            } catch (ParseException e) {
                e.printStackTrace();
            }
            listDiagramParameters.put(key, listDataChart);
        }

        return listDiagramParameters;
    }

}
