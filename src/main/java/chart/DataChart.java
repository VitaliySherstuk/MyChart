package chart;

import java.util.Map;

public class DataChart {


    private String period;
    private Map<String, String> data;

    public DataChart(String period, Map<String, String> data)
    {

        this.period = period;
        this.data = data;
    }

    public String getPeriod() {
        return period;
    }

    public Map<String, String> getData() {
        return data;
    }

    @Override
    public String toString() {
        return "DataChart{" +
                "period='" + period + '\'' +
                ", data=" + data +
                '}';
    }
}
