package chart;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;

import java.util.List;

public class ChartLineNPS {

    public void setLineGoal(XSLFSlide slide, String value, List<DataChart> dataLineList) {


        XSLFShape shapeTrendLine = null;
        for(XSLFShape shape : slide.getShapes())
        {
            //if(shape.getShapeName().equals("Chart 16") || shape.getShapeName().equals("Chart 7"))
            if(shape.getShapeName().contains("Chart"))
            {
                shapeTrendLine = shape;
            }
        }

        XSLFChart chartTrendLine = null;
        CTChart ctChrtTrendLine = null;

        for(POIXMLDocumentPart docPart : shapeTrendLine.getSheet().getRelations())
        {
            if (docPart instanceof XSLFChart) {

                chartTrendLine = (XSLFChart) docPart;
                if(chartTrendLine.getCTChart().getPlotArea().getBarChartList().size()==0)
                {
                    ctChrtTrendLine = chartTrendLine.getCTChart();
                }

            }
        }

        double minValue = 0.0;
        if(dataLineList.size() == 1)
        {
            minValue = Double.valueOf(dataLineList.get(0).getData().get("response rate"));
        }
        else
        {
            minValue = getMinValueResponseRate(dataLineList);
        }

        if(minValue<0)
        {
            double minMaajorGridLine = getMinMajorGridline(minValue);
            ctChrtTrendLine.getPlotArea().getValAxList().get(0).getScaling().getMin().setVal(minMaajorGridLine*(-1));
        }

        List<CTNumVal> ctNums = ctChrtTrendLine.getPlotArea().getLineChartArray(0).getSerList().get(0).getVal().getNumRef().getNumCache().getPtList();
        for (CTNumVal ctNumVal : ctNums) {
            ctNumVal.setV(value);
        }

    }

    //get min value of "NPS"
    private double getMinValueResponseRate(List<DataChart> dataLineList)
    {
        double minValue = Double.valueOf(dataLineList.get(0).getData().get("nps"));
        for(int i=1; i<dataLineList.size(); i++)
        {
            double checkValue = Double.valueOf(dataLineList.get(i).getData().get("nps"));
            if(checkValue < minValue)
            {
                minValue = checkValue;
            }
        }

        return minValue;
    }

    //define min major gridline of second column bar chart's axis(absolute value)
    private double getMinMajorGridline(double minValue)
    {
        double minMajorGridline = 0;
        int majorLineScale = 0;
        if(minValue<0 && minValue>-90.5)
        {

            majorLineScale = 20;
        }
        else if(minValue<-90.5 && minValue>-377)
        {
            majorLineScale = 50;
        }



        double koef = Math.ceil((Math.abs(minValue)/majorLineScale));

        int realGridline = majorLineScale * (int)koef;

        if((realGridline-Math.abs(minValue)) < 7)
        {
            minMajorGridline = (realGridline + majorLineScale);
        }
        else
        {
            minMajorGridline = realGridline;
        }

        /*System.out.println("koef " + koef);
        System.out.println("realGridline "+ realGridline);
        System.out.println("minValue/realGridline " + (realGridline-Math.abs(minValue)));
        System.out.println("minMajorGridline " + minMajorGridline);*/

        return minMajorGridline;
    }

}
