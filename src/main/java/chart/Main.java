package chart;

import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import java.io.*;
import java.util.*;

public class Main {

    public static void main(String[] args) throws Exception {

        String filePptx = args[0];
        String filenameExcel = args[1];
        String pathForNewPptx = args[2];
        String npsTarget = args[3];
        Integer statusRound = Integer.valueOf(args[4].trim());
        String jsonBar = args[5].replace("\"", "'").replace(" ", "@");
        String jsonRadar = args[6].replace("\"", "'").replace(" ", "@");



        /*String filePptx = "d:/TestBot/template_new_diagram_3.pptx";
        String filenameExcel = "d:/TestBot/ChartPointss.xlsx";
        String pathForNewPptx = "d:/TestBot/ChartPointss.pptx";
       // String jsonBar = "[{'_Response_Rate':'0.4','_Detractors':'0.30','_Promoters':'0.20','_CSAT_round':'1H@2017','_Passives':'0.5','_NPS':'-30','_presentation_id':'1522739203821'},{'_Response_Rate':'0.10','_Detractors':'0.2','_Promoters':'0.7','_CSAT_round':'2H@2017','_Passives':'0.1','_NPS':'25.076923076923077','_presentation_id':'1522739203821'},{'_Response_Rate':'0.4','_Detractors':'0.30','_Promoters':'0.20','_CSAT_round':'1H@2017','_Passives':'0.5','_NPS':'-30','_presentation_id':'1522739203821'}]";
        String jsonBar = "[{'_Response_Rate':'0.4','_Detractors':'0.30','_Promoters':'0.20','_CSAT_round':'1H@2017','_Passives':'0.5','_NPS':'-30','_presentation_id':'1522739203821'},{'_Response_Rate':'0.10','_Detractors':'0.2','_Promoters':'0.7','_CSAT_round':'2H@2017','_Passives':'0.1','_NPS':'25.076923076923077','_presentation_id':'1522739203821'}]";
        String jsonRadar = "[{'_Communication':'3.3333333333333335','_Innovation':'1.2857142857142856','_Agreed_upon_Timeline':'2.888888888888889','_CSAT_round':'1h@2017','_Value':'4.2222222222222223','_Process':'2.875','_Comprehensive_Capabilities':'3.875','_Adaptability':'4.2222222222222223','_Technical_Excellence':'1.7777777777777777','_presentation_id':'1522739203821','_Appropriate_Team':'4.6666666666666665'},{'_Communication':'2.0833333333333335','_Innovation':'3.75','_Agreed_upon_Timeline':'1.25','_CSAT_round':'2h@2017','_Value':'3.5','_Process':'2.5','_Comprehensive_Capabilities':'2.8333333333333335','_Adaptability':'3.4166666666666665','_Technical_Excellence':'3.9166666666666665','_presentation_id':'1522739203821','_Appropriate_Team':'3.25'}, {'_Communication':'3.0833333333333335','_Innovation':'2.75','_Agreed_upon_Timeline':'3.25','_CSAT_round':'2h@2017','_Value':'3.5','_Process':'2.5','_Comprehensive_Capabilities':'2.8333333333333335','_Adaptability':'3.4166666666666665','_Technical_Excellence':'3.9166666666666665','_presentation_id':'1522739203821','_Appropriate_Team':'3.25'}]";
        String npsTarget = "70";
        int statusRound = 3;*/
        // System.out.println("Bar--->" + jsonBar);
        // System.out.println("Radar--->" + jsonRadar);


        //create slideshow and get list of slides
        XMLSlideShow slideShow = new XMLSlideShow(new FileInputStream(filePptx));
        List<XSLFSlide> slideList = slideShow.getSlides();

        //parse json for radar and bar charts, so we will get list of parameters for Bar and list of parameters for lines
        ParseJson parseJson = new ParseJson();

        //create massive in order to set the order of parameters for chart tree.
        //We use space in start name of parameters in order to compare with json(json has names of parameters with space after correction)
        String[] parametersBar = {" Promoters", " Passives", " Detractors"};
        String[] parametersLine = {" NPS", " Response Rate"};
        String[] parametersRadar = {" Technical Excellence", " Agreed-upon Timeline", " Process", " Communication", " Value", " Comprehensive Capabilities", " Innovation", " Adaptability", " Appropriate Team"};

        List<DataChart> dataRadar = parseJson.parseJsonRadar(jsonRadar, parametersRadar);
        Map<String, String[]> parametersBarColumn = new HashMap<>();
        parametersBarColumn.put("bar", parametersBar);
        parametersBarColumn.put("line", parametersLine);
        Map<String, List<DataChart>> dataBar = parseJson.parseJsonBar(jsonBar, parametersBarColumn);
        List<DataChart> dataBarList = dataBar.get("bar");
        List<DataChart> dataLineList = dataBar.get("line");

        //to movie 'Trend Line'
        ChartLineNPS npsGoal = new ChartLineNPS();
        npsGoal.setLineGoal(slideList.get(1), npsTarget, dataLineList);

        //Create Bar chart
        XSLFSlide slideTwo = slideList.get(1);
        XSLFShape chartShapeBar = null;
        List<XSLFShape> shapes = slideTwo.getShapes();

        //Code for ungroup chart
        for (XSLFShape shape : shapes) {
            //System.out.println(shape.getShapeName());
            if (shape.getShapeName().contains("Chart")) {
                chartShapeBar = shape;
            }
        }

        CreateBar createBar = new CreateBar(slideTwo, chartShapeBar);
        MyXSLFChart myXSLFChartBar = createBar.createChart();
        MyXSLFChartShape myXSLFChartBarShape = createBar.createChartShape(myXSLFChartBar);
        createBar.drawBar(myXSLFChartBarShape, dataBarList, dataLineList);

        //create Radar chart
        XSLFSlide slideFive = slideList.get(4);
        XSLFShape chartShapeRadar = null;
        List<XSLFShape> shapesRadar = slideFive.getShapes();
        for (XSLFShape sh : shapesRadar) {
            //System.out.println(sh.getShapeName());
            if (sh.getShapeName().contains("Chart")) {

                chartShapeRadar = sh;
            }
        }

        CreateRadar createRadar = new CreateRadar(slideFive, chartShapeRadar);
        MyXSLFChart myXSLFChartRadar = createRadar.createChart();
        MyXSLFChartShape myXSLFChartRadarShape = createRadar.createChartShape(myXSLFChartRadar);
        createRadar.drawRadar(myXSLFChartRadarShape, dataRadar);


        //to copy Chart points from excel file
        InputStream fs = new FileInputStream(filenameExcel);
        XSSFWorkbook book = new XSSFWorkbook(fs);
        XSSFSheet sheetOneH = book.getSheet("2x2 1R");
        XSSFSheet sheetTwoH = book.getSheet("2x2 2R");
        XSSFSheet sheetThreeH = book.getSheet("2x2 3R");
        //XSSFSheet sheetFourH = book.getSheet("2x2 4R");

        GrabChart grabChart = new GrabChart();

        switch(statusRound)
        {
            case 2:
                if(sheetTwoH != null)
                {
                    CTChart ctChartSecondRound = grabChart.getChart(sheetTwoH);
                    XSLFSlide slidePointsSecondRound = slideList.get(5);
                    grabChart.setChart(slidePointsSecondRound, ctChartSecondRound);
                }
                if (sheetThreeH != null) {
                    CTChart ctChartThreedRound = grabChart.getChart(sheetThreeH);
                    XSLFSlide slidePointsThreeRound = slideList.get(6);
                    grabChart.setChart(slidePointsThreeRound, ctChartThreedRound);
                }
                break;
            case 3:
                if(sheetOneH != null)
                {
                    CTChart ctChartFirstRound = grabChart.getChart(sheetOneH);
                    XSLFSlide slidePointsFirstRound = slideList.get(5);
                    grabChart.setChart(slidePointsFirstRound, ctChartFirstRound);
                }
                if(sheetTwoH != null)
                {
                    CTChart ctChartSecondRound = grabChart.getChart(sheetTwoH);
                    XSLFSlide slidePointsSecondRound = slideList.get(6);
                    grabChart.setChart(slidePointsSecondRound, ctChartSecondRound);
                }
                if (sheetThreeH != null) {
                    CTChart ctChartThreedRound = grabChart.getChart(sheetThreeH);
                    XSLFSlide slidePointsThreeRound = slideList.get(7);
                    grabChart.setChart(slidePointsThreeRound, ctChartThreedRound);
                }
                break;

        }

       /* if (sheetFourH != null) {
            CTChart ctChartFourRound = grabChart.getChart(sheetThreeH);
            XSLFSlide slidePointsFourRound = slideList.get(7);
            grabChart.setChart(slidePointsFourRound, ctChartFourRound);
        }*/

      try (FileOutputStream out = new FileOutputStream(pathForNewPptx)) {

            slideShow.write(out);
      }

    }
}
