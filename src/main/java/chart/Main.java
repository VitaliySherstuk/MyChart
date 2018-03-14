package chart;


import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {

    public static void main(String[] args) throws Exception {

       String filePptx = args[0];
       String filenameExcel = args[1];
       String pathForNewPptx = args[2];
       String jsonBar = args[3];
       String jsonRadar = args[4];


        //create slideshow and get list of slides
        XMLSlideShow slideShow = new XMLSlideShow(new FileInputStream(filePptx));
        List<XSLFSlide> slideList = slideShow.getSlides();

        //parse json for radar and bar charts, so we will get list of parameters for Bar and list of parameters for lines
        ParseJson parseJson = new ParseJson();
        String[] parametersBar = {" Detractors", " Passives", " Promoters"};
        String[] parametersLine = {" NPS", " Response Rate"};
        String[] parametersRadar = {" Technical Excellence", " Agreed-upon Timeline", " Process", " Communication", " Value", " Comprehensive Capabilities", " Innovation", " Adaptability", " Appropriate Team"};
        List<DataChart> dataRadar = parseJson.parseJsonRadar(jsonRadar, parametersRadar);
        Map<String, String[]> parametersBarColumn = new HashMap<>();
        parametersBarColumn.put("bar", parametersBar);
        parametersBarColumn.put("line", parametersLine);
        Map<String, List<DataChart>> dataBar = parseJson.parseJsonBar(jsonBar, parametersBarColumn);
        List<DataChart> dataBarList = dataBar.get("bar");
        List<DataChart> dataLineList = dataBar.get("line");


            //Create Bar chart
            XSLFSlide slideTwo = slideList.get(1);
            XSLFShape chartShapeBar = null;
            List<XSLFShape> shapes = slideTwo.getShapes();
            for (XSLFShape shape : shapes) {

                if (shape.getShapeName().contains("Chart 8")) {
                    chartShapeBar = shape;
                }
            }

            /* List<DataChart> dataBarList = createBar.getDataChart(new XSSFWorkbook(jsonBar), "bar");
            List<DataChart> dataLineList = createBar.getDataChart(new XSSFWorkbook(jsonBar), "line");*/

            CreateBar createBar = new CreateBar(slideTwo, chartShapeBar);
            MyXSLFChart myXSLFChartBar = createBar.createChart();
            MyXSLFChartShape myXSLFChartBarShape = createBar.createChartShape(myXSLFChartBar);
            createBar.drawBar(myXSLFChartBarShape, dataBarList, dataLineList);


            //create Radar chart
            XSLFSlide slideFive = slideList.get(4);
            XSLFShape chartShapeRadar = null;
            List<XSLFShape> shapesRadar = slideFive.getShapes();
            for (XSLFShape sh : shapesRadar) {
                if (sh.getShapeName().contains("Chart")) {

                    chartShapeRadar = sh;
                }
            }

            /*List<DataChart> dataRadar = createRadar.getDataRadar(new XSSFWorkbook(jsonRadar));*/

            CreateRadar createRadar = new CreateRadar(slideFive, chartShapeRadar);
            MyXSLFChart myXSLFChartRadar = createRadar.createChart();
            MyXSLFChartShape myXSLFChartRadarShape = createRadar.createChartShape(myXSLFChartRadar);
            createRadar.drawRadar(myXSLFChartRadarShape, dataRadar);


            //Chart points
            InputStream fs = new FileInputStream(filenameExcel);
            XSSFWorkbook book = new XSSFWorkbook(fs);
            XSSFSheet sheetOneH = book.getSheet("2x2 1H");
            XSSFSheet sheetTwoH = book.getSheet("2x2 2H");

            XSLFSlide slidePointsFirstRound = slideList.get(5);
            XSLFSlide slidePointsSecondRound = slideList.get(6);

            GrabChart grabChart = new GrabChart();
            CTChart ctChartFirstRound = grabChart.getChart(sheetOneH);
            CTChart ctChartSecondRound = grabChart.getChart(sheetTwoH);

            grabChart.setChart(slidePointsFirstRound, ctChartFirstRound);
            grabChart.setChart(slidePointsSecondRound, ctChartSecondRound);


            ChartLineNPS npsGoal = new ChartLineNPS();
            npsGoal.setLineGoal(slideList.get(1));

            /*String nameNewPptx = filenameExcel.substring(filenameExcel.lastIndexOf("\\"), filenameExcel.lastIndexOf(".")) + ".pptx";
            String outFile = pathForNewPptx+ "\\" + nameNewPptx.substring(1);
        System.out.println(nameNewPptx);
            System.out.println(outFile);*/
           try (FileOutputStream out = new FileOutputStream(pathForNewPptx)) {

                slideShow.write(out);
            }

    }
}
