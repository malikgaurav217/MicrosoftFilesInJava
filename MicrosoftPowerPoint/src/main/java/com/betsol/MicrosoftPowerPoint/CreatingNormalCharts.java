package com.betsol.MicrosoftPowerPoint;

import java.awt.Color;

import com.aspose.slides.ChartType;
import com.aspose.slides.FillType;
import com.aspose.slides.IChart;
import com.aspose.slides.IChartDataWorkbook;
import com.aspose.slides.IChartSeries;
import com.aspose.slides.IDataLabel;
import com.aspose.slides.ISlide;
import com.aspose.slides.NullableBool;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;
import com.betsol.MicrosoftPowerPoint.Utils;

public class CreatingNormalCharts {

	public static Presentation createNornalCharts(Presentation pres) {
		// Access first slide
		ISlide sld = pres.getSlides().get_Item(0);
		sld = pres.getSlides().addEmptySlide(sld.getLayoutSlide());
		
		// Add chart with default data
		IChart chart = sld.getShapes().addChart(ChartType.ClusteredColumn, 0, 0, 500, 500);

		// Setting chart Title
		// chart.ChartTitle.TextFrameForOverriding.Text = "Sample Title";
		chart.getChartTitle().addTextFrameForOverriding("Betsol");
		chart.getChartTitle().getTextFrameForOverriding().getTextFrameFormat().setCenterText(NullableBool.True);
		chart.getChartTitle().setHeight(20);
		chart.hasTitle();

		// Set first series to Show Values
		chart.getChartData().getSeries().get_Item(0).getLabels().getDefaultDataLabelFormat().setShowValue(true);

		// Setting the index of chart data sheet
		int defaultWorksheetIndex = 0;

		// Getting the chart data WorkSheet
		IChartDataWorkbook fact = chart.getChartData().getChartDataWorkbook();

		// Delete default generated series and categories
		chart.getChartData().getSeries().clear();
		chart.getChartData().getCategories().clear();
		int s = chart.getChartData().getSeries().size();
		s = chart.getChartData().getCategories().size();

		// Adding new series
		chart.getChartData().getSeries().add(fact.getCell(defaultWorksheetIndex, 0, 1, "Denver"),chart.getType());
		chart.getChartData().getSeries().add(fact.getCell(defaultWorksheetIndex, 0, 2, "Lone Tree"),chart.getType());

		// Adding new categories
		chart.getChartData().getCategories().add(fact.getCell(defaultWorksheetIndex, 1, 0, "Team 1"));
		chart.getChartData().getCategories().add(fact.getCell(defaultWorksheetIndex, 2, 0, "Team 2"));
		chart.getChartData().getCategories().add(fact.getCell(defaultWorksheetIndex, 3, 0, "Team 3"));

		// Take first chart series
		IChartSeries series = chart.getChartData().getSeries().get_Item(0);

		// Now populating series data

		series.getDataPoints().addDataPointForBarSeries(fact.getCell(defaultWorksheetIndex, 1, 1, 20));
		series.getDataPoints().addDataPointForBarSeries(fact.getCell(defaultWorksheetIndex, 2, 1, 50));
		series.getDataPoints().addDataPointForBarSeries(fact.getCell(defaultWorksheetIndex, 3, 1, 30));

		// Setting fill color for series
		series.getFormat().getFill().setFillType(FillType.Solid);
		series.getFormat().getFill().getSolidFillColor().setColor(Color.BLUE);

		// Take second chart series
		series = chart.getChartData().getSeries().get_Item(1);

		// Now populating series data
		series.getDataPoints().addDataPointForBarSeries(fact.getCell(defaultWorksheetIndex, 1, 2, 30));
		series.getDataPoints().addDataPointForBarSeries(fact.getCell(defaultWorksheetIndex, 2, 2, 10));
		series.getDataPoints().addDataPointForBarSeries(fact.getCell(defaultWorksheetIndex, 3, 2, 60));

		// Setting fill color for series
		series.getFormat().getFill().setFillType(FillType.Solid);
		series.getFormat().getFill().getSolidFillColor().setColor(Color.LIGHT_GRAY);

		// create custom labels for each of categories for new series
		// first label will be show Category name
		IDataLabel lbl = series.getDataPoints().get_Item(0).getLabel();
		lbl.getDataLabelFormat().setShowCategoryName(true);

		lbl = series.getDataPoints().get_Item(1).getLabel();
		lbl.getDataLabelFormat().setShowSeriesName(true);

		// Show value for third label
		lbl = series.getDataPoints().get_Item(2).getLabel();
		lbl.getDataLabelFormat().setShowValue(true);
		lbl.getDataLabelFormat().setShowSeriesName(true);
		lbl.getDataLabelFormat().setSeparator("/");

		System.out.println("AsposeChart.pptx" + " created");
		return pres;
	}

}
