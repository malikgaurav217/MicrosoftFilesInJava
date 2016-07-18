package com.betsol.MicrosoftPowerPoint;


import java.awt.Color;

import com.aspose.slides.FillType;
import com.aspose.slides.ISlide;
import com.aspose.slides.ITable;
import com.aspose.slides.Presentation;

public class CreateATableFromScratch {

	public static Presentation create(Presentation pres) {
		
		// Access first slide
		ISlide slide = pres.getSlides().get_Item(0);
		
		slide = pres.getSlides().addEmptySlide(slide.getLayoutSlide());
		
		// Define columns with widths and rows with heights
		double[] dblCols = { 150, 150, 150,150  };
		double[] dblRows = { 30, 30, 30, 30, 30,30, 30 };

		// Add table shape to slide
		ITable tbl = slide.getShapes().addTable(100, 50, dblCols, dblRows);
	
		// Set border format for each cell
		for (int row = 0; row < tbl.getRows().size(); row++) {
			for (int cell = 0; cell < tbl.getRows().get_Item(row).size(); cell++) {
				tbl.getRows().get_Item(row).get_Item(cell).getBorderTop().getFillFormat().setFillType(FillType.Solid);
				tbl.getRows().get_Item(row).get_Item(cell).getBorderTop().getFillFormat().getSolidFillColor()
						.setColor(Color.BLACK);
				tbl.getRows().get_Item(row).get_Item(cell).getBorderTop().setWidth(1);

				tbl.getRows().get_Item(row).get_Item(cell).getBorderBottom().getFillFormat()
						.setFillType(FillType.Solid);
				tbl.getRows().get_Item(row).get_Item(cell).getBorderBottom().getFillFormat().getSolidFillColor()
						.setColor(Color.BLACK);
				tbl.getRows().get_Item(row).get_Item(cell).getBorderBottom().setWidth(1);

				tbl.getRows().get_Item(row).get_Item(cell).getBorderLeft().getFillFormat().setFillType(FillType.Solid);
				tbl.getRows().get_Item(row).get_Item(cell).getBorderLeft().getFillFormat().getSolidFillColor()
						.setColor(Color.BLACK);
				tbl.getRows().get_Item(row).get_Item(cell).getBorderLeft().setWidth(1);

				tbl.getRows().get_Item(row).get_Item(cell).getBorderRight().getFillFormat().setFillType(FillType.Solid);
				tbl.getRows().get_Item(row).get_Item(cell).getBorderRight().getFillFormat().getSolidFillColor()
						.setColor(Color.BLACK);
				tbl.getRows().get_Item(row).get_Item(cell).getBorderRight().setWidth(1);
				
				tbl.getRows().get_Item(row).get_Item(cell).getTextFrame().setText(
						"row" + Integer.toString(row)+
						" cell" + Integer.toString(cell));
			}
		}
		// Merge cells 1 & 2 of row 1
		//tbl.mergeCells(tbl.getRows().get_Item(0).get_Item(0), tbl.getRows().get_Item(1).get_Item(0), false);

		// Add text to the merged cell
		//tbl.getRows().get_Item(0).get_Item(0).getTextFrame().setText("Merged Cells");

		System.out.println("created table");
		
		return pres;
	}

}
