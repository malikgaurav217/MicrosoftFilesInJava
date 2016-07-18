package com.betsol.MicrosoftPowerPoint;

import java.awt.Color;
import java.util.Date;
import com.aspose.slides.*;
import com.betsol.MicrosoftPowerPoint.Utils;

public class ProjectService {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ProjectService.class);

		AddLicense.add();
		Presentation pres = new Presentation();
			
		ISlide sld = pres.getSlides().get_Item(0);
		sld.getBackground().getStyleColor().setColor(Color.ORANGE);
		
		
		// Add an AutoShape of Rectangle type
		IAutoShape ashp = sld.getShapes().addAutoShape(ShapeType.Rectangle, 300, 300, 300, 300);
		// Add TextFrame to the Rectangle
		ashp.addTextFrame(" ");
		// Accessing the text frame
		ITextFrame txtFrame = ashp.getTextFrame();
		// Create the Paragraph object for text frame
		IParagraph para = txtFrame.getParagraphs().get_Item(0);
		// Create Portion object for paragraph
		IPortion portion = para.getPortions().get_Item(0);
		// Set Text
		portion.setText("Betsol - Slides");
		portion.getPortionFormat().setFontHeight(20);
		
		pres.getHeaderFooterManager().setFooterText("Betsol");
		pres.getHeaderFooterManager().setFooterVisible(true);
		pres.getHeaderFooterManager().setVisibilityOnTitleSlide(true);
		pres.getHeaderFooterManager().setDateTimeText(new Date().toGMTString());
		pres.getHeaderFooterManager().setDateTimeVisible(true);
		pres.getHeaderFooterManager().setSlideNumberVisible(true);
		
		pres = CreatingNormalCharts.createNornalCharts(pres);
		pres = CreatingScatteredChart.create(pres);
		pres = CreateATableFromScratch.create(pres);

		
		Date date = new Date();

		// Save the presentation to disk
		pres.save(dataDir + "ProjectOutput-" + date.getTime() + ".pptx", com.aspose.slides.SaveFormat.Pptx);

		System.out.println("Process completed Successfully");
	}
}
