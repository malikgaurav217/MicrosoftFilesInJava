package com.betsol.MicrosoftProjectPlan;

import java.io.FileInputStream;
import java.io.IOException;

import com.aspose.tasks.License;


public class AddLicense {

	public static void main(String[] args) {
		add();
	}
	public static void  add(){
		FileInputStream fstream = null;
		try {
			// Create a stream object containing the license file
			fstream = new FileInputStream("C:\\work\\betsol\\kieser\\MicrosoftProjectPlan\\src\\main\\java\\com\\betsol\\MicrosoftProjectPlan\\Aspose.Tasks.lic");

			// Instantiate the License class
			com.aspose.tasks.License license = new License();

			// Set the license through the stream object
			license.setLicense(fstream);
		} catch (Exception ex) {
			// Printing the exception, if it occurs
			System.out.println(ex.toString());
		} finally {
			// Closing the stream finally
			if (fstream != null)
				try {
					fstream.close();
				} catch (IOException e) {
					System.out.println(e.toString());
				}
		}
		System.out.println("License added");

	}

}
